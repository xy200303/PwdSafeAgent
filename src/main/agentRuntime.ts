import { join } from "node:path";
import OpenAI from "openai";
import type {
  ChatCompletionAssistantMessageParam,
  ChatCompletionMessageParam,
  ChatCompletionMessageToolCall,
  ChatCompletionToolMessageParam
} from "openai/resources/chat/completions";
import type { AgentRuntimeKind, AppSettings, ChatSession, StreamItem } from "../shared/types";
import { compactText, sanitizeFileName } from "./agentTools";
import { hasDiagramArtifactIntent, hasSchemeArtifactIntent } from "./artifactIntent";
import { buildAgentChatTools, executeAgentToolCall } from "./agentToolRegistry";
import type { BundledPythonRuntime } from "./bundledRuntime";
import { loadPiAgentRuntime } from "./piAgentAdapter";

type OpenAIChatClient = Pick<OpenAI, "chat">;

export type MessageStreamItem = Extract<StreamItem, { kind: "message" }>;

export interface AgentRuntimeTurnInput {
  sessionId: string;
  userPrompt: string;
  controller: AbortController;
  onAssistantCreated: (assistantItem: MessageStreamItem) => void;
}

export interface AgentRuntime {
  kind: AgentRuntimeKind;
  runTurn(input: AgentRuntimeTurnInput): Promise<MessageStreamItem>;
}

export interface AgentRuntimeHost {
  rootDir: string;
  docsDir: string;
  inputDir: string;
  outputDir: string;
  loadSettings(): AppSettings;
  getSession(sessionId: string): ChatSession | undefined;
  buildMessages(session: ChatSession): ChatCompletionMessageParam[];
  createAssistantMessage(sessionId: string): MessageStreamItem;
  updateItem(sessionId: string, item: StreamItem): void;
  startToolCall(sessionId: string, toolName: string, summary: string): StreamItem;
  finishToolCall(sessionId: string, item: StreamItem, status: "success" | "failed", summary: string): void;
  addStage(sessionId: string, title: string, detail?: string): void;
  appendSessionMemory(sessionId: string, source: string, content: string): void;
  formatSessionMemory(sessionId: string): string;
  getSessionReadableFiles(sessionId: string): string[];
  getBundledPythonRuntime(): BundledPythonRuntime | undefined;
  createArtifact(sessionId: string, filePath: string): void;
  streamMockResponse(sessionId: string, assistantItem: MessageStreamItem): Promise<void>;
}

const MAX_AGENT_TOOL_ROUNDS = 3;
const TOOL_PREVIEW_CHARS = 4000;
let openAIClientFactory = (settings: AppSettings): OpenAIChatClient =>
  new OpenAI({
    apiKey: process.env.OPENAI_API_KEY,
    baseURL: settings.openai.baseUrl,
    timeout: settings.openai.requestTimeoutMs
  });

export function setOpenAIClientFactoryForTesting(factory?: (settings: AppSettings) => OpenAIChatClient): void {
  openAIClientFactory =
    factory ??
    ((settings: AppSettings): OpenAIChatClient =>
      new OpenAI({
        apiKey: process.env.OPENAI_API_KEY,
        baseURL: settings.openai.baseUrl,
        timeout: settings.openai.requestTimeoutMs
      }));
}

export function normalizeAgentRuntimeKind(value: string | undefined): AgentRuntimeKind {
  return value?.trim().toLowerCase() === "pi-agent" ? "pi-agent" : "openai-chat";
}

export function createAgentRuntime(kind: AgentRuntimeKind, host: AgentRuntimeHost): AgentRuntime {
  if (kind === "pi-agent") {
    return createPiAgentRuntime(host);
  }
  return createOpenAIChatRuntime(host);
}

function createOpenAIChatRuntime(host: AgentRuntimeHost): AgentRuntime {
  return {
    kind: "openai-chat",
    async runTurn(input) {
      const settings = host.loadSettings();
      const session = host.getSession(input.sessionId);
      if (!session) throw new Error("Session not found");

      if (!process.env.OPENAI_API_KEY) {
        const assistantItem = host.createAssistantMessage(input.sessionId);
        input.onAssistantCreated(assistantItem);
        await host.streamMockResponse(input.sessionId, assistantItem);
        return assistantItem;
      }

      const client = openAIClientFactory(settings);

      const messages = await runOpenAIToolPlanning(
        input.sessionId,
        client,
        settings,
        host.buildMessages(session),
        input.controller,
        input.userPrompt,
        host
      );
      messages.push({
        role: "system",
        content: "工具调用阶段已结束。请基于用户需求、模板上下文、附件内容和工具结果，输出最终中文 Markdown 回复。"
      });

      const assistantItem = host.createAssistantMessage(input.sessionId);
      input.onAssistantCreated(assistantItem);
      const stream = await createFinalResponseStream({
        sessionId: input.sessionId,
        client,
        settings,
        messages,
        controller: input.controller,
        host
      });

      for await (const part of stream) {
        const delta = part.choices[0]?.delta?.content;
        if (!delta) continue;
        assistantItem.content += delta;
        host.updateItem(input.sessionId, assistantItem);
      }

      return assistantItem;
    }
  };
}

interface ChatCompletionStreamChunk {
  choices: Array<{
    delta?: {
      content?: string | null;
    };
  }>;
}

type ChatCompletionStream = AsyncIterable<ChatCompletionStreamChunk>;

async function createFinalResponseStream({
  sessionId,
  client,
  settings,
  messages,
  controller,
  host
}: {
  sessionId: string;
  client: OpenAIChatClient;
  settings: AppSettings;
  messages: ChatCompletionMessageParam[];
  controller: AbortController;
  host: AgentRuntimeHost;
}): Promise<ChatCompletionStream> {
  try {
    return (await client.chat.completions.create(
      {
        model: settings.openai.chatModel,
        messages,
        stream: true,
        max_completion_tokens: settings.openai.maxOutputTokens
      },
      { signal: controller.signal }
    )) as unknown as ChatCompletionStream;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (!isFunctionArgumentsJsonError(message)) {
      throw error;
    }

    host.addStage(
      sessionId,
      "最终回复已兼容降级",
      "模型服务不接受工具调用历史，已将工具结果转换为普通上下文后重试"
    );
    return (await client.chat.completions.create(
      {
        model: settings.openai.chatModel,
        messages: flattenToolHistoryForFinalResponse(messages),
        stream: true,
        max_completion_tokens: settings.openai.maxOutputTokens
      },
      { signal: controller.signal }
    )) as unknown as ChatCompletionStream;
  }
}

function flattenToolHistoryForFinalResponse(messages: ChatCompletionMessageParam[]): ChatCompletionMessageParam[] {
  const flattened: ChatCompletionMessageParam[] = [];

  for (const message of messages) {
    if (message.role === "tool") {
      flattened.push({
        role: "system",
        content: `以下是已执行工具的结果，请作为上下文参考：\n\n${stringifyChatContent(message.content)}`
      });
      continue;
    }

    if (message.role === "assistant" && "tool_calls" in message && message.tool_calls?.length) {
      const content = stringifyChatContent(message.content);
      if (content.trim()) {
        flattened.push({ role: "assistant", content });
      }
      flattened.push({
        role: "system",
        content: formatToolCallHistory(message.tool_calls)
      });
      continue;
    }

    flattened.push(message);
  }

  return flattened;
}

function formatToolCallHistory(calls: ChatCompletionMessageToolCall[]): string {
  return [
    "模型在前一步曾请求以下工具调用，这些调用已由应用执行；请不要再次依赖 Chat Completions tool_call 历史格式：",
    ...calls.map((call, index) => {
      if (call.type !== "function") return `${index + 1}. ${call.type}`;
      return `${index + 1}. ${call.function.name}(${sanitizeToolPreview(call.function.arguments)})`;
    })
  ].join("\n");
}

function stringifyChatContent(content: unknown): string {
  if (content === null || content === undefined) return "";
  if (typeof content === "string") return content;
  try {
    return JSON.stringify(content);
  } catch {
    return String(content);
  }
}

function createPiAgentRuntime(host: AgentRuntimeHost): AgentRuntime {
  const fallbackRuntime = createOpenAIChatRuntime(host);

  return {
    kind: "pi-agent",
    async runTurn(input) {
      const settings = host.loadSettings();
      const toolItem = host.startToolCall(
        input.sessionId,
        "pi-agent.runtime",
        settings.agent.piAgentPackage
          ? `加载 Pi Agent 插件 ${settings.agent.piAgentPackage}`
          : "未配置 Pi Agent 插件，准备回退"
      );
      const adapter = await loadPiAgentRuntime(settings, host, fallbackRuntime);
      host.finishToolCall(input.sessionId, toolItem, adapter.status === "failed" ? "failed" : "success", adapter.summary);

      if (adapter.runtime) {
        return adapter.runtime.runTurn(input);
      }
      return fallbackRuntime.runTurn(input);
    }
  };
}

async function runOpenAIToolPlanning(
  sessionId: string,
  client: OpenAIChatClient,
  settings: AppSettings,
  messages: ChatCompletionMessageParam[],
  controller: AbortController,
  userPrompt: string,
  host: AgentRuntimeHost
): Promise<ChatCompletionMessageParam[]> {
  const includeArtifactTools = shouldExposeArtifactTools(userPrompt, host.formatSessionMemory(sessionId));
  const tools = buildAgentChatTools({ includeExecBash: settings.agent.execBashEnabled, includeArtifactTools });
  const toolMessages = [...messages];

  if (!includeArtifactTools) {
    toolMessages.push({
      role: "system",
      content:
        "当前处于项目资料收集阶段：最终交付工具 write_word、write_pdf、image_generate、send_file 暂不开放。请通过对话继续澄清信息，必要时调用 remember_project 更新项目档案；不要声称已经生成文件。"
    });
  } else {
    toolMessages.push({
      role: "system",
      content:
        "当前已进入方案生成/交付阶段：如需交付文件，先确保项目事实已沉淀，再调用 write_word 生成 Word；需要图示时再调用 image_generate；生成完成后可调用 send_file。"
    });
  }

  for (let round = 1; round <= MAX_AGENT_TOOL_ROUNDS; round += 1) {
    throwIfAborted(controller);
    const completion = await createPlanningCompletion({
      sessionId,
      client,
      settings,
      toolMessages,
      tools,
      controller,
      host
    });
    if (!completion) return toolMessages;

    const message = completion.choices[0]?.message;
    const calls = message?.tool_calls ?? [];
    if (!calls.length) {
      if (message?.content?.trim()) {
        toolMessages.push({ role: "assistant", content: message.content });
      }
      return toolMessages;
    }

    const normalizedCalls = normalizeToolCallsForChatHistory(calls);
    const assistantToolMessage: ChatCompletionAssistantMessageParam = {
      role: "assistant",
      content: message.content ?? null,
      tool_calls: normalizedCalls
    };
    toolMessages.push(assistantToolMessage);

    for (const call of normalizedCalls) {
      throwIfAborted(controller);
      const functionName = call.type === "function" ? call.function.name : call.type;
      const toolItem = host.startToolCall(sessionId, functionName, `执行 ${functionName}`);
      if (toolItem.kind === "tool" && call.type === "function") {
        toolItem.inputPreview = sanitizeToolPreview(call.function.arguments);
        host.updateItem(sessionId, toolItem);
      }
      try {
        const result = await executeAgentToolCall(call, {
          rootDir: host.rootDir,
          docsDir: host.docsDir,
          outputDir: host.outputDir,
          sessionTitle: sanitizeFileName(host.getSession(sessionId)?.title || "密码应用方案"),
          memory: host.formatSessionMemory(sessionId),
          settings,
          userPrompt,
          signal: controller.signal,
          allowedReadDirs: [host.docsDir, host.inputDir, host.outputDir],
          allowedReadFiles: host.getSessionReadableFiles(sessionId),
          execBashEnabled: settings.agent.execBashEnabled,
          bundledPythonRuntime: host.getBundledPythonRuntime()
        });
        if (toolItem.kind === "tool") {
          toolItem.outputPreview = compactText(result.content || result.summary, TOOL_PREVIEW_CHARS);
        }
        host.finishToolCall(sessionId, toolItem, "success", result.summary);

        if (result.content.trim()) {
          host.appendSessionMemory(sessionId, `工具结果：${result.toolName}`, result.content);
        }
        if (result.artifactPath) {
          host.createArtifact(sessionId, result.artifactPath);
        }

        const nextToolMessage: ChatCompletionToolMessageParam = {
          role: "tool",
          tool_call_id: call.id,
          content: compactText(result.content || result.summary, 16000)
        };
        toolMessages.push(nextToolMessage);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        if (toolItem.kind === "tool") {
          toolItem.errorPreview = sanitizeToolPreview(message);
        }
        host.finishToolCall(sessionId, toolItem, "failed", message);
        toolMessages.push({
          role: "tool",
          tool_call_id: call.id,
          content: `Tool failed: ${message}`
        });
      }
    }
  }

  host.addStage(sessionId, "工具规划轮次已达上限", `${MAX_AGENT_TOOL_ROUNDS} 轮`);
  return toolMessages;
}

function shouldExposeArtifactTools(userPrompt: string, memory: string): boolean {
  if (hasSchemeArtifactIntent(userPrompt) || hasDiagramArtifactIntent(userPrompt)) return true;
  return /生成就绪[：:]\s*是|ready_for_generation[：:=]\s*true/i.test(memory);
}

async function createPlanningCompletion({
  sessionId,
  client,
  settings,
  toolMessages,
  tools,
  controller,
  host
}: {
  sessionId: string;
  client: OpenAIChatClient;
  settings: AppSettings;
  toolMessages: ChatCompletionMessageParam[];
  tools: ReturnType<typeof buildAgentChatTools>;
  controller: AbortController;
  host: AgentRuntimeHost;
}) {
  try {
    return await client.chat.completions.create(
      {
        model: settings.openai.chatModel,
        messages: toolMessages,
        tools,
        tool_choice: "auto",
        parallel_tool_calls: false,
        max_completion_tokens: Math.min(settings.openai.maxOutputTokens, 2048)
      },
      { signal: controller.signal }
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (!isFunctionArgumentsJsonError(message)) {
      throw error;
    }

    host.addStage(sessionId, "工具规划已兼容跳过", "模型服务返回了非法工具参数，已改为直接回复");
    toolMessages.push({
      role: "system",
      content:
        "本轮工具规划被跳过：兼容模型服务报告 function.arguments 不是合法 JSON。请不要继续调用工具，直接基于已有对话内容回复用户；如需要文件或模板，请提示用户重试或补充信息。"
    });
    return undefined;
  }
}

function throwIfAborted(controller: AbortController): void {
  if (controller.signal.aborted) {
    throw new Error("用户已停止生成");
  }
}

function sanitizeToolPreview(value: string): string {
  if (!value.trim()) return "";
  return compactText(
    value
      .replace(/(api[_-]?key|authorization|token|secret|password|passwd|OPENAI_API_KEY|OPENAI_IMAGE_API_KEY)("\s*:\s*"|'\s*:\s*'|=)([^"',\s}]+)/gi, "$1$2[REDACTED]")
      .replace(/Bearer\s+[A-Za-z0-9._~+/=-]+/gi, "Bearer [REDACTED]")
      .replace(/\bsk-[A-Za-z0-9_-]{12,}\b/g, "sk-[REDACTED]"),
    TOOL_PREVIEW_CHARS
  );
}

function normalizeToolCallsForChatHistory(calls: ChatCompletionMessageToolCall[]): ChatCompletionMessageToolCall[] {
  return calls.map((call) => {
    if (call.type !== "function") return call;
    return {
      ...call,
      function: {
        ...call.function,
        arguments: normalizeFunctionArguments(call.function.arguments, call.function.name)
      }
    };
  });
}

function normalizeFunctionArguments(raw: string, functionName: string): string {
  const parsed = parseJsonObject(raw);
  if (parsed) return JSON.stringify(parsed);

  const stripped = stripJsonCodeFence(raw);
  const parsedStripped = stripped === raw ? undefined : parseJsonObject(stripped);
  if (parsedStripped) return JSON.stringify(parsedStripped);

  const repaired = repairLooseJsonObject(stripped);
  const parsedRepaired = repaired ? parseJsonObject(repaired) : undefined;
  if (parsedRepaired) return JSON.stringify(parsedRepaired);

  return JSON.stringify(inferArgumentsFromPlainText(functionName, stripped));
}

function parseJsonObject(raw: string): Record<string, unknown> | undefined {
  if (!raw.trim()) return {};
  try {
    const parsed = JSON.parse(raw) as unknown;
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? (parsed as Record<string, unknown>) : undefined;
  } catch {
    return undefined;
  }
}

function stripJsonCodeFence(raw: string): string {
  return raw
    .trim()
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/i, "")
    .trim();
}

function repairLooseJsonObject(raw: string): string | undefined {
  const trimmed = raw.trim();
  if (!trimmed || trimmed.startsWith("{") || !trimmed.includes(":")) return undefined;
  return `{${trimmed}}`;
}

function inferArgumentsFromPlainText(functionName: string, raw: string): Record<string, unknown> {
  const value = raw.trim();
  if (!value || functionName === "time") return {};
  if (functionName === "web_search") return { query: value };
  if (["read_file", "read_word", "read_pdf", "send_file", "write_pdf"].includes(functionName)) return { path: value };
  if (functionName === "exec_bash") return { command: value };
  if (functionName === "image_generate") return { kind: "architecture", prompt: value };
  if (functionName === "remember_project") return { summary: value, facts: [], gaps: [], ready_for_generation: false };
  if (["write_file", "write_word"].includes(functionName)) return { content: value };
  return {};
}

function isFunctionArgumentsJsonError(message: string): boolean {
  return /function\.arguments/i.test(message) && /JSON/i.test(message);
}
