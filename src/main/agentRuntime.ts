import { join } from "node:path";
import OpenAI from "openai";
import type {
  ChatCompletionAssistantMessageParam,
  ChatCompletionMessageParam,
  ChatCompletionToolMessageParam
} from "openai/resources/chat/completions";
import type { AgentRuntimeKind, AppSettings, ChatSession, StreamItem } from "../shared/types";
import { compactText, sanitizeFileName } from "./agentTools";
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
      const stream = await client.chat.completions.create(
        {
          model: settings.openai.chatModel,
          messages,
          stream: true,
          max_completion_tokens: settings.openai.maxOutputTokens
        },
        { signal: input.controller.signal }
      );

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
  const tools = buildAgentChatTools({ includeExecBash: settings.agent.execBashEnabled });
  const toolMessages = [...messages];

  for (let round = 1; round <= MAX_AGENT_TOOL_ROUNDS; round += 1) {
    throwIfAborted(controller);
    const plannerTool = host.startToolCall(sessionId, "openai.chat.tools", `工具规划第 ${round} 轮`);
    const completion = await client.chat.completions.create(
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

    const message = completion.choices[0]?.message;
    const calls = message?.tool_calls ?? [];
    if (!calls.length) {
      host.finishToolCall(sessionId, plannerTool, "success", "模型判断无需继续调用工具");
      if (message?.content?.trim()) {
        toolMessages.push({ role: "assistant", content: message.content });
      }
      return toolMessages;
    }

    host.finishToolCall(sessionId, plannerTool, "success", `模型请求 ${calls.length} 个工具调用`);
    const assistantToolMessage: ChatCompletionAssistantMessageParam = {
      role: "assistant",
      content: message.content ?? null,
      tool_calls: calls
    };
    toolMessages.push(assistantToolMessage);

    for (const call of calls) {
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
