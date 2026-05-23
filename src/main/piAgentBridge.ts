import { mkdirSync } from "node:fs";
import { resolve } from "node:path";
import type { ChatCompletionMessageParam, ChatCompletionMessageToolCall } from "openai/resources/chat/completions";
import {
  AuthStorage,
  DefaultResourceLoader,
  ModelRegistry,
  SessionManager,
  createAgentSession,
  defineTool,
  type AgentSession,
  type AgentSessionEvent,
  type ToolDefinition
} from "@mariozechner/pi-coding-agent";
import type { Api, AssistantMessage, Model, TextContent } from "@mariozechner/pi-ai";
import { compactText } from "./agentTools";
import { buildAgentChatTools, executeAgentToolCall, type AgentToolExecutionResult } from "./agentToolRegistry";
import type { AgentRuntime, AgentRuntimeHost, AgentRuntimeTurnInput, MessageStreamItem } from "./agentRuntime";
import type { AppSettings, StreamItem } from "../shared/types";

const PWD_SAFE_PROVIDER = "pwdsafe-openai";
const ZERO_USAGE = {
  input: 0,
  output: 0,
  cacheRead: 0,
  cacheWrite: 0,
  totalTokens: 0,
  cost: {
    input: 0,
    output: 0,
    cacheRead: 0,
    cacheWrite: 0,
    total: 0
  }
};

interface PiSessionState {
  session: AgentSession;
  loader: DefaultResourceLoader;
  configSignature: string;
  toolNames: string[];
  systemPrompt: string;
  activeUserPrompt: string;
  activeSignal?: AbortSignal;
  model: Model<Api>;
  unsubscribe?: () => void;
}

export function createPiAgentBridge(host: AgentRuntimeHost): AgentRuntime {
  const sessionStates = new Map<string, PiSessionState>();

  return {
    async runTurn(input) {
      const appSession = host.getSession(input.sessionId);
      if (!appSession) {
        throw new Error(`Session not found: ${input.sessionId}`);
      }

      const currentUserPrompt = getLatestUserPrompt(host, input.sessionId) || input.userPrompt;
      const state = await ensurePiSessionState(host, input, sessionStates, currentUserPrompt);
      const assistantItem = await runPiPrompt(host, state, input, currentUserPrompt);
      return assistantItem;
    }
  };
}

async function ensurePiSessionState(
  host: AgentRuntimeHost,
  input: AgentRuntimeTurnInput,
  sessionStates: Map<string, PiSessionState>,
  currentUserPrompt: string
): Promise<PiSessionState> {
  const settings = host.loadSettings();
  const signature = buildConfigSignature(settings);
  const existing = sessionStates.get(input.sessionId);
  if (existing?.configSignature === signature) {
    existing.systemPrompt = buildPiSystemPrompt(host, input.sessionId);
    existing.activeUserPrompt = currentUserPrompt;
    existing.activeSignal = input.controller.signal;
    await refreshPiSession(existing);
    return existing;
  }

  existing?.unsubscribe?.();
  existing?.session.dispose();

  const state = await createPiSessionState(host, input, settings, signature, currentUserPrompt);
  sessionStates.set(input.sessionId, state);
  return state;
}

async function createPiSessionState(
  host: AgentRuntimeHost,
  input: AgentRuntimeTurnInput,
  settings: AppSettings,
  configSignature: string,
  currentUserPrompt: string
): Promise<PiSessionState> {
  const agentDir = resolve(host.outputDir, "..", "pi-agent");
  mkdirSync(agentDir, { recursive: true });

  const authStorage = AuthStorage.inMemory();
  const modelRegistry = ModelRegistry.inMemory(authStorage);
  const model = registerPwdSafeOpenAiModel(modelRegistry, settings);
  const toolDefinitions = createPwdSafePiTools(host, input.sessionId, settings);
  const toolNames = toolDefinitions.map((tool) => tool.name);
  const systemPrompt = buildPiSystemPrompt(host, input.sessionId);
  const stateRef: Pick<PiSessionState, "systemPrompt" | "activeUserPrompt" | "activeSignal"> = {
    systemPrompt,
    activeUserPrompt: currentUserPrompt,
    activeSignal: input.controller.signal
  };
  const loader = new DefaultResourceLoader({
    cwd: host.rootDir,
    agentDir,
    noExtensions: true,
    noSkills: true,
    noPromptTemplates: true,
    noThemes: true,
    noContextFiles: true,
    systemPromptOverride: () => stateRef.systemPrompt,
    appendSystemPromptOverride: () => []
  });
  await loader.reload();

  const { session } = await createAgentSession({
    cwd: host.rootDir,
    agentDir,
    authStorage,
    modelRegistry,
    model,
    thinkingLevel: resolveThinkingLevel(settings),
    customTools: toolDefinitions,
    tools: toolNames,
    noTools: "builtin",
    resourceLoader: loader,
    sessionManager: SessionManager.inMemory(host.rootDir)
  });
  session.agent.state.messages = buildInitialPiHistory(host, input.sessionId, model);

  const state: PiSessionState = {
    session,
    loader,
    configSignature,
    toolNames,
    systemPrompt,
    activeUserPrompt: currentUserPrompt,
    activeSignal: input.controller.signal,
    model
  };

  Object.defineProperty(stateRef, "systemPrompt", {
    get: () => state.systemPrompt
  });
  Object.defineProperty(stateRef, "activeUserPrompt", {
    get: () => state.activeUserPrompt
  });
  Object.defineProperty(stateRef, "activeSignal", {
    get: () => state.activeSignal
  });

  await refreshPiSession(state);
  return state;
}

async function refreshPiSession(state: PiSessionState): Promise<void> {
  await state.loader.reload();
  await state.session.reload();
  state.session.setActiveToolsByName(state.toolNames);
}

async function runPiPrompt(
  host: AgentRuntimeHost,
  state: PiSessionState,
  input: AgentRuntimeTurnInput,
  prompt: string
): Promise<MessageStreamItem> {
  let assistantItem: MessageStreamItem | undefined;
  const toolItems = new Map<string, StreamItem>();
  const abort = () => {
    void state.session.abort();
  };

  input.controller.signal.addEventListener("abort", abort, { once: true });
  state.unsubscribe = state.session.subscribe((event) => {
    handlePiSessionEvent(event, {
      host,
      input,
      state,
      toolItems,
      getAssistantItem: () => assistantItem,
      setAssistantItem: (item) => {
        assistantItem = item;
      }
    });
  });

  try {
    await state.session.prompt(prompt, {
      expandPromptTemplates: false,
      source: "interactive"
    });

    if (!assistantItem) {
      assistantItem = host.createAssistantMessage(input.sessionId);
      input.onAssistantCreated(assistantItem);
    }
    const finalText = state.session.getLastAssistantText();
    if (finalText && assistantItem.content !== finalText) {
      assistantItem.content = finalText;
      host.updateItem(input.sessionId, assistantItem);
    }
    return assistantItem;
  } finally {
    input.controller.signal.removeEventListener("abort", abort);
    state.unsubscribe?.();
    state.unsubscribe = undefined;
    state.activeSignal = undefined;
  }
}

function handlePiSessionEvent(
  event: AgentSessionEvent,
  context: {
    host: AgentRuntimeHost;
    input: AgentRuntimeTurnInput;
    state: PiSessionState;
    toolItems: Map<string, StreamItem>;
    getAssistantItem: () => MessageStreamItem | undefined;
    setAssistantItem: (item: MessageStreamItem) => void;
  }
): void {
  switch (event.type) {
    case "message_update": {
      const update = event.assistantMessageEvent;
      if (update.type !== "text_delta") return;
      const assistantItem = ensureAssistantItem(context);
      assistantItem.content += update.delta;
      context.host.updateItem(context.input.sessionId, assistantItem);
      return;
    }
    case "message_end": {
      if (event.message.role !== "assistant") return;
      const text = extractAssistantText(event.message);
      if (!text) return;
      const assistantItem = ensureAssistantItem(context);
      assistantItem.content = text;
      context.host.updateItem(context.input.sessionId, assistantItem);
      return;
    }
    case "tool_execution_start": {
      const item = context.host.startToolCall(
        context.input.sessionId,
        event.toolName,
        `Pi Agent 调用工具 ${event.toolName}`
      );
      if (item.kind === "tool") {
        item.toolCallId = event.toolCallId;
        item.inputPreview = compactText(JSON.stringify(event.args ?? {}, null, 2), 2000);
        context.host.updateItem(context.input.sessionId, item);
      }
      context.toolItems.set(event.toolCallId, item);
      return;
    }
    case "tool_execution_update": {
      const item = context.toolItems.get(event.toolCallId);
      if (item?.kind !== "tool") return;
      item.outputPreview = compactText(formatToolResultContent(event.partialResult), 4000);
      context.host.updateItem(context.input.sessionId, item);
      return;
    }
    case "tool_execution_end": {
      const item = context.toolItems.get(event.toolCallId);
      if (item?.kind !== "tool") return;
      const details = readToolExecutionDetails(event.result);
      const summary = details?.summary || formatToolResultContent(event.result) || `工具 ${event.toolName} 执行完成`;
      const output = details?.content || formatToolResultContent(event.result);
      if (event.isError) {
        item.errorPreview = compactText(output, 6000);
      } else {
        item.outputPreview = compactText(output, 6000);
      }
      context.host.finishToolCall(context.input.sessionId, item, event.isError ? "failed" : "success", summary);
      return;
    }
    default:
      return;
  }
}

function ensureAssistantItem(context: {
  host: AgentRuntimeHost;
  input: AgentRuntimeTurnInput;
  getAssistantItem: () => MessageStreamItem | undefined;
  setAssistantItem: (item: MessageStreamItem) => void;
}): MessageStreamItem {
  const existing = context.getAssistantItem();
  if (existing) return existing;

  const item = context.host.createAssistantMessage(context.input.sessionId);
  context.input.onAssistantCreated(item);
  context.setAssistantItem(item);
  return item;
}

function createPwdSafePiTools(
  host: AgentRuntimeHost,
  sessionId: string,
  settings: AppSettings
): ToolDefinition[] {
  return buildAgentChatTools({
    includeExecBash: settings.agent.execBashEnabled,
    includeArtifactTools: true
  })
    .filter((tool) => tool.type === "function")
    .map((tool) => {
      const definition = tool.function;
      return defineTool({
        name: definition.name,
        label: definition.name,
        description: definition.description || definition.name,
        promptSnippet: `${definition.name}: ${firstLine(definition.description || definition.name)}`,
        parameters: definition.parameters as ToolDefinition["parameters"],
        executionMode: "sequential",
        execute: async (toolCallId, params, signal) => {
          const activeSettings = host.loadSettings();
          const toolCall: ChatCompletionMessageToolCall = {
            id: toolCallId,
            type: "function",
            function: {
              name: definition.name,
              arguments: JSON.stringify(params ?? {})
            }
          };
          const result = await executeAgentToolCall(toolCall, {
            rootDir: host.rootDir,
            docsDir: host.docsDir,
            outputDir: host.outputDir,
            sessionTitle: host.getSession(sessionId)?.title || "密码应用方案",
            memory: host.formatSessionMemory(sessionId),
            settings: activeSettings,
            userPrompt: getLatestUserPrompt(host, sessionId),
            signal: signal ?? undefined,
            allowedReadDirs: [host.docsDir, host.inputDir, host.outputDir],
            allowedReadFiles: host.getSessionReadableFiles(sessionId),
            execBashEnabled: activeSettings.agent.execBashEnabled,
            bundledPythonRuntime: host.getBundledPythonRuntime()
          });

          host.appendSessionMemory(sessionId, `工具 ${result.toolName}`, result.content);
          if (result.artifactPath) {
            host.createArtifact(sessionId, result.artifactPath);
          }

          return {
            content: [{ type: "text", text: result.content } satisfies TextContent],
            details: result
          };
        }
      });
    });
}

function registerPwdSafeOpenAiModel(modelRegistry: ModelRegistry, settings: AppSettings): Model<Api> {
  const modelId = settings.openai.chatModel.trim() || "gpt-5.5";
  const baseUrl = settings.openai.baseUrl.trim() || "https://api.openai.com/v1";
  const maxTokens = Number.isFinite(settings.openai.maxOutputTokens) ? settings.openai.maxOutputTokens : 16000;
  const qwenCompatible = isQwenCompatibleModel(modelId, baseUrl);

  modelRegistry.registerProvider(PWD_SAFE_PROVIDER, {
    name: "PwdSafeAgent OpenAI Chat",
    baseUrl,
    apiKey: "OPENAI_API_KEY",
    api: "openai-completions",
    authHeader: false,
    models: [
      {
        id: modelId,
        name: modelId,
        api: "openai-completions",
        reasoning: settings.openai.thinkingEnabled,
        input: ["text", "image"],
        cost: {
          input: 0,
          output: 0,
          cacheRead: 0,
          cacheWrite: 0
        },
        contextWindow: 200000,
        maxTokens,
        compat: {
          supportsDeveloperRole: false,
          supportsReasoningEffort: settings.openai.thinkingEnabled && !qwenCompatible,
          supportsStrictMode: false,
          maxTokensField: "max_tokens",
          thinkingFormat: qwenCompatible ? "qwen" : "openai"
        }
      }
    ]
  });

  const model = modelRegistry.find(PWD_SAFE_PROVIDER, modelId);
  if (!model) {
    throw new Error(`Pi Agent 模型注册失败：${PWD_SAFE_PROVIDER}/${modelId}`);
  }
  return model;
}

function isQwenCompatibleModel(modelId: string, baseUrl: string): boolean {
  return /qwen|dashscope|aliyuncs/i.test(`${modelId} ${baseUrl}`);
}

function buildConfigSignature(settings: AppSettings): string {
  return JSON.stringify({
    baseUrl: settings.openai.baseUrl,
    chatModel: settings.openai.chatModel,
    thinkingEnabled: settings.openai.thinkingEnabled,
    reasoningEffort: settings.openai.reasoningEffort,
    maxOutputTokens: settings.openai.maxOutputTokens,
    execBashEnabled: settings.agent.execBashEnabled
  });
}

function resolveThinkingLevel(settings: AppSettings): "off" | "minimal" | "low" | "medium" | "high" | "xhigh" {
  if (!settings.openai.thinkingEnabled) return "off";
  const effort = settings.openai.reasoningEffort.trim();
  if (effort === "minimal" || effort === "low" || effort === "medium" || effort === "high" || effort === "xhigh") {
    return effort;
  }
  return "medium";
}

function buildPiSystemPrompt(host: AgentRuntimeHost, sessionId: string): string {
  const session = host.getSession(sessionId);
  if (!session) return "";
  const messages = host.buildMessages(session);
  const systemParts = messages
    .filter((message) => message.role === "system" || message.role === "developer")
    .map((message) => chatContentToText(message.content))
    .filter(Boolean);
  return systemParts.join("\n\n");
}

function buildInitialPiHistory(host: AgentRuntimeHost, sessionId: string, model: Model<Api>) {
  const session = host.getSession(sessionId);
  if (!session) return [];

  const messages = session.items.filter((item): item is MessageStreamItem => item.kind === "message");
  const lastUserIndex = findLastIndex(messages, (item) => item.role === "user");
  const history = lastUserIndex >= 0 ? messages.slice(0, lastUserIndex) : messages;

  return history
    .filter((item) => item.content.trim())
    .map((item) => {
      const timestamp = Date.parse(item.createdAt) || Date.now();
      if (item.role === "user") {
        return {
          role: "user" as const,
          content: item.content,
          timestamp
        };
      }
      return {
        role: "assistant" as const,
        content: [{ type: "text" as const, text: item.content }],
        api: model.api,
        provider: model.provider,
        model: model.id,
        usage: ZERO_USAGE,
        stopReason: "stop" as const,
        timestamp
      };
    });
}

function getLatestUserPrompt(host: AgentRuntimeHost, sessionId: string): string {
  const session = host.getSession(sessionId);
  const items = session?.items ?? [];
  for (let index = items.length - 1; index >= 0; index -= 1) {
    const item = items[index];
    if (item.kind === "message" && item.role === "user") {
      return item.content;
    }
  }
  return "";
}

function extractAssistantText(message: AssistantMessage): string {
  return message.content
    .filter((content): content is TextContent => content.type === "text")
    .map((content) => content.text)
    .join("");
}

function formatToolResultContent(result: unknown): string {
  if (!result || typeof result !== "object") return "";
  const record = result as { content?: Array<{ type: string; text?: string; mimeType?: string }> };
  if (!Array.isArray(record.content)) return "";
  return record.content
    .map((item) => {
      if (item.type === "text") return item.text || "";
      if (item.type === "image") return `[image:${item.mimeType || "unknown"}]`;
      return "";
    })
    .filter(Boolean)
    .join("\n");
}

function readToolExecutionDetails(result: unknown): AgentToolExecutionResult | undefined {
  if (!result || typeof result !== "object") return undefined;
  const details = (result as { details?: unknown }).details;
  if (!details || typeof details !== "object") return undefined;
  const record = details as Partial<AgentToolExecutionResult>;
  return typeof record.toolName === "string" && typeof record.summary === "string" && typeof record.content === "string"
    ? (record as AgentToolExecutionResult)
    : undefined;
}

function chatContentToText(content: ChatCompletionMessageParam["content"]): string {
  if (typeof content === "string") return content;
  if (!Array.isArray(content)) return "";
  return content
    .map((part) => {
      if (typeof part === "string") return part;
      if (part && typeof part === "object" && "text" in part && typeof part.text === "string") return part.text;
      return "";
    })
    .filter(Boolean)
    .join("\n");
}

function firstLine(value: string): string {
  return value.split(/\r?\n/, 1)[0]?.trim() || value;
}

function findLastIndex<T>(items: T[], predicate: (item: T) => boolean): number {
  for (let index = items.length - 1; index >= 0; index -= 1) {
    if (predicate(items[index])) return index;
  }
  return -1;
}
