import { mkdirSync } from "node:fs";
import { basename, resolve } from "node:path";
import type { ChatCompletionMessageParam, ChatCompletionMessageToolCall } from "openai/resources/chat/completions";
import { Type, type TSchema } from "typebox";
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
import { resolveAssistantDisplayTextOrThrow } from "./piAgentResult";
import type { AgentRuntime, AgentRuntimeHost, AgentRuntimeTurnInput, MessageStreamItem } from "./agentRuntime";
import { extractTemplateAnchorIds } from "./schemeProgress";
import type { AppSettings, SchemeProgressItem, StreamItem } from "../shared/types";

const PWD_SAFE_PROVIDER = "pwdsafe-openai";
const PI_TOOL_SCHEMA_VERSION = 2;
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

type JsonSchemaPrimitive = string | number | boolean | null;
type JsonSchemaValue = JsonSchemaPrimitive | JsonSchemaObject | JsonSchemaValue[];
type JsonSchemaTypeName = "object" | "array" | "string" | "integer" | "number" | "boolean" | "null";

interface JsonSchemaObject {
  type?: JsonSchemaTypeName | JsonSchemaTypeName[];
  properties?: Record<string, JsonSchemaValue>;
  required?: string[];
  items?: JsonSchemaValue | JsonSchemaValue[];
  additionalProperties?: boolean | JsonSchemaValue;
  enum?: JsonSchemaPrimitive[];
  anyOf?: JsonSchemaValue[];
  oneOf?: JsonSchemaValue[];
  allOf?: JsonSchemaValue[];
  description?: string;
  title?: string;
  default?: unknown;
  examples?: unknown;
  [key: string]: unknown;
}

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
  const agentDir = resolve(host.outputDir, "..", "pi-agent", input.sessionId);
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
  const toolArgs = new Map<string, unknown>();
  const sealedAssistantTexts: string[] = [];
  const assistantCompletion: {
    finalText?: string;
    stopReason?: AssistantMessage["stopReason"];
    errorMessage?: string;
  } = {};
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
      toolArgs,
      sealedAssistantTexts,
      assistantCompletion,
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

    const finalText = resolveAssistantDisplayTextOrThrow({
      finalText:
        assistantCompletion.finalText ??
        stripSealedAssistantText(state.session.getLastAssistantText(), sealedAssistantTexts),
      assistantContent: assistantItem?.content,
      toolItems,
      stopReason: assistantCompletion.stopReason,
      errorMessage: assistantCompletion.errorMessage
    });
    if (!assistantItem) {
      assistantItem = host.createAssistantMessage(input.sessionId);
      input.onAssistantCreated(assistantItem);
    }
    if (assistantItem.content !== finalText) {
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
    toolArgs: Map<string, unknown>;
    sealedAssistantTexts: string[];
    assistantCompletion: {
      finalText?: string;
      stopReason?: AssistantMessage["stopReason"];
      errorMessage?: string;
    };
    getAssistantItem: () => MessageStreamItem | undefined;
    setAssistantItem: (item: MessageStreamItem | undefined) => void;
  }
): void {
  switch (event.type) {
    case "message_update": {
      const update = event.assistantMessageEvent;
      if (update.type === "text_delta") {
        const assistantItem = ensureAssistantItem(context);
        assistantItem.content += update.delta;
        context.host.updateItem(context.input.sessionId, assistantItem);
        return;
      }
      if (update.type === "done") {
        captureAssistantCompletion(update.message, context);
        return;
      }
      if (update.type === "error") {
        captureAssistantCompletion(update.error, context);
        return;
      }
      return;
    }
    case "message_end": {
      if (event.message.role !== "assistant") return;
      captureAssistantCompletion(event.message, context);
      return;
    }
    case "tool_execution_start": {
      sealCurrentAssistantAsProcessNote(context);
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
      context.toolArgs.set(event.toolCallId, event.args);
      trackSchemeToolStart(event.toolName, event.args, context);
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
      trackSchemeToolEnd(event.toolName, context.toolArgs.get(event.toolCallId), details, event.isError, context);
      context.toolArgs.delete(event.toolCallId);
      return;
    }
    default:
      return;
  }
}

function trackSchemeToolStart(
  toolName: string,
  args: unknown,
  context: { host: AgentRuntimeHost; input: AgentRuntimeTurnInput }
): void {
  if (toolName === "draft_scheme_sections") {
    const sections = readDraftToolSections(args);
    if (sections.length) {
      for (const section of sections) {
        context.host.updateSchemeSectionProgress(context.input.sessionId, {
          section,
          status: "drafting",
          detail: `${section} 正在并行起草正文`
        });
      }
    } else {
      context.host.ensureSchemeProgress(context.input.sessionId, "正在并行起草章节正文");
    }
    return;
  }

  if (toolName === "create_word") {
    context.host.ensureSchemeProgress(
      context.input.sessionId,
      "已进入 Word 模板生成流程，等待批量写入章节内容",
      readToolStringArg(args, "name")
    );
    return;
  }

  if (toolName !== "write_word") return;
  const section = readSchemeToolSection(args);
  if (section) {
    context.host.updateSchemeSectionProgress(context.input.sessionId, {
      section,
      status: "running",
      detail: `正在生成 ${section}`
    });
    return;
  }

  context.host.ensureSchemeProgress(context.input.sessionId, "正在按模板写入 Word 内容");
}

function trackSchemeToolEnd(
  toolName: string,
  args: unknown,
  details: AgentToolExecutionResult | undefined,
  isError: boolean,
  context: { host: AgentRuntimeHost; input: AgentRuntimeTurnInput }
): void {
  if (toolName === "draft_scheme_sections") {
    if (isError) {
      for (const section of readDraftToolSections(args)) {
        context.host.updateSchemeSectionProgress(context.input.sessionId, {
          section,
          status: "failed",
          detail: `${section} 起草失败`
        });
      }
      return;
    }
    for (const update of details?.schemeProgressUpdates ?? []) {
      context.host.updateSchemeSectionProgress(context.input.sessionId, update);
    }
    return;
  }

  if (toolName === "create_word") {
    if (isError) {
      context.host.settleSchemeProgress(context.input.sessionId, "failed");
      return;
    }
    context.host.ensureSchemeProgress(
      context.input.sessionId,
      "Word 模板副本已创建，后续将按章节增量写入",
      details?.artifactPath ? basename(details.artifactPath) : readToolStringArg(args, "name")
    );
    return;
  }

  if (toolName !== "write_word") return;
  const section = readSchemeToolSection(args);
  const anchorIds = extractTemplateAnchorIds(details?.content ?? "");
  const artifactName = details?.artifactPath ? basename(details.artifactPath) : undefined;
  if (isError) {
    if (section || anchorIds.length) {
      context.host.updateSchemeSectionProgress(context.input.sessionId, {
        section,
        anchorIds,
        status: "failed",
        detail: section ? `${section} 生成失败` : "Word 写入失败",
        artifactName
      });
      return;
    }
    context.host.settleSchemeProgress(context.input.sessionId, "failed");
    return;
  }

  context.host.updateSchemeSectionProgress(context.input.sessionId, {
    section,
    anchorIds,
    status: "completed",
    detail: section ? `${section} 已写入模板` : "已按模板写入 Word",
    artifactName
  });
}

function readSchemeToolSection(args: unknown): string | undefined {
  return readToolStringArg(args, "section") || readToolStringArg(args, "section_title");
}

function readDraftToolSections(args: unknown): string[] {
  if (!args || typeof args !== "object" || Array.isArray(args)) return [];
  const value = (args as Record<string, unknown>).sections;
  if (!Array.isArray(value)) return [];
  return value
    .map((item) => {
      if (typeof item === "string") return item.trim();
      if (!item || typeof item !== "object" || Array.isArray(item)) return "";
      const section = (item as Record<string, unknown>).section;
      return typeof section === "string" ? section.trim() : "";
    })
    .filter(Boolean);
}

function readToolStringArg(args: unknown, key: string): string | undefined {
  if (!args || typeof args !== "object" || Array.isArray(args)) return undefined;
  const value = (args as Record<string, unknown>)[key];
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function ensureAssistantItem(context: {
  host: AgentRuntimeHost;
  input: AgentRuntimeTurnInput;
  getAssistantItem: () => MessageStreamItem | undefined;
  setAssistantItem: (item: MessageStreamItem | undefined) => void;
}): MessageStreamItem {
  const existing = context.getAssistantItem();
  if (existing) return existing;

  const item = context.host.createAssistantMessage(context.input.sessionId);
  context.input.onAssistantCreated(item);
  context.setAssistantItem(item);
  return item;
}

function sealCurrentAssistantAsProcessNote(context: {
  host: AgentRuntimeHost;
  input: AgentRuntimeTurnInput;
  sealedAssistantTexts: string[];
  getAssistantItem: () => MessageStreamItem | undefined;
  setAssistantItem: (item: MessageStreamItem | undefined) => void;
}): void {
  const assistantItem = context.getAssistantItem();
  if (!assistantItem?.content.trim()) return;
  assistantItem.isFinished = true;
  context.host.updateItem(context.input.sessionId, assistantItem);
  context.sealedAssistantTexts.push(assistantItem.content);
  context.setAssistantItem(undefined);
}

function stripSealedAssistantText(text: string | undefined, sealedTexts: string[]): string {
  let nextText = (text || "").trim();
  for (const sealedText of sealedTexts) {
    const prefix = sealedText.trim();
    if (prefix && nextText.startsWith(prefix)) {
      nextText = nextText.slice(prefix.length).trimStart();
    }
  }
  return nextText.trim();
}

function captureAssistantCompletion(
  message: AssistantMessage,
  context: {
    host: AgentRuntimeHost;
    input: AgentRuntimeTurnInput;
    sealedAssistantTexts: string[];
    assistantCompletion: {
      finalText?: string;
      stopReason?: AssistantMessage["stopReason"];
      errorMessage?: string;
    };
    getAssistantItem: () => MessageStreamItem | undefined;
    setAssistantItem: (item: MessageStreamItem | undefined) => void;
  }
): void {
  context.assistantCompletion.stopReason = message.stopReason;
  context.assistantCompletion.errorMessage = message.errorMessage;
  const text = stripSealedAssistantText(extractAssistantText(message), context.sealedAssistantTexts);
  context.assistantCompletion.finalText = text || context.assistantCompletion.finalText;
  if (!text) return;
  const assistantItem = ensureAssistantItem(context);
  assistantItem.content = text;
  context.host.updateItem(context.input.sessionId, assistantItem);
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
        parameters: convertJsonSchemaToTypeBoxSchema(definition.parameters),
        executionMode: resolvePwdSafeToolExecutionMode(definition.name),
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
            schemeProgress: getLatestSchemeProgress(host, sessionId),
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

export function convertJsonSchemaToTypeBoxSchema(schema: unknown, path = "parameters"): TSchema {
  if (schema === true) return Type.Unknown();
  if (schema === false) return Type.Never();
  const source = readJsonSchemaObject(schema, path);

  if (Array.isArray(source.enum)) {
    return Type.Unsafe(cloneJsonSchemaObject(source));
  }

  if (Array.isArray(source.allOf)) {
    return Type.Intersect(
      source.allOf.map((item, index) => convertJsonSchemaToTypeBoxSchema(item, `${path}.allOf[${index}]`)),
      readTypeBoxOptions(source, ["allOf"])
    );
  }

  if (Array.isArray(source.anyOf)) {
    return Type.Union(
      source.anyOf.map((item, index) => convertJsonSchemaToTypeBoxSchema(item, `${path}.anyOf[${index}]`)),
      readTypeBoxOptions(source, ["anyOf"])
    );
  }

  if (Array.isArray(source.oneOf)) {
    return Type.Union(
      source.oneOf.map((item, index) => convertJsonSchemaToTypeBoxSchema(item, `${path}.oneOf[${index}]`)),
      readTypeBoxOptions(source, ["oneOf"])
    );
  }

  const schemaTypes = readJsonSchemaTypes(source, path);
  if (schemaTypes.length > 1) {
    return Type.Union(
      schemaTypes.map((type) => convertJsonSchemaToTypeBoxSchema({ ...source, type }, `${path}<${type}>`)),
      readTypeBoxOptions(source, ["type"])
    );
  }

  const schemaType = schemaTypes[0] ?? inferJsonSchemaType(source);
  switch (schemaType) {
    case "object":
      return convertJsonObjectSchemaToTypeBox(source, path);
    case "array":
      return convertJsonArraySchemaToTypeBox(source, path);
    case "string":
      return Type.String(readTypeBoxOptions(source, ["type"]));
    case "integer":
      return Type.Integer(readTypeBoxOptions(source, ["type"]));
    case "number":
      return Type.Number(readTypeBoxOptions(source, ["type"]));
    case "boolean":
      return Type.Boolean(readTypeBoxOptions(source, ["type"]));
    case "null":
      return Type.Null(readTypeBoxOptions(source, ["type"]));
    default:
      return Type.Unknown(readTypeBoxOptions(source, ["type"]));
  }
}

function convertJsonObjectSchemaToTypeBox(schema: JsonSchemaObject, path: string): TSchema {
  const required = new Set(readRequiredProperties(schema, path));
  const properties = schema.properties ?? {};
  const propertySchemas: Record<string, TSchema> = {};

  for (const requiredKey of required) {
    if (!(requiredKey in properties)) {
      throw new Error(`Unsupported JSON schema at ${path}: required property "${requiredKey}" is not defined`);
    }
  }

  for (const [key, propertySchema] of Object.entries(properties)) {
    const converted = convertJsonSchemaToTypeBoxSchema(propertySchema, `${path}.properties.${key}`);
    propertySchemas[key] = required.has(key) ? converted : Type.Optional(converted);
  }

  return Type.Object(propertySchemas, readTypeBoxOptions(schema, ["type", "properties", "required"]));
}

function convertJsonArraySchemaToTypeBox(schema: JsonSchemaObject, path: string): TSchema {
  const options = readTypeBoxOptions(schema, ["type", "items"]);
  if (Array.isArray(schema.items)) {
    return Type.Tuple(
      schema.items.map((item, index) => convertJsonSchemaToTypeBoxSchema(item, `${path}.items[${index}]`)),
      options
    );
  }
  return Type.Array(
    schema.items === undefined ? Type.Unknown() : convertJsonSchemaToTypeBoxSchema(schema.items, `${path}.items`),
    options
  );
}

function readJsonSchemaObject(schema: unknown, path: string): JsonSchemaObject {
  if (!isRecord(schema) || Array.isArray(schema)) {
    throw new Error(`Unsupported JSON schema at ${path}: expected an object`);
  }
  return schema as JsonSchemaObject;
}

function readJsonSchemaTypes(schema: JsonSchemaObject, path: string): JsonSchemaTypeName[] {
  if (schema.type === undefined) return [];
  const types = Array.isArray(schema.type) ? schema.type : [schema.type];
  for (const type of types) {
    if (
      type !== "object" &&
      type !== "array" &&
      type !== "string" &&
      type !== "integer" &&
      type !== "number" &&
      type !== "boolean" &&
      type !== "null"
    ) {
      throw new Error(`Unsupported JSON schema at ${path}: unknown type "${String(type)}"`);
    }
  }
  return types;
}

function readRequiredProperties(schema: JsonSchemaObject, path: string): string[] {
  if (schema.required === undefined) return [];
  if (!Array.isArray(schema.required) || schema.required.some((item) => typeof item !== "string")) {
    throw new Error(`Unsupported JSON schema at ${path}: required must be a string array`);
  }
  return schema.required;
}

function inferJsonSchemaType(schema: JsonSchemaObject): JsonSchemaTypeName | undefined {
  if (schema.properties || schema.additionalProperties !== undefined) return "object";
  if (schema.items) return "array";
  return undefined;
}

function readTypeBoxOptions(schema: JsonSchemaObject, omittedKeys: string[]): Record<string, unknown> {
  const omitted = new Set(omittedKeys);
  const options: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(schema)) {
    if (omitted.has(key) || value === undefined) continue;
    if (key === "additionalProperties" && isRecord(value) && !Array.isArray(value)) {
      options.additionalProperties = convertJsonSchemaToTypeBoxSchema(value, "additionalProperties");
      continue;
    }
    options[key] = cloneJsonSchemaValue(value);
  }
  return options;
}

function cloneJsonSchemaObject(schema: JsonSchemaObject): TSchema {
  return cloneJsonSchemaValue(schema) as TSchema;
}

function cloneJsonSchemaValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map((item) => cloneJsonSchemaValue(item));
  if (!isRecord(value)) return value;
  return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, cloneJsonSchemaValue(item)]));
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function resolvePwdSafeToolExecutionMode(toolName: string): "parallel" | "sequential" {
  if (
    toolName === "time" ||
    toolName === "web_search" ||
    toolName === "read_file" ||
    toolName === "read_word" ||
    toolName === "read_pdf" ||
    toolName === "read_image" ||
    toolName === "plan_scheme_batches" ||
    toolName === "plan_scheme_assets" ||
    toolName === "draft_scheme_sections" ||
    toolName === "image_generate"
  ) {
    return "parallel";
  }
  return "sequential";
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

export function buildPwdSafePiToolSchemaSignature(settings: AppSettings): string {
  return JSON.stringify({
    version: PI_TOOL_SCHEMA_VERSION,
    tools: buildAgentChatTools({
      includeExecBash: settings.agent.execBashEnabled,
      includeArtifactTools: true
    })
      .filter((tool) => tool.type === "function")
      .map((tool) => ({
        name: tool.function.name,
        description: tool.function.description,
        parameters: tool.function.parameters,
        strict: tool.function.strict,
        executionMode: resolvePwdSafeToolExecutionMode(tool.function.name)
      }))
  });
}

function buildConfigSignature(settings: AppSettings): string {
  return JSON.stringify({
    baseUrl: settings.openai.baseUrl,
    chatModel: settings.openai.chatModel,
    thinkingEnabled: settings.openai.thinkingEnabled,
    reasoningEffort: settings.openai.reasoningEffort,
    maxOutputTokens: settings.openai.maxOutputTokens,
    execBashEnabled: settings.agent.execBashEnabled,
    draftSectionParallelism: settings.agent.draftSectionParallelism,
    imageGenerationParallelism: settings.agent.imageGenerationParallelism,
    toolSchemaSignature: buildPwdSafePiToolSchemaSignature(settings)
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

function getLatestSchemeProgress(host: AgentRuntimeHost, sessionId: string): SchemeProgressItem | undefined {
  const session = host.getSession(sessionId);
  const items = session?.items ?? [];
  for (let index = items.length - 1; index >= 0; index -= 1) {
    const item = items[index];
    if (item.kind === "scheme_progress") return item;
  }
  return undefined;
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
