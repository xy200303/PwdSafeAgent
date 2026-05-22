import electron from "electron";
import type { BrowserWindow as BrowserWindowType } from "electron";
import { join, extname, basename } from "node:path";
import { existsSync, mkdirSync, statSync, writeFileSync } from "node:fs";
import OpenAI from "openai";
import type {
  ChatCompletionAssistantMessageParam,
  ChatCompletionMessageParam,
  ChatCompletionToolMessageParam
} from "openai/resources/chat/completions";
import dotenv from "dotenv";
import {
  compactText,
  getCurrentTimeText,
  getReadToolName,
  readDocumentText,
  sanitizeFileName,
  writeUtf8File
} from "./agentTools";
import { buildAgentChatTools, executeAgentToolCall } from "./agentToolRegistry";
import { shouldCreateDraftArtifact } from "./artifactIntent";
import { buildArtifactPreview } from "./artifactPreview";
import { exportDocxToPdf } from "./documentExport";
import { generateDiagramImage, shouldGenerateDiagramArtifacts, type DiagramKind } from "./imageGeneration";
import { writeSchemeDocxFromTemplate, type SchemeDiagramAsset } from "./schemeDocument";
import { loadPersistedState, savePersistedState, type PersistedStateSnapshot, type SessionMemoryEntry } from "./sessionPersistence";
import type {
  AppSettings,
  ArtifactKind,
  ArtifactListInput,
  ArtifactSummary,
  AttachmentRef,
  ChatPromptInput,
  ChatSession,
  ClipboardAttachmentInput,
  PickAttachmentInput,
  RenameSessionInput,
  RendererEvent,
  StreamItem,
  UpdateAppSettingsInput
} from "../shared/types";

type MessageStreamItem = Extract<StreamItem, { kind: "message" }>;

const rootDir = process.cwd();
const { app, BrowserWindow, dialog, ipcMain, shell } = electron;
const dataDir = join(rootDir, "data");
const inputDir = join(dataDir, "input");
const outputDir = join(dataDir, "output");
const statePath = join(dataDir, "state.json");
const envLocalPath = join(rootDir, ".env.local");
const envPath = join(rootDir, ".env");

let mainWindow: BrowserWindowType | null = null;
const sessions = new Map<string, ChatSession>();
const attachments = new Map<string, AttachmentRef>();
const artifacts = new Map<string, ArtifactSummary>();
const abortControllers = new Map<string, AbortController>();
const templateLoadedSessions = new Set<string>();
const sessionMemories = new Map<string, SessionMemoryEntry[]>();
let persistTimer: NodeJS.Timeout | undefined;
const MAX_DOCUMENT_CONTEXT_CHARS = 24000;
const MAX_TEMPLATE_CONTEXT_CHARS = 36000;
const MAX_SESSION_MEMORY_CHARS = 90000;
const MAX_AGENT_TOOL_ROUNDS = 3;

function ensureDataDirs(): void {
  for (const dir of [dataDir, inputDir, outputDir]) {
    if (!existsSync(dir)) {
      mkdirSync(dir, { recursive: true });
    }
  }
}

function createId(prefix: string): string {
  return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
}

function now(): string {
  return new Date().toISOString();
}

function parseBooleanEnv(value: string | undefined, fallback: boolean): boolean {
  if (value === undefined) return fallback;
  return /^(1|true|yes|on)$/i.test(value.trim());
}

function sendEvent(event: RendererEvent): void {
  mainWindow?.webContents.send("app:event", event);
}

function snapshotState(): PersistedStateSnapshot {
  const memoryRecord: Record<string, SessionMemoryEntry[]> = {};
  for (const [sessionId, memory] of sessionMemories) {
    memoryRecord[sessionId] = memory;
  }

  return {
    sessions: Array.from(sessions.values()),
    attachments: Array.from(attachments.values()),
    artifacts: Array.from(artifacts.values()),
    templateLoadedSessionIds: Array.from(templateLoadedSessions),
    sessionMemories: memoryRecord
  };
}

function restorePersistedState(): void {
  try {
    const state = loadPersistedState(statePath);
    if (!state) return;

    sessions.clear();
    attachments.clear();
    artifacts.clear();
    templateLoadedSessions.clear();
    sessionMemories.clear();

    for (const session of state.sessions) sessions.set(session.id, session);
    for (const attachment of state.attachments) attachments.set(attachment.id, attachment);
    for (const artifact of state.artifacts) artifacts.set(artifact.id, artifact);
    for (const sessionId of state.templateLoadedSessionIds) templateLoadedSessions.add(sessionId);
    for (const [sessionId, memory] of Object.entries(state.sessionMemories)) sessionMemories.set(sessionId, memory);
  } catch (error) {
    console.warn("Failed to restore persisted state:", error);
  }
}

function schedulePersistState(): void {
  if (persistTimer) clearTimeout(persistTimer);
  persistTimer = setTimeout(() => {
    persistTimer = undefined;
    flushPersistedState();
  }, 250);
}

function flushPersistedState(): void {
  if (persistTimer) {
    clearTimeout(persistTimer);
    persistTimer = undefined;
  }
  try {
    savePersistedState(statePath, snapshotState());
  } catch (error) {
    console.warn("Failed to persist state:", error);
  }
}

function loadEnv(): AppSettings {
  const envFilePath = existsSync(envLocalPath) ? envLocalPath : existsSync(envPath) ? envPath : "";
  if (envFilePath) {
    dotenv.config({ path: envFilePath, override: true });
  }

  return {
    runtime: {
      envFilePath,
      configSource: existsSync(envLocalPath) ? ".env.local" : existsSync(envPath) ? ".env" : "process"
    },
    openai: {
      baseUrl: process.env.OPENAI_BASE_URL || "https://api.openai.com/v1",
      imageBaseUrl: process.env.OPENAI_IMAGE_BASE_URL || "",
      chatModel: process.env.OPENAI_CHAT_MODEL || "gpt-5.5",
      imageModel: process.env.OPENAI_IMAGE_MODEL || "gpt-image-2",
      imageSize: process.env.OPENAI_IMAGE_SIZE || "1536x1024",
      imageQuality: process.env.OPENAI_IMAGE_QUALITY || "high",
      autoImageGeneration: parseBooleanEnv(process.env.OPENAI_AUTO_IMAGE_GENERATION, true),
      requestTimeoutMs: Number(process.env.OPENAI_REQUEST_TIMEOUT_MS || 120000),
      maxOutputTokens: Number(process.env.OPENAI_MAX_OUTPUT_TOKENS || 16000),
      apiKeyConfigured: Boolean(process.env.OPENAI_API_KEY),
      imageApiKeyConfigured: Boolean(process.env.OPENAI_IMAGE_API_KEY)
    },
    document: {
      autoPdfExport: parseBooleanEnv(process.env.AGENT_AUTO_PDF_EXPORT, false),
      libreOfficePath: process.env.LIBREOFFICE_PATH || ""
    },
    agent: {
      execBashEnabled: parseBooleanEnv(process.env.AGENT_EXEC_BASH_ENABLED, false)
    }
  };
}

function saveEnv(input: UpdateAppSettingsInput): AppSettings {
  const currentKey = process.env.OPENAI_API_KEY || "";
  const currentImageKey = process.env.OPENAI_IMAGE_API_KEY || "";
  const lines = [
    `OPENAI_API_KEY=${input.openai.apiKey ?? currentKey}`,
    `OPENAI_BASE_URL=${input.openai.baseUrl}`,
    `OPENAI_IMAGE_BASE_URL=${input.openai.imageBaseUrl}`,
    `OPENAI_IMAGE_API_KEY=${input.openai.imageApiKey ?? currentImageKey}`,
    `OPENAI_CHAT_MODEL=${input.openai.chatModel}`,
    `OPENAI_IMAGE_MODEL=${input.openai.imageModel}`,
    `OPENAI_IMAGE_SIZE=${input.openai.imageSize}`,
    `OPENAI_IMAGE_QUALITY=${input.openai.imageQuality}`,
    `OPENAI_AUTO_IMAGE_GENERATION=${input.openai.autoImageGeneration ? "true" : "false"}`,
    `OPENAI_REQUEST_TIMEOUT_MS=${input.openai.requestTimeoutMs}`,
    `OPENAI_MAX_OUTPUT_TOKENS=${input.openai.maxOutputTokens}`,
    `AGENT_EXEC_BASH_ENABLED=${input.agent.execBashEnabled ? "true" : "false"}`,
    `AGENT_AUTO_PDF_EXPORT=${input.document.autoPdfExport ? "true" : "false"}`,
    `LIBREOFFICE_PATH=${input.document.libreOfficePath}`
  ];
  writeFileSync(envLocalPath, `${lines.join("\n")}\n`, "utf-8");
  dotenv.config({ path: envLocalPath, override: true });
  return loadEnv();
}

function createSession(): ChatSession {
  const session: ChatSession = {
    id: createId("session"),
    title: "新的密码方案对话",
    status: "idle",
    createdAt: now(),
    updatedAt: now(),
    items: []
  };
  sessions.set(session.id, session);
  schedulePersistState();
  sendEvent({ id: createId("event"), type: "session.created", payload: session });
  return session;
}

function renameSession(input: RenameSessionInput): ChatSession {
  const session = sessions.get(input.sessionId);
  if (!session) {
    throw new Error(`Session not found: ${input.sessionId}`);
  }

  const title = input.title.replace(/[\r\n\t]+/g, " ").trim();
  if (!title) {
    throw new Error("Session title cannot be empty");
  }

  session.title = title.slice(0, 80);
  session.updatedAt = now();
  sessions.set(session.id, session);
  schedulePersistState();
  sendEvent({ id: createId("event"), type: "session.updated", payload: session });
  return session;
}

function deleteSession(sessionId: string): ChatSession[] {
  const session = sessions.get(sessionId);
  if (!session) {
    throw new Error(`Session not found: ${sessionId}`);
  }

  abortControllers.get(sessionId)?.abort();
  abortControllers.delete(sessionId);
  sessions.delete(sessionId);
  templateLoadedSessions.delete(sessionId);
  sessionMemories.delete(sessionId);

  for (const [attachmentId, attachment] of attachments) {
    if (attachment.sessionId === sessionId) {
      attachments.delete(attachmentId);
    }
  }

  for (const item of session.items) {
    if (item.kind === "file") {
      artifacts.delete(item.artifactId);
    }
  }

  schedulePersistState();
  sendEvent({ id: createId("event"), type: "session.deleted", sessionId });

  if (sessions.size === 0) {
    createSession();
  }

  return Array.from(sessions.values());
}

function addItem(sessionId: string, item: StreamItem): StreamItem {
  const session = sessions.get(sessionId);
  if (!session) {
    throw new Error(`Session not found: ${sessionId}`);
  }
  session.items.push(item);
  session.updatedAt = now();
  sessions.set(sessionId, session);
  schedulePersistState();
  sendEvent({ id: createId("event"), type: "stream.item.added", sessionId, payload: item });
  sendEvent({ id: createId("event"), type: "session.updated", payload: session });
  return item;
}

function updateItem(sessionId: string, item: StreamItem): void {
  const session = sessions.get(sessionId);
  if (!session) {
    throw new Error(`Session not found: ${sessionId}`);
  }
  const index = session.items.findIndex((existing) => existing.id === item.id);
  if (index >= 0) {
    session.items[index] = item;
  }
  session.updatedAt = now();
  sessions.set(sessionId, session);
  schedulePersistState();
  sendEvent({ id: createId("event"), type: "stream.item.updated", sessionId, payload: item });
  sendEvent({ id: createId("event"), type: "session.updated", payload: session });
}

function setSessionStatus(sessionId: string, status: ChatSession["status"]): void {
  const session = sessions.get(sessionId);
  if (!session) return;
  session.status = status;
  session.updatedAt = now();
  sessions.set(sessionId, session);
  schedulePersistState();
  sendEvent({ id: createId("event"), type: "session.updated", payload: session });
}

function startToolCall(sessionId: string, toolName: string, summary: string): StreamItem {
  return addItem(sessionId, {
    id: createId("tool"),
    kind: "tool",
    toolCallId: createId("toolcall"),
    toolName,
    status: "running",
    summary,
    createdAt: now()
  });
}

function finishToolCall(sessionId: string, item: StreamItem, status: "success" | "failed", summary: string): void {
  if (item.kind !== "tool") return;
  item.status = status;
  item.summary = summary;
  updateItem(sessionId, item);
}

function appendSessionMemory(sessionId: string, source: string, content: string): void {
  const memory = sessionMemories.get(sessionId) ?? [];
  memory.push({ source, content: compactText(content, MAX_DOCUMENT_CONTEXT_CHARS) });

  let total = memory.reduce((sum, item) => sum + item.content.length, 0);
  while (total > MAX_SESSION_MEMORY_CHARS && memory.length > 1) {
    const removed = memory.shift();
    total -= removed?.content.length ?? 0;
  }
  sessionMemories.set(sessionId, memory);
  schedulePersistState();
}

function formatSessionMemory(sessionId: string): string {
  const memory = sessionMemories.get(sessionId) ?? [];
  if (!memory.length) return "";
  return memory.map((item) => `## ${item.source}\n${item.content}`).join("\n\n");
}

function getArtifactKind(filePath: string): ArtifactKind {
  const ext = extname(filePath).toLowerCase();
  if (ext === ".docx") return "docx";
  if (ext === ".pdf") return "pdf";
  if (ext === ".png") return "png";
  if (ext === ".jpg") return "jpg";
  if (ext === ".jpeg") return "jpeg";
  if (ext === ".webp") return "webp";
  if (ext === ".svg") return "svg";
  if (ext === ".json") return "json";
  if (ext === ".md") return "md";
  if (ext === ".txt") return "txt";
  if (ext === ".log") return "log";
  return "other";
}

function createArtifact(sessionId: string, filePath: string, name = basename(filePath)): ArtifactSummary {
  const artifact: ArtifactSummary = {
    id: createId("artifact"),
    sessionId,
    name,
    kind: getArtifactKind(filePath),
    path: filePath,
    size: existsSync(filePath) ? statSync(filePath).size : 0,
    createdAt: now()
  };
  artifacts.set(artifact.id, artifact);
  addItem(sessionId, {
    id: createId("file"),
    kind: "file",
    artifactId: artifact.id,
    name: artifact.name,
    fileKind: artifact.kind,
    createdAt: now()
  });
  sendEvent({ id: createId("event"), type: "artifact.created", sessionId, payload: artifact });
  schedulePersistState();
  return artifact;
}

function listArtifacts(input?: ArtifactListInput): ArtifactSummary[] {
  const allArtifacts = Array.from(artifacts.values());
  const filtered = input?.sessionId
    ? allArtifacts.filter((artifact) => artifactBelongsToSession(artifact, input.sessionId!))
    : allArtifacts;

  return filtered.sort((left, right) => right.createdAt.localeCompare(left.createdAt));
}

function artifactBelongsToSession(artifact: ArtifactSummary, sessionId: string): boolean {
  if (artifact.sessionId === sessionId) return true;
  const session = sessions.get(sessionId);
  return Boolean(session?.items.some((item) => item.kind === "file" && item.artifactId === artifact.id));
}

function buildMessages(session: ChatSession): ChatCompletionMessageParam[] {
  const systemPrompt = [
    "你是专业的密码应用方案生成 Agent。",
    "你需要严格参考用户给出的 Word 模板结构，生成专业、可复核、可落地的密码应用方案内容。",
    "方案需要覆盖系统概况、密码应用需求、密码应用设计、密钥管理、实施计划、风险与符合性说明等章节。",
    "回复使用中文 Markdown，必要时给出缺失资料清单。",
    "不要编造用户未提供的关键事实；若资料不足，用“待补充/需确认”标识。",
    "工具调用由你按任务需要自主决策：寒暄、普通问答和资料澄清阶段不要默认读取模板或附件；只有生成/完善方案、分析附件、导出文件、查询最新资料等确有需要时，才调用对应工具。",
    `当前时间：${getCurrentTimeText()}`
  ].join("\n");

  const messages: ChatCompletionMessageParam[] = [{ role: "system", content: systemPrompt }];
  const resources = buildAvailableResourceContext(session);
  if (resources) {
    messages.push({
      role: "system",
      content: resources
    });
  }

  const memory = formatSessionMemory(session.id);
  if (memory) {
    messages.push({
      role: "system",
      content: `以下是本会话内置工具读取到的模板与附件上下文，生成方案时必须优先参考：\n\n${memory}`
    });
  }

  for (const item of session.items) {
    if (item.kind !== "message") continue;
    if (!item.content.trim()) continue;
    messages.push({
      role: item.role === "user" ? "user" : "assistant",
      content: item.content
    });
  }
  return messages;
}

function buildAvailableResourceContext(session: ChatSession): string {
  const lines = [
    "以下是本会话可按需读取的文件资源。注意：这些文件尚未读取；只有在用户任务需要时才调用 read_word/read_pdf/read_file。",
    "Word 方案模板：docs/密码应用方案.docx"
  ];
  const sessionAttachments = getSessionAttachments(session.id);
  if (sessionAttachments.length) {
    lines.push(
      "用户附件：",
      ...sessionAttachments.map((attachment, index) => `${index + 1}. ${attachment.name} (${attachment.path})`)
    );
  }

  return lines.join("\n");
}

function getSessionAttachments(sessionId: string): AttachmentRef[] {
  const session = sessions.get(sessionId);
  const ids = new Set<string>();
  for (const item of session?.items ?? []) {
    if (item.kind === "message") {
      for (const attachmentId of item.attachmentIds ?? []) {
        ids.add(attachmentId);
      }
    }
  }

  return Array.from(attachments.values()).filter(
    (attachment) => attachment.sessionId === sessionId || ids.has(attachment.id)
  );
}

function getSessionReadableFiles(sessionId: string): string[] {
  return getSessionAttachments(sessionId).map((attachment) => attachment.path);
}

async function streamMockResponse(sessionId: string, assistantItem: StreamItem): Promise<void> {
  if (assistantItem.kind !== "message") return;
  const memory = sessionMemories.get(sessionId) ?? [];
  const hasContext = memory.length > 0;
  const chunks = [
    "已收到资料，并已通过内置工具读取模板与附件上下文。\n\n",
    hasContext ? `当前会话已纳入 ${memory.length} 份上下文材料。\n\n` : "当前还没有可解析的附件内容。\n\n",
    "我会按密码应用方案模板推进：\n\n",
    "1. 建立系统事实模型：建设单位、系统边界、业务场景、数据类型和等保级别。\n",
    "2. 对照 `GB/T 39786-2021` 组织密码应用需求与差距分析。\n",
    "3. 输出密码应用总体架构、典型流程、密钥管理、产品部署和实施计划。\n",
    "4. 对缺失材料使用“待补充/需确认”标识，避免编造。\n\n",
    "当前处于本地演示模式。配置 `OPENAI_API_KEY` 后，会切换为 OpenAI Chat Completions 的真实流式生成。"
  ];

  for (const chunk of chunks) {
    await new Promise((resolvePromise) => setTimeout(resolvePromise, 140));
    assistantItem.content += chunk;
    updateItem(sessionId, assistantItem);
  }
}

function createAssistantMessage(sessionId: string): MessageStreamItem {
  return addItem(sessionId, {
    id: createId("msg"),
    kind: "message",
    role: "assistant",
    content: "",
    isFinished: false,
    createdAt: now()
  }) as MessageStreamItem;
}

async function runOpenAIResponse(
  sessionId: string,
  controller: AbortController,
  userPrompt: string,
  onAssistantCreated: (assistantItem: MessageStreamItem) => void
): Promise<MessageStreamItem> {
  const settings = loadEnv();
  const session = sessions.get(sessionId);
  if (!session) throw new Error("Session not found");

  if (!process.env.OPENAI_API_KEY) {
    const assistantItem = createAssistantMessage(sessionId);
    onAssistantCreated(assistantItem);
    await streamMockResponse(sessionId, assistantItem);
    return assistantItem;
  }

  const client = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY,
    baseURL: settings.openai.baseUrl,
    timeout: settings.openai.requestTimeoutMs
  });

  const messages = await runOpenAIToolPlanning(sessionId, client, settings, buildMessages(session), controller, userPrompt);
  messages.push({
    role: "system",
    content: "工具调用阶段已结束。请基于用户需求、模板上下文、附件内容和工具结果，输出最终中文 Markdown 回复。"
  });

  const assistantItem = createAssistantMessage(sessionId);
  onAssistantCreated(assistantItem);
  const stream = await client.chat.completions.create(
    {
      model: settings.openai.chatModel,
      messages,
      stream: true,
      max_completion_tokens: settings.openai.maxOutputTokens
    },
    { signal: controller.signal }
  );

  for await (const part of stream) {
    const delta = part.choices[0]?.delta?.content;
    if (!delta) continue;
    assistantItem.content += delta;
    updateItem(sessionId, assistantItem);
  }

  return assistantItem;
}

async function runOpenAIToolPlanning(
  sessionId: string,
  client: OpenAI,
  settings: AppSettings,
  messages: ChatCompletionMessageParam[],
  controller: AbortController,
  userPrompt: string
): Promise<ChatCompletionMessageParam[]> {
  const tools = buildAgentChatTools({ includeExecBash: settings.agent.execBashEnabled });
  const toolMessages = [...messages];

  for (let round = 1; round <= MAX_AGENT_TOOL_ROUNDS; round += 1) {
    throwIfAborted(controller);
    const plannerTool = startToolCall(sessionId, "openai.chat.tools", `工具规划第 ${round} 轮`);
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
      finishToolCall(sessionId, plannerTool, "success", "模型判断无需继续调用工具");
      if (message?.content?.trim()) {
        toolMessages.push({ role: "assistant", content: message.content });
      }
      return toolMessages;
    }

    finishToolCall(sessionId, plannerTool, "success", `模型请求 ${calls.length} 个工具调用`);
    const assistantToolMessage: ChatCompletionAssistantMessageParam = {
      role: "assistant",
      content: message.content ?? null,
      tool_calls: calls
    };
    toolMessages.push(assistantToolMessage);

    for (const call of calls) {
      throwIfAborted(controller);
      const functionName = call.type === "function" ? call.function.name : call.type;
      const toolItem = startToolCall(sessionId, functionName, `执行 ${functionName}`);
      try {
        const result = await executeAgentToolCall(call, {
          rootDir,
          outputDir,
          sessionTitle: sanitizeFileName(sessions.get(sessionId)?.title || "密码应用方案"),
          memory: formatSessionMemory(sessionId),
          settings,
          userPrompt,
          signal: controller.signal,
          allowedReadDirs: [join(rootDir, "docs"), inputDir, outputDir],
          allowedReadFiles: getSessionReadableFiles(sessionId),
          execBashEnabled: settings.agent.execBashEnabled
        });
        finishToolCall(sessionId, toolItem, "success", result.summary);

        if (result.content.trim()) {
          appendSessionMemory(sessionId, `工具结果：${result.toolName}`, result.content);
        }
        if (result.artifactPath) {
          createArtifact(sessionId, result.artifactPath);
        }

        const nextToolMessage: ChatCompletionToolMessageParam = {
          role: "tool",
          tool_call_id: call.id,
          content: compactText(result.content || result.summary, 16000)
        };
        toolMessages.push(nextToolMessage);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        finishToolCall(sessionId, toolItem, "failed", message);
        toolMessages.push({
          role: "tool",
          tool_call_id: call.id,
          content: `Tool failed: ${message}`
        });
      }
    }
  }

  addItem(sessionId, {
    id: createId("stage"),
    kind: "stage",
    title: "工具规划轮次已达上限",
    detail: `${MAX_AGENT_TOOL_ROUNDS} 轮`,
    createdAt: now()
  });
  return toolMessages;
}

async function prepareAgentContext(sessionId: string, input: ChatPromptInput, controller: AbortController): Promise<void> {
  throwIfAborted(controller);
  const timeTool = startToolCall(sessionId, "time", "获取当前时间");
  finishToolCall(sessionId, timeTool, "success", getCurrentTimeText());

  await ensureTemplateContext(sessionId, controller);

  if (!input.attachments?.length) return;
  addItem(sessionId, {
    id: createId("stage"),
    kind: "stage",
    title: "解析用户附件",
    detail: `${input.attachments.length} 个文件`,
    createdAt: now()
  });

  for (const attachment of input.attachments) {
    throwIfAborted(controller);
    const toolName = getReadToolName(attachment.path);
    const tool = startToolCall(sessionId, toolName, `读取 ${attachment.name}`);
    try {
      const result = await readDocumentText(attachment.path, MAX_DOCUMENT_CONTEXT_CHARS);
      if (result.content.trim()) {
        appendSessionMemory(sessionId, `用户附件：${result.sourceName}`, result.content);
      }
      finishToolCall(sessionId, tool, "success", result.summary);
    } catch (error) {
      finishToolCall(sessionId, tool, "failed", error instanceof Error ? error.message : String(error));
    }
  }
}

async function ensureTemplateContext(sessionId: string, controller: AbortController): Promise<void> {
  if (templateLoadedSessions.has(sessionId)) return;
  throwIfAborted(controller);
  const templatePath = join(rootDir, "docs", "密码应用方案.docx");
  const tool = startToolCall(sessionId, "read_word", "读取 Word 方案模板");
  try {
    const result = await readDocumentText(templatePath, MAX_TEMPLATE_CONTEXT_CHARS);
    appendSessionMemory(sessionId, `Word 模板：${result.sourceName}`, result.content);
    templateLoadedSessions.add(sessionId);
    finishToolCall(sessionId, tool, "success", result.summary);
  } catch (error) {
    finishToolCall(sessionId, tool, "failed", error instanceof Error ? error.message : String(error));
  }
}

async function emitDraftArtifact(
  sessionId: string,
  prompt: string,
  assistantItem: StreamItem,
  controller: AbortController
): Promise<void> {
  if (assistantItem.kind !== "message") return;
  if (!shouldCreateDraftArtifact(prompt, assistantItem.content)) return;

  const session = sessions.get(sessionId);
  const baseName = sanitizeFileName(session?.title || "密码应用方案草稿");
  const artifactStamp = Date.now().toString(36);
  const filePath = join(outputDir, `${baseName}-${artifactStamp}.md`);
  const writeTool = startToolCall(sessionId, "write_file", "写入方案草稿 Markdown");
  try {
    await writeUtf8File(filePath, assistantItem.content);
    finishToolCall(sessionId, writeTool, "success", `已写入 ${basename(filePath)}`);
  } catch (error) {
    finishToolCall(sessionId, writeTool, "failed", error instanceof Error ? error.message : String(error));
    return;
  }

  const sendTool = startToolCall(sessionId, "send_file", "发送文件给前端");
  createArtifact(sessionId, filePath);
  finishToolCall(sessionId, sendTool, "success", `已发送 ${basename(filePath)}`);

  const diagramAssets = await emitDiagramArtifacts(sessionId, prompt, assistantItem, baseName, artifactStamp, controller);
  const docxPath = join(outputDir, `${baseName}-${artifactStamp}.docx`);
  const writeWordTool = startToolCall(sessionId, "write_word", "按 Word 模板生成方案文档");
  try {
    const result = await writeSchemeDocxFromTemplate(join(rootDir, "docs", "密码应用方案.docx"), docxPath, {
      prompt,
      memory: formatSessionMemory(sessionId),
      generatedMarkdown: assistantItem.content,
      diagrams: diagramAssets
    });
    const diagramSummary = result.embeddedDiagrams.length ? `，嵌入图示 ${result.embeddedDiagrams.length} 张` : "";
    finishToolCall(
      sessionId,
      writeWordTool,
      "success",
      `已生成 ${result.fileName}，填充 ${result.filledFields.length} 项，待补充 ${result.missingFields.length} 项${diagramSummary}`
    );
  } catch (error) {
    finishToolCall(sessionId, writeWordTool, "failed", error instanceof Error ? error.message : String(error));
    return;
  }

  const sendWordTool = startToolCall(sessionId, "send_file", "发送 Word 文档给前端");
  createArtifact(sessionId, docxPath);
  finishToolCall(sessionId, sendWordTool, "success", `已发送 ${basename(docxPath)}`);

  await emitPdfArtifact(sessionId, docxPath, controller);
}

async function emitPdfArtifact(sessionId: string, docxPath: string, controller: AbortController): Promise<void> {
  const settings = loadEnv();
  if (!settings.document.autoPdfExport) return;
  throwIfAborted(controller);

  const pdfTool = startToolCall(sessionId, "write_pdf", "导出 PDF 方案文档");
  const result = await exportDocxToPdf(docxPath, outputDir, {
    libreOfficePath: settings.document.libreOfficePath,
    timeoutMs: settings.openai.requestTimeoutMs
  });

  if (result.status === "success") {
    finishToolCall(sessionId, pdfTool, "success", result.summary);
    const sendPdfTool = startToolCall(sessionId, "send_file", "发送 PDF 文档给前端");
    createArtifact(sessionId, result.outputPath);
    finishToolCall(sessionId, sendPdfTool, "success", `已发送 ${result.fileName}`);
    return;
  }

  finishToolCall(sessionId, pdfTool, result.status === "unavailable" ? "success" : "failed", result.summary);
}

async function emitDiagramArtifacts(
  sessionId: string,
  prompt: string,
  assistantItem: MessageStreamItem,
  baseName: string,
  artifactStamp: string,
  controller: AbortController
): Promise<SchemeDiagramAsset[]> {
  const settings = loadEnv();
  if (!settings.openai.autoImageGeneration) return [];
  if (!shouldGenerateDiagramArtifacts(prompt, assistantItem.content)) return [];

  addItem(sessionId, {
    id: createId("stage"),
    kind: "stage",
    title: "生成方案配图",
    detail: "技术架构图 / 业务流程图",
    createdAt: now()
  });

  const diagrams: Array<{ kind: DiagramKind; label: string }> = [
    { kind: "architecture", label: "密码应用技术架构图" },
    { kind: "flow", label: "典型业务密码应用流程图" }
  ];
  const assets: SchemeDiagramAsset[] = [];

  for (const diagram of diagrams) {
    throwIfAborted(controller);
    const imageTool = startToolCall(sessionId, "image_generate", `生成${diagram.label}`);
    try {
      const result = await generateDiagramImage(
        {
          apiKey: process.env.OPENAI_IMAGE_API_KEY || process.env.OPENAI_API_KEY,
          baseUrl: settings.openai.imageBaseUrl || settings.openai.baseUrl,
          imageModel: settings.openai.imageModel,
          imageSize: settings.openai.imageSize,
          imageQuality: settings.openai.imageQuality,
          requestTimeoutMs: settings.openai.requestTimeoutMs
        },
        {
          kind: diagram.kind,
          sessionTitle: baseName,
          prompt,
          memory: formatSessionMemory(sessionId),
          generatedMarkdown: assistantItem.content,
          outputDir,
          artifactStamp
        },
        controller.signal
      );
      const modeText = result.mode === "openai" ? result.model : "本地 SVG 占位图";
      finishToolCall(sessionId, imageTool, "success", `已生成 ${result.fileName} (${modeText})`);

      const sendImageTool = startToolCall(sessionId, "send_file", `发送${diagram.label}`);
      createArtifact(sessionId, result.outputPath);
      finishToolCall(sessionId, sendImageTool, "success", `已发送 ${result.fileName}`);
      assets.push({
        label: diagram.label,
        kind: diagram.kind,
        path: result.outputPath
      });
    } catch (error) {
      finishToolCall(sessionId, imageTool, "failed", error instanceof Error ? error.message : String(error));
    }
  }

  return assets;
}

function throwIfAborted(controller: AbortController): void {
  if (controller.signal.aborted) {
    throw new Error("用户已停止生成");
  }
}

async function handlePrompt(input: ChatPromptInput): Promise<{ accepted: true }> {
  const session = sessions.get(input.sessionId);
  if (!session) {
    throw new Error("Session not found");
  }

  if (session.title === "新的密码方案对话" && input.message.trim()) {
    session.title = input.message.trim().slice(0, 24);
  }

  addItem(input.sessionId, {
    id: createId("msg"),
    kind: "message",
    role: "user",
    content: input.message,
    isFinished: true,
    createdAt: now(),
    attachmentIds: input.attachments?.map((attachment) => attachment.id)
  });

  setSessionStatus(input.sessionId, "running");
  const controller = new AbortController();
  abortControllers.set(input.sessionId, controller);
  let toolItem: StreamItem | undefined;
  let assistantItem: MessageStreamItem | undefined;

  void Promise.resolve()
    .then(async () => {
      toolItem = startToolCall(
        input.sessionId,
        "openai.chat.completions",
        "正在使用 OpenAI Chat Completions 生成流式回复"
      );

      assistantItem = await runOpenAIResponse(input.sessionId, controller, input.message, (createdItem) => {
        assistantItem = createdItem;
      });
      assistantItem.isFinished = true;
      updateItem(input.sessionId, assistantItem);
      finishToolCall(input.sessionId, toolItem, "success", "流式回复完成");
      await emitDraftArtifact(input.sessionId, input.message, assistantItem, controller);
    })
    .then(() => {
      setSessionStatus(input.sessionId, "completed");
    })
    .catch((error: unknown) => {
      if (!sessions.has(input.sessionId)) {
        return;
      }
      if (assistantItem?.kind === "message") {
        assistantItem.isFinished = true;
        assistantItem.content += `\n\n> 生成中断：${error instanceof Error ? error.message : String(error)}`;
        updateItem(input.sessionId, assistantItem);
      } else {
        addItem(input.sessionId, {
          id: createId("msg"),
          kind: "message",
          role: "assistant",
          content: `> 生成中断：${error instanceof Error ? error.message : String(error)}`,
          isFinished: true,
          createdAt: now()
        });
      }
      if (toolItem?.kind === "tool") {
        finishToolCall(input.sessionId, toolItem, "failed", error instanceof Error ? error.message : String(error));
      }
      setSessionStatus(input.sessionId, "failed");
    })
    .finally(() => {
      abortControllers.delete(input.sessionId);
    });

  return { accepted: true };
}

async function pickAttachments(input: PickAttachmentInput): Promise<AttachmentRef[]> {
  const result = await dialog.showOpenDialog({
    properties: input.multiple === false ? ["openFile"] : ["openFile", "multiSelections"],
    filters: [
      { name: "Supported files", extensions: ["docx", "pdf", "md", "txt", "png", "jpg", "jpeg", "webp", "svg", "xlsx"] },
      { name: "All files", extensions: ["*"] }
    ]
  });
  if (result.canceled) return [];

  return result.filePaths.map((filePath) => {
    const attachment: AttachmentRef = {
      id: createId("attachment"),
      sessionId: input.sessionId,
      name: basename(filePath),
      mimeType: "application/octet-stream",
      size: statSync(filePath).size,
      path: filePath,
      source: "picker",
      createdAt: now()
    };
    attachments.set(attachment.id, attachment);
    schedulePersistState();
    return attachment;
  });
}

function importClipboardAttachments(input: ClipboardAttachmentInput): AttachmentRef[] {
  ensureDataDirs();
  return input.files.map((file) => {
    const safeName = file.name.replace(/[<>:"/\\|?*]/g, "_");
    const id = createId("attachment");
    const filePath = join(inputDir, `${id}-${safeName}`);
    writeFileSync(filePath, Buffer.from(file.dataBase64, "base64"));
    const attachment: AttachmentRef = {
      id,
      sessionId: input.sessionId,
      name: safeName,
      mimeType: file.mimeType,
      size: statSync(filePath).size,
      path: filePath,
      source: "clipboard",
      createdAt: now()
    };
    attachments.set(attachment.id, attachment);
    schedulePersistState();
    return attachment;
  });
}

function registerIpc(): void {
  ipcMain.handle("session:list", () => Array.from(sessions.values()));
  ipcMain.handle("session:create", () => createSession());
  ipcMain.handle("session:rename", (_event, input: RenameSessionInput) => renameSession(input));
  ipcMain.handle("session:delete", (_event, sessionId: string) => deleteSession(sessionId));
  ipcMain.handle("chat:prompt", (_event, input: ChatPromptInput) => handlePrompt(input));
  ipcMain.handle("chat:abort", (_event, sessionId: string) => {
    abortControllers.get(sessionId)?.abort();
    abortControllers.delete(sessionId);
    setSessionStatus(sessionId, "idle");
  });
  ipcMain.handle("attachment:pick", (_event, input: PickAttachmentInput) => pickAttachments(input));
  ipcMain.handle("attachment:import-clipboard", (_event, input: ClipboardAttachmentInput) =>
    importClipboardAttachments(input)
  );
  ipcMain.handle("attachment:remove", (_event, attachmentId: string) => {
    attachments.delete(attachmentId);
    schedulePersistState();
  });
  ipcMain.handle("artifact:list", (_event, input?: ArtifactListInput) => listArtifacts(input));
  ipcMain.handle("artifact:open", (_event, artifactId: string) => {
    const artifact = artifacts.get(artifactId);
    if (artifact) void shell.openPath(artifact.path);
  });
  ipcMain.handle("artifact:reveal", (_event, artifactId: string) => {
    const artifact = artifacts.get(artifactId);
    if (artifact) shell.showItemInFolder(artifact.path);
  });
  ipcMain.handle("artifact:preview", (_event, artifactId: string) => {
    const artifact = artifacts.get(artifactId);
    if (!artifact) {
      throw new Error("Artifact not found");
    }
    return buildArtifactPreview(artifact);
  });
  ipcMain.handle("settings:get", () => loadEnv());
  ipcMain.handle("settings:save", (_event, input: UpdateAppSettingsInput) => saveEnv(input));
}

function createWindow(): void {
  mainWindow = new BrowserWindow({
    width: 1360,
    height: 860,
    minWidth: 1180,
    minHeight: 760,
    title: "PwdSafeAgent",
    webPreferences: {
      preload: join(__dirname, "../preload/index.cjs"),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  if (process.env.ELECTRON_RENDERER_URL) {
    void mainWindow.loadURL(process.env.ELECTRON_RENDERER_URL);
  } else {
    void mainWindow.loadFile(join(__dirname, "../renderer/index.html"));
  }
}

app.whenReady().then(() => {
  ensureDataDirs();
  loadEnv();
  restorePersistedState();
  registerIpc();
  if (sessions.size === 0) {
    createSession();
  }
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  flushPersistedState();
  if (process.platform !== "darwin") app.quit();
});

app.on("before-quit", () => {
  flushPersistedState();
});
