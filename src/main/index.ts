import electron from "electron";
import type { BrowserWindow as BrowserWindowType } from "electron";
import { join, extname, basename, resolve } from "node:path";
import { existsSync, mkdirSync, statSync, writeFileSync } from "node:fs";
import type { ChatCompletionMessageParam } from "openai/resources/chat/completions";
import dotenv from "dotenv";
import { compactText, getCurrentTimeText } from "./agentTools";
import { createAgentRuntime, normalizeAgentRuntimeKind, type AgentRuntimeHost, type MessageStreamItem } from "./agentRuntime";
import { prepareImportedAttachments } from "./attachmentImport";
import { resolveAppPaths } from "./appPaths";
import { buildArtifactPreview } from "./artifactPreview";
import { findBundledPythonRuntime, getProcessResourcesDir } from "./bundledRuntime";
import { serializeEnvFile } from "./envFile";
import {
  checkRuntime as checkRuntimeDiagnostics,
  getBundledPythonRuntimeStatus
} from "./runtimeDiagnostics";
import { registerContentSecurityPolicy } from "./securityHeaders";
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

const { app, BrowserWindow, dialog, ipcMain, shell } = electron;
const projectRootDir = process.cwd();
const resourcesDir = getProcessResourcesDir();
const appPaths = resolveAppPaths({
  projectRootDir,
  resourcesDir,
  userDataDir: app.getPath("userData"),
  packaged: app.isPackaged
});
const { rootDir, docsDir, dataDir, inputDir, outputDir, statePath, envLocalPath, envPath } = appPaths;

let mainWindow: BrowserWindowType | null = null;
const sessions = new Map<string, ChatSession>();
const attachments = new Map<string, AttachmentRef>();
const artifacts = new Map<string, ArtifactSummary>();
const abortControllers = new Map<string, AbortController>();
const sessionMemories = new Map<string, SessionMemoryEntry[]>();
let persistTimer: NodeJS.Timeout | undefined;
const MAX_DOCUMENT_CONTEXT_CHARS = 24000;
const MAX_SESSION_MEMORY_CHARS = 90000;

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
    sessionMemories.clear();

    for (const session of state.sessions) sessions.set(session.id, session);
    for (const attachment of state.attachments) attachments.set(attachment.id, attachment);
    for (const artifact of state.artifacts) artifacts.set(artifact.id, artifact);
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
      configSource: existsSync(envLocalPath) ? ".env.local" : existsSync(envPath) ? ".env" : "process",
      bundledPython: getBundledPythonRuntimeStatus({ projectRootDir, resourcesDir })
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
      runtime: normalizeAgentRuntimeKind(process.env.AGENT_RUNTIME),
      execBashEnabled: parseBooleanEnv(process.env.AGENT_EXEC_BASH_ENABLED, true),
      piAgentPackage: process.env.PI_AGENT_PACKAGE || "",
      piAgentExport: process.env.PI_AGENT_EXPORT || ""
    }
  };
}

function saveEnv(input: UpdateAppSettingsInput): AppSettings {
  const currentKey = process.env.OPENAI_API_KEY || "";
  const currentImageKey = process.env.OPENAI_IMAGE_API_KEY || "";
  const content = serializeEnvFile([
    { key: "OPENAI_API_KEY", value: input.openai.apiKey ?? currentKey },
    { key: "OPENAI_BASE_URL", value: input.openai.baseUrl },
    { key: "OPENAI_IMAGE_BASE_URL", value: input.openai.imageBaseUrl },
    { key: "OPENAI_IMAGE_API_KEY", value: input.openai.imageApiKey ?? currentImageKey },
    { key: "OPENAI_CHAT_MODEL", value: input.openai.chatModel },
    { key: "OPENAI_IMAGE_MODEL", value: input.openai.imageModel },
    { key: "OPENAI_IMAGE_SIZE", value: input.openai.imageSize },
    { key: "OPENAI_IMAGE_QUALITY", value: input.openai.imageQuality },
    { key: "OPENAI_AUTO_IMAGE_GENERATION", value: input.openai.autoImageGeneration },
    { key: "OPENAI_REQUEST_TIMEOUT_MS", value: input.openai.requestTimeoutMs },
    { key: "OPENAI_MAX_OUTPUT_TOKENS", value: input.openai.maxOutputTokens },
    { key: "AGENT_RUNTIME", value: input.agent.runtime },
    { key: "AGENT_EXEC_BASH_ENABLED", value: input.agent.execBashEnabled },
    { key: "PI_AGENT_PACKAGE", value: input.agent.piAgentPackage },
    { key: "PI_AGENT_EXPORT", value: input.agent.piAgentExport },
    { key: "AGENT_AUTO_PDF_EXPORT", value: input.document.autoPdfExport },
    { key: "LIBREOFFICE_PATH", value: input.document.libreOfficePath }
  ]);
  writeFileSync(envLocalPath, content, "utf-8");
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

function addStage(sessionId: string, title: string, detail?: string): void {
  addItem(sessionId, {
    id: createId("stage"),
    kind: "stage",
    title,
    detail,
    createdAt: now()
  });
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
  const existing = findSessionArtifactByPath(sessionId, filePath);
  if (existing) {
    return existing;
  }

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

function findSessionArtifactByPath(sessionId: string, filePath: string): ArtifactSummary | undefined {
  const normalizedPath = resolve(filePath).toLowerCase();
  return Array.from(artifacts.values()).find(
    (artifact) => resolve(artifact.path).toLowerCase() === normalizedPath && artifactBelongsToSession(artifact, sessionId)
  );
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
    "你的核心工作流是：先通过多轮对话收集项目事实，持续总结项目档案；中途按需读取附件、检索资料或做中间分析；当信息基本充分或用户明确要求交付时，再生成最终方案文件。",
    "不要在第一轮或资料明显不足时直接生成 Word、PDF、图片或发送文件；此时应先总结已知信息、指出缺口，并继续澄清关键事实。",
    "每当用户补充了项目关键信息，优先调用 remember_project 沉淀已确认事实和待补充信息。项目档案应覆盖：应用系统、建设单位、单位省份、单位地址、邮编、等保级别、系统边界、业务场景、部署架构、应用子系统、关键数据、用户角色、密码产品、机房/云平台、外部接口和交付要求。",
    "只有当用户明确说“生成/导出/输出/形成方案/出 Word/PDF/画图”等交付意图，或项目档案已经标记为生成就绪时，才调用 write_word、write_pdf、image_generate、send_file 等最终交付工具。",
    "最终生成时需要严格参考用户给出的 Word 模板结构，生成专业、可复核、可落地的密码应用方案内容。",
    "方案需要覆盖系统概况、密码应用需求、密码应用设计、密钥管理、实施计划、风险与符合性说明等章节。",
    "回复使用中文 Markdown，必要时给出缺失资料清单。",
    "不要编造用户未提供的关键事实；若资料不足，用“待补充/需确认”标识。",
    "工具调用由你按任务需要自主决策：寒暄、普通问答和资料澄清阶段不要默认读取模板或附件；只有分析附件、查询最新资料、沉淀项目档案或交付文件确有需要时，才调用对应工具。",
    "当进入最终交付阶段并需要 Word、PDF、Markdown 或图片文件时，必须通过 write_word/write_pdf/write_file/image_generate 等工具真实生成文件；不要只在文本回复中声称已经生成。",
    "最终文件生成后，如果工具结果没有自动展示文件卡片，再调用 send_file 将 data/output 中的产物发送给前端。",
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
      content: `以下是本会话已沉淀的项目档案、模板/附件上下文和工具结果。继续对话时优先基于这些信息更新项目档案；最终生成方案时必须优先参考：\n\n${memory}`
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

function buildUserMessageContent(input: ChatPromptInput): string {
  const message = input.message.trim();
  const attachmentNames = input.attachments?.map((attachment) => attachment.name).filter(Boolean) ?? [];

  if (!attachmentNames.length) return message;

  const attachmentText = `附件：${attachmentNames.join("、")}`;
  if (!message) return `已上传 ${attachmentNames.length} 个附件。\n${attachmentText}`;
  return `${message}\n\n${attachmentText}`;
}

async function streamMockResponse(sessionId: string, assistantItem: MessageStreamItem): Promise<void> {
  const memory = sessionMemories.get(sessionId) ?? [];
  const hasContext = memory.length > 0;
  const chunks = [
    "已收到你的需求。当前不会自动读取模板或附件，我会在生成方案、分析资料或导出文件时按需调用工具。\n\n",
    hasContext ? `当前会话已有 ${memory.length} 份工具读取上下文。\n\n` : "当前还没有工具读取的模板或附件上下文。\n\n",
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

function createAgentRuntimeHost(): AgentRuntimeHost {
  return {
    rootDir,
    docsDir,
    inputDir,
    outputDir,
    loadSettings: loadEnv,
    getSession: (sessionId) => sessions.get(sessionId),
    buildMessages,
    createAssistantMessage,
    updateItem,
    startToolCall,
    finishToolCall,
    addStage,
    appendSessionMemory,
    formatSessionMemory,
    getSessionReadableFiles,
    getBundledPythonRuntime: () => findBundledPythonRuntime({ rootDir: projectRootDir, resourcesDir }),
    createArtifact: (sessionId, filePath) => {
      createArtifact(sessionId, filePath);
    },
    streamMockResponse
  };
}

async function runAgentResponse(
  sessionId: string,
  controller: AbortController,
  userPrompt: string,
  onAssistantCreated: (assistantItem: MessageStreamItem) => void
): Promise<MessageStreamItem> {
  const runtime = createAgentRuntime(loadEnv().agent.runtime, createAgentRuntimeHost());
  return runtime.runTurn({
    sessionId,
    userPrompt,
    controller,
    onAssistantCreated
  });
}

async function handlePrompt(input: ChatPromptInput): Promise<{ accepted: true }> {
  const session = sessions.get(input.sessionId);
  if (!session) {
    throw new Error("Session not found");
  }
  const userContent = buildUserMessageContent(input);

  if (session.title === "新的密码方案对话" && userContent) {
    session.title = userContent.replace(/\s+/g, " ").slice(0, 24);
  }

  addItem(input.sessionId, {
    id: createId("msg"),
    kind: "message",
    role: "user",
    content: userContent,
    isFinished: true,
    createdAt: now(),
    attachmentIds: input.attachments?.map((attachment) => attachment.id)
  });

  setSessionStatus(input.sessionId, "running");
  const controller = new AbortController();
  abortControllers.set(input.sessionId, controller);
  let assistantItem: MessageStreamItem | undefined;

  void Promise.resolve()
    .then(async () => {
      assistantItem = await runAgentResponse(input.sessionId, controller, input.message, (createdItem) => {
        assistantItem = createdItem;
      });
      assistantItem.isFinished = true;
      updateItem(input.sessionId, assistantItem);
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
  return prepareImportedAttachments(input, {
    inputDir,
    createId: () => createId("attachment")
  }).map((file) => {
    writeFileSync(file.outputPath, file.buffer);
    const attachment: AttachmentRef = {
      id: file.id,
      sessionId: input.sessionId,
      name: file.safeName,
      mimeType: file.mimeType,
      size: statSync(file.outputPath).size,
      path: file.outputPath,
      source: file.source,
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
  ipcMain.handle("artifact:open", async (_event, artifactId: string) => {
    const artifact = artifacts.get(artifactId);
    if (!artifact) {
      throw new Error("Artifact not found");
    }
    const errorMessage = await shell.openPath(artifact.path);
    if (errorMessage) {
      throw new Error(errorMessage);
    }
  });
  ipcMain.handle("artifact:reveal", (_event, artifactId: string) => {
    const artifact = artifacts.get(artifactId);
    if (!artifact) {
      throw new Error("Artifact not found");
    }
    shell.showItemInFolder(artifact.path);
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
  ipcMain.handle("settings:check-runtime", () =>
    checkRuntimeDiagnostics({
      projectRootDir,
      resourcesDir,
      cwd: rootDir,
      env: process.env,
      now
    })
  );
}

function resolveWindowIconPath(): string | undefined {
  const candidates = [
    resourcesDir ? join(resourcesDir, "build", "icon.png") : "",
    resourcesDir ? join(resourcesDir, "build", "icon.ico") : "",
    join(projectRootDir, "build", "icon.png"),
    join(projectRootDir, "build", "icon.ico")
  ];
  return candidates.find((candidate) => Boolean(candidate) && existsSync(candidate));
}

function createWindow(): void {
  const windowIconPath = resolveWindowIconPath();
  mainWindow = new BrowserWindow({
    width: 1360,
    height: 860,
    minWidth: 1180,
    minHeight: 760,
    title: "PwdSafeAgent",
    autoHideMenuBar: true,
    ...(windowIconPath ? { icon: windowIconPath } : {}),
    webPreferences: {
      preload: join(__dirname, "../preload/index.cjs"),
      contextIsolation: true,
      nodeIntegration: false
    }
  });
  mainWindow.setMenuBarVisibility(false);
  mainWindow.removeMenu();

  registerContentSecurityPolicy(mainWindow.webContents.session, {
    dev: Boolean(process.env.ELECTRON_RENDERER_URL)
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
