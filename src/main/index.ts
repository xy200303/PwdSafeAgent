import electron from "electron";
import type { BrowserWindow as BrowserWindowType } from "electron";
import { join, extname, basename, resolve } from "node:path";
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import type { ChatCompletionMessageParam } from "openai/resources/chat/completions";
import type { ImageContent } from "@mariozechner/pi-ai";
import dotenv from "dotenv";
import {
  MAX_VISION_IMAGE_BYTES,
  compactText,
  getCurrentTimeText,
  isSupportedImageFile,
  readImageDataUrl
} from "./agentTools";
import { createAgentRuntime, type AgentRuntimeHost, type MessageStreamItem } from "./agentRuntime";
import { prepareImportedAttachments, sanitizeAttachmentName } from "./attachmentImport";
import { resolveAppPaths } from "./appPaths";
import { buildArtifactPreview } from "./artifactPreview";
import { findBundledPythonRuntime, getProcessResourcesDir } from "./bundledRuntime";
import { serializeEnvFile } from "./envFile";
import {
  checkRuntime as checkRuntimeDiagnostics,
  getBundledPythonRuntimeStatus
} from "./runtimeDiagnostics";
import { registerContentSecurityPolicy } from "./securityHeaders";
import {
  applySchemeProgressUpdate,
  createSchemeProgressItem,
  settleSchemeProgressItem,
  type SchemeProgressUpdateInput
} from "./schemeProgress";
import { stripSyntheticSchemeCompletionNotice } from "./schemeContinuation";
import { loadPersistedState, savePersistedState, type PersistedStateSnapshot, type SessionMemoryEntry } from "./sessionPersistence";
import { resolveSessionWorkspaceDirs, type SessionWorkspaceDirs } from "./sessionWorkspace";
import { registerDocumentTemplateFromDocx, type RegisteredDocumentTemplate } from "./documentTemplateRegistration";
import {
  BUILT_IN_STANDARD_REFERENCE_RELATIVE_PATH,
  BUILT_IN_TEMPLATE_DOCX_RELATIVE_PATH,
  BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH
} from "./templatePaths";
import {
  IMAGE_GENERATION_REQUEST_TIMEOUT_DEFAULT_MS,
  clampDraftSectionParallelism,
  clampImageGenerationParallelism
} from "../shared/types";
import { APP_DISPLAY_NAME, createAppMetadata } from "../shared/appMetadata";
import type {
  AppMetadata,
  AppSettings,
  ArtifactKind,
  ArtifactListInput,
  ArtifactSummary,
  AttachmentRef,
  ChatPromptInput,
  ChatSession,
  ClipboardAttachmentInput,
  DocumentTemplateSummary,
  PickAttachmentInput,
  RenameSessionInput,
  RendererEvent,
  SchemeProgressItem,
  SchemeProgressSection,
  StreamItem,
  UpdateAppSettingsInput
} from "../shared/types";

const { app, BrowserWindow, dialog, ipcMain, shell } = electron;
app.setName(APP_DISPLAY_NAME);
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
const pendingStreamItemUpdateEvents = new Map<string, RendererEvent>();
const pendingStreamItemUpdateTimers = new Map<string, NodeJS.Timeout>();
let persistTimer: NodeJS.Timeout | undefined;
let selectedDocumentTemplateId = "default";
const MAX_DOCUMENT_CONTEXT_CHARS = 24000;
const MAX_SESSION_MEMORY_CHARS = 90000;
const STREAM_ITEM_UPDATE_THROTTLE_MS = 80;
const DEFAULT_DOCUMENT_TEMPLATE_ID = "default";
const DEFAULT_DOCUMENT_TEMPLATE_PROFILE_PATH = "docs/document-profiles/generic_document.json";

function getAppMetadata(): AppMetadata {
  return createAppMetadata(app.getVersion());
}

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

function getSessionWorkspaceDirs(sessionId: string): SessionWorkspaceDirs {
  return resolveSessionWorkspaceDirs(dataDir, sessionId);
}

function ensureSessionWorkspaceDirs(sessionId: string): SessionWorkspaceDirs {
  const dirs = getSessionWorkspaceDirs(sessionId);
  for (const dir of [dirs.workspaceDir, dirs.inputDir, dirs.outputDir]) {
    if (!existsSync(dir)) {
      mkdirSync(dir, { recursive: true });
    }
  }
  return dirs;
}

function getDefaultDocumentTemplate(): DocumentTemplateSummary {
  return {
    id: DEFAULT_DOCUMENT_TEMPLATE_ID,
    name: "默认模板",
    source: "builtin",
    profilePath: DEFAULT_DOCUMENT_TEMPLATE_PROFILE_PATH,
    templatePath: BUILT_IN_TEMPLATE_DOCX_RELATIVE_PATH,
    templateJsonPath: BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH,
    renderMode: "template_sections"
  };
}

function getDocumentTemplatesDir(): string {
  return join(outputDir, "document-templates");
}

function listDocumentTemplates(): DocumentTemplateSummary[] {
  ensureDataDirs();
  const templatesDir = getDocumentTemplatesDir();
  const uploadedTemplates = existsSync(templatesDir)
    ? readdirSync(templatesDir, { withFileTypes: true })
        .filter((entry) => entry.isDirectory())
        .filter((entry) => entry.name !== DEFAULT_DOCUMENT_TEMPLATE_ID)
        .map((entry) => readUploadedDocumentTemplateSummary(entry.name))
        .filter((template): template is DocumentTemplateSummary => Boolean(template))
        .sort(compareUploadedDocumentTemplates)
    : [];

  return [getDefaultDocumentTemplate(), ...uploadedTemplates];
}

function getSelectedDocumentTemplate(): DocumentTemplateSummary {
  const templates = listDocumentTemplates();
  const selected = templates.find((template) => template.id === selectedDocumentTemplateId);
  if (selected) return selected;
  selectedDocumentTemplateId = DEFAULT_DOCUMENT_TEMPLATE_ID;
  return templates.find((template) => template.id === DEFAULT_DOCUMENT_TEMPLATE_ID) ?? getDefaultDocumentTemplate();
}

function selectDocumentTemplate(templateId: string): DocumentTemplateSummary {
  const normalizedId = typeof templateId === "string" ? templateId.trim() : "";
  const selected = listDocumentTemplates().find((template) => template.id === normalizedId);
  if (!selected) {
    throw new Error(`文档模板不存在：${normalizedId || "(空)"}`);
  }
  selectedDocumentTemplateId = selected.id;
  schedulePersistState();
  return selected;
}

async function uploadDocumentTemplate(): Promise<DocumentTemplateSummary | undefined> {
  ensureDataDirs();
  const result = await dialog.showOpenDialog({
    properties: ["openFile"],
    filters: [{ name: "Word templates (.docx)", extensions: ["docx"] }]
  });
  if (result.canceled) return undefined;

  const sourcePath = result.filePaths[0];
  if (!sourcePath) return undefined;
  if (extname(sourcePath).toLowerCase() !== ".docx") {
    throw new Error(`上传模板只支持 .docx 文件：${basename(sourcePath)}`);
  }

  const registered = await registerDocumentTemplateFromDocx({
    sourceDocxPath: sourcePath,
    outputDir
  });
  const summary = createUploadedDocumentTemplateSummary(registered);
  selectedDocumentTemplateId = summary.id;
  schedulePersistState();
  return summary;
}

function createUploadedDocumentTemplateSummary(template: RegisteredDocumentTemplate): DocumentTemplateSummary {
  const createdAt = existsSync(template.profilePath) ? statSync(template.profilePath).mtime.toISOString() : now();
  return {
    id: template.templateId,
    name: template.displayName,
    source: "uploaded",
    profilePath: template.profilePathForTool,
    templatePath: template.templatePathForTool,
    templateJsonPath: template.templateJsonPathForTool,
    renderMode: template.renderMode,
    sectionCount: template.sectionCount,
    sectionGroupCount: template.sectionGroupCount,
    createdAt
  };
}

function readUploadedDocumentTemplateSummary(templateId: string): DocumentTemplateSummary | undefined {
  const templateDir = join(getDocumentTemplatesDir(), templateId);
  const profilePath = join(templateDir, "profile.json");
  const templatePath = join(templateDir, "template.docx");
  const templateJsonPath = join(templateDir, "template.json");
  if (!existsSync(profilePath) || !existsSync(templatePath)) return undefined;

  const profile = readJsonObject(profilePath);
  const templateJson = existsSync(templateJsonPath) ? readJsonObject(templateJsonPath) : undefined;
  const templateStrategy = readRecord(templateJson?.strategy);
  return {
    id: templateId,
    name: readString(profile?.titleSuffix) || readString(profile?.profile) || templateId,
    source: "uploaded",
    profilePath: `document-templates/${templateId}/profile.json`,
    templatePath: `document-templates/${templateId}/template.docx`,
    templateJsonPath: existsSync(templateJsonPath) ? `document-templates/${templateId}/template.json` : undefined,
    renderMode: readRenderMode(templateStrategy?.defaultWriteMode),
    sectionCount: readNumber(readRecord(templateJson?.statistics)?.sectionCount) ?? readArray(templateJson?.sections)?.length,
    sectionGroupCount: readArray(profile?.sectionRules)?.length,
    createdAt: statSync(profilePath).mtime.toISOString()
  };
}

function compareUploadedDocumentTemplates(left: DocumentTemplateSummary, right: DocumentTemplateSummary): number {
  const leftTime = left.createdAt ? Date.parse(left.createdAt) : 0;
  const rightTime = right.createdAt ? Date.parse(right.createdAt) : 0;
  if (rightTime !== leftTime) return rightTime - leftTime;
  return left.name.localeCompare(right.name, "zh-Hans-CN");
}

function readJsonObject(filePath: string): Record<string, unknown> | undefined {
  try {
    return readRecord(JSON.parse(readFileSync(filePath, "utf-8")));
  } catch {
    return undefined;
  }
}

function readRecord(value: unknown): Record<string, unknown> | undefined {
  return typeof value === "object" && value !== null && !Array.isArray(value) ? (value as Record<string, unknown>) : undefined;
}

function readArray(value: unknown): unknown[] | undefined {
  return Array.isArray(value) ? value : undefined;
}

function readString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function readNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function readRenderMode(value: unknown): DocumentTemplateSummary["renderMode"] | undefined {
  return value === "full_document" || value === "template_sections" ? value : undefined;
}

function parseBooleanEnv(value: string | undefined, fallback: boolean): boolean {
  if (value === undefined) return fallback;
  return /^(1|true|yes|on)$/i.test(value.trim());
}

function sendEvent(event: RendererEvent): void {
  mainWindow?.webContents.send("app:event", event);
}

function sendStreamItemUpdatedEvent(sessionId: string, item: StreamItem): void {
  const event: RendererEvent = { id: createId("event"), type: "stream.item.updated", sessionId, payload: item };
  const shouldThrottle = item.kind === "message" && item.role === "assistant" && !item.isFinished;
  if (!shouldThrottle) {
    flushPendingStreamItemUpdate(sessionId, item.id);
    sendEvent(event);
    return;
  }

  const key = getStreamItemUpdateKey(sessionId, item.id);
  pendingStreamItemUpdateEvents.set(key, event);
  if (pendingStreamItemUpdateTimers.has(key)) return;
  pendingStreamItemUpdateTimers.set(
    key,
    setTimeout(() => {
      flushPendingStreamItemUpdate(sessionId, item.id);
    }, STREAM_ITEM_UPDATE_THROTTLE_MS)
  );
}

function flushPendingStreamItemUpdate(sessionId: string, itemId: string): void {
  const key = getStreamItemUpdateKey(sessionId, itemId);
  const timer = pendingStreamItemUpdateTimers.get(key);
  if (timer) {
    clearTimeout(timer);
    pendingStreamItemUpdateTimers.delete(key);
  }
  const event = pendingStreamItemUpdateEvents.get(key);
  if (!event) return;
  pendingStreamItemUpdateEvents.delete(key);
  sendEvent(event);
}

function getStreamItemUpdateKey(sessionId: string, itemId: string): string {
  return `${sessionId}:${itemId}`;
}

function clearPendingStreamItemUpdatesForSession(sessionId: string): void {
  const prefix = `${sessionId}:`;
  for (const key of Array.from(pendingStreamItemUpdateEvents.keys())) {
    if (!key.startsWith(prefix)) continue;
    const timer = pendingStreamItemUpdateTimers.get(key);
    if (timer) clearTimeout(timer);
    pendingStreamItemUpdateTimers.delete(key);
    pendingStreamItemUpdateEvents.delete(key);
  }
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
    sessionMemories: memoryRecord,
    selectedDocumentTemplateId
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
    selectedDocumentTemplateId = state.selectedDocumentTemplateId || DEFAULT_DOCUMENT_TEMPLATE_ID;
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
      visionBaseUrl: process.env.OPENAI_VISION_BASE_URL || "",
      chatModel: process.env.OPENAI_CHAT_MODEL || "gpt-5.5",
      chatImageInputEnabled: parseBooleanEnv(process.env.OPENAI_CHAT_IMAGE_INPUT_ENABLED, false),
      imageModel: process.env.OPENAI_IMAGE_MODEL || "gpt-image-2",
      visionModel: process.env.OPENAI_VISION_MODEL || "",
      imageSize: process.env.OPENAI_IMAGE_SIZE || "1536x1024",
      imageQuality: process.env.OPENAI_IMAGE_QUALITY || "high",
      autoImageGeneration: parseBooleanEnv(process.env.OPENAI_AUTO_IMAGE_GENERATION, true),
      thinkingEnabled: parseBooleanEnv(process.env.OPENAI_THINKING_ENABLED, true),
      reasoningEffort: process.env.OPENAI_REASONING_EFFORT || "",
      requestTimeoutMs: Number(process.env.OPENAI_REQUEST_TIMEOUT_MS || 120000),
      imageRequestTimeoutMs: Number(
        process.env.OPENAI_IMAGE_REQUEST_TIMEOUT_MS || IMAGE_GENERATION_REQUEST_TIMEOUT_DEFAULT_MS
      ),
      maxOutputTokens: Number(process.env.OPENAI_MAX_OUTPUT_TOKENS || 16000),
      apiKeyConfigured: Boolean(process.env.OPENAI_API_KEY),
      imageApiKeyConfigured: Boolean(process.env.OPENAI_IMAGE_API_KEY),
      visionApiKeyConfigured: Boolean(process.env.OPENAI_VISION_API_KEY)
    },
    document: {
      autoPdfExport: parseBooleanEnv(process.env.AGENT_AUTO_PDF_EXPORT, false),
      libreOfficePath: process.env.LIBREOFFICE_PATH || ""
    },
    agent: {
      execBashEnabled: parseBooleanEnv(process.env.AGENT_EXEC_BASH_ENABLED, true),
      draftSectionParallelism: clampDraftSectionParallelism(process.env.AGENT_DRAFT_SECTION_PARALLELISM),
      imageGenerationParallelism: clampImageGenerationParallelism(process.env.AGENT_IMAGE_GENERATION_PARALLELISM)
    }
  };
}

function saveEnv(input: UpdateAppSettingsInput): AppSettings {
  const currentKey = process.env.OPENAI_API_KEY || "";
  const currentImageKey = process.env.OPENAI_IMAGE_API_KEY || "";
  const currentVisionKey = process.env.OPENAI_VISION_API_KEY || "";
  const content = serializeEnvFile([
    { key: "OPENAI_API_KEY", value: input.openai.apiKey ?? currentKey },
    { key: "OPENAI_BASE_URL", value: input.openai.baseUrl },
    { key: "OPENAI_IMAGE_BASE_URL", value: input.openai.imageBaseUrl },
    { key: "OPENAI_VISION_BASE_URL", value: input.openai.visionBaseUrl },
    { key: "OPENAI_IMAGE_API_KEY", value: input.openai.imageApiKey ?? currentImageKey },
    { key: "OPENAI_VISION_API_KEY", value: input.openai.visionApiKey ?? currentVisionKey },
    { key: "OPENAI_CHAT_MODEL", value: input.openai.chatModel },
    { key: "OPENAI_CHAT_IMAGE_INPUT_ENABLED", value: input.openai.chatImageInputEnabled },
    { key: "OPENAI_IMAGE_MODEL", value: input.openai.imageModel },
    { key: "OPENAI_VISION_MODEL", value: input.openai.visionModel },
    { key: "OPENAI_IMAGE_SIZE", value: input.openai.imageSize },
    { key: "OPENAI_IMAGE_QUALITY", value: input.openai.imageQuality },
    { key: "OPENAI_AUTO_IMAGE_GENERATION", value: input.openai.autoImageGeneration },
    { key: "OPENAI_THINKING_ENABLED", value: input.openai.thinkingEnabled },
    { key: "OPENAI_REASONING_EFFORT", value: input.openai.reasoningEffort },
    { key: "OPENAI_REQUEST_TIMEOUT_MS", value: input.openai.requestTimeoutMs },
    { key: "OPENAI_IMAGE_REQUEST_TIMEOUT_MS", value: input.openai.imageRequestTimeoutMs },
    { key: "OPENAI_MAX_OUTPUT_TOKENS", value: input.openai.maxOutputTokens },
    { key: "AGENT_EXEC_BASH_ENABLED", value: input.agent.execBashEnabled },
    { key: "AGENT_DRAFT_SECTION_PARALLELISM", value: clampDraftSectionParallelism(input.agent.draftSectionParallelism) },
    { key: "AGENT_IMAGE_GENERATION_PARALLELISM", value: clampImageGenerationParallelism(input.agent.imageGenerationParallelism) },
    { key: "AGENT_AUTO_PDF_EXPORT", value: input.document.autoPdfExport },
    { key: "LIBREOFFICE_PATH", value: input.document.libreOfficePath }
  ]);
  writeFileSync(envLocalPath, content, "utf-8");
  dotenv.config({ path: envLocalPath, override: true });
  return loadEnv();
}

function createSession(): ChatSession {
  const sessionId = createId("session");
  ensureSessionWorkspaceDirs(sessionId);
  const session: ChatSession = {
    id: sessionId,
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
  clearPendingStreamItemUpdatesForSession(sessionId);
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
  sendStreamItemUpdatedEvent(sessionId, item);
}

function setSessionStatus(sessionId: string, status: ChatSession["status"]): void {
  const session = sessions.get(sessionId);
  if (!session) return;
  session.status = status;
  session.updatedAt = now();
  sessions.set(sessionId, session);
  schedulePersistState();
  sendEvent({ id: createId("event"), type: "session.updated", payload: session });

  if (status === "completed" || status === "failed") {
    settleSchemeProgress(sessionId, status);
  }
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

function getLatestSchemeProgressItem(sessionId: string): SchemeProgressItem | undefined {
  const session = sessions.get(sessionId);
  if (!session) return undefined;
  for (let index = session.items.length - 1; index >= 0; index -= 1) {
    const item = session.items[index];
    if (item.kind === "scheme_progress") return item;
  }
  return undefined;
}

function ensureSchemeProgress(
  sessionId: string,
  detail?: string,
  artifactName?: string,
  options: { title?: string; sections?: SchemeProgressSection[]; reset?: boolean } = {}
): void {
  const existing = options.reset ? undefined : getLatestSchemeProgressItem(sessionId);
  const timestamp = now();
  if (!existing) {
    const item = createSchemeProgressItem({
      id: createId("scheme_progress"),
      docsDir,
      title: options.title,
      sections: options.sections,
      createdAt: timestamp
    });
    addItem(sessionId, {
      ...item,
      status: "running",
      detail,
      artifactName,
      updatedAt: timestamp
    });
    return;
  }

  if (!detail && !artifactName && existing.status === "running") return;
  updateItem(sessionId, {
    ...existing,
    status: "running",
    title: options.title ?? existing.title,
    detail: detail ?? existing.detail,
    artifactName: artifactName ?? existing.artifactName,
    sections: options.sections ?? existing.sections,
    total: options.sections?.length ?? existing.total,
    updatedAt: timestamp
  });
}

function updateSchemeSectionProgress(sessionId: string, update: SchemeProgressUpdateInput): void {
  ensureSchemeProgress(sessionId, update.detail, update.artifactName);
  const existing = getLatestSchemeProgressItem(sessionId);
  if (!existing) return;
  updateItem(sessionId, applySchemeProgressUpdate(existing, update, now()));
}

function settleSchemeProgress(sessionId: string, status: "completed" | "failed"): void {
  const existing = getLatestSchemeProgressItem(sessionId);
  if (!existing) return;
  let nextStatus: SchemeProgressItem["status"] = status;
  if (status === "completed") {
    if (existing.failed > 0) {
      nextStatus = "failed";
    } else if (existing.total > 0 && existing.completed === existing.total) {
      nextStatus = "completed";
    } else if (existing.completed > 0 || (existing.drafted ?? 0) > 0) {
      nextStatus = "partial";
    } else {
      nextStatus = "pending";
    }
  }
  updateItem(sessionId, settleSchemeProgressItem(existing, nextStatus, now()));
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
    const refreshed: ArtifactSummary = {
      ...existing,
      sessionId,
      name,
      size: existsSync(filePath) ? statSync(filePath).size : existing.size,
      createdAt: now()
    };
    artifacts.set(refreshed.id, refreshed);
    ensureArtifactStreamItem(sessionId, refreshed, { refreshExisting: true });
    sendEvent({ id: createId("event"), type: "artifact.created", sessionId, payload: refreshed });
    schedulePersistState();
    return refreshed;
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
  ensureArtifactStreamItem(sessionId, artifact);
  sendEvent({ id: createId("event"), type: "artifact.created", sessionId, payload: artifact });
  schedulePersistState();
  return artifact;
}

function ensureArtifactStreamItem(
  sessionId: string,
  artifact: ArtifactSummary,
  options: { refreshExisting?: boolean } = {}
): void {
  const session = sessions.get(sessionId);
  if (!session) {
    throw new Error(`Session not found: ${sessionId}`);
  }
  const existingIndex = session.items.findIndex((item) => item.kind === "file" && item.artifactId === artifact.id);
  if (existingIndex >= 0 && !options.refreshExisting) {
    return;
  }
  if (existingIndex >= 0) {
    const existingItem = session.items[existingIndex];
    if (existingItem.kind !== "file") return;
    const fileItem: StreamItem = {
      ...existingItem,
      name: artifact.name,
      fileKind: artifact.kind,
      createdAt: artifact.createdAt
    };
    session.items.splice(existingIndex, 1);
    session.items.push(fileItem);
    session.updatedAt = now();
    sessions.set(sessionId, session);
    schedulePersistState();
    sendStreamItemUpdatedEvent(sessionId, fileItem);
    return;
  }
  addItem(sessionId, {
    id: createId("file"),
    kind: "file",
    artifactId: artifact.id,
    name: artifact.name,
    fileKind: artifact.kind,
    createdAt: now()
  });
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
    "事实证据规则：只有用户明确提供、附件/工具明确读到、或 remember_project 已记录为“已确认事实”的内容，才能在正文中写成确定事实；标准条文、模板提示、示例文档、行业常识和你的推断只能作为要求或建议，不能写成项目已经具备的现状。",
    "未知信息处理规则：凡是建设单位、系统名称、等保级别、设备型号、密码产品、算法、网络区域、接口、机房、云平台、数量、地址、责任主体、部署位置或调用路径没有证据，必须写“待补充/需确认”，不得用看似合理的内容补全。",
    "项目档案中的“待补充信息”是禁止转写为事实的缺口清单；如果后续没有新证据，不得在正文、摘要或最终回复中声称这些缺口已经具备或已经完成。",
    "只有当用户明确说“生成/导出/输出/形成方案/出 Word/PDF/画图”等交付意图，或项目档案已经标记为生成就绪时，才调用 write_document_word、write_pdf、image_generate、send_file 等交付工具。",
    "Word 交付统一使用 write_document_word；模板锚点由工具优先从 Word 模板 .docx 动态解析，模板结构 JSON 是模板索引文件，不得把缺少 JSON 当成无法写 Word。",
    `密码应用要求的标准依据优先参考 ${BUILT_IN_STANDARD_REFERENCE_RELATIVE_PATH}；涉及等级要求、密码技术框架、身份鉴别、密钥管理、机密性/完整性/不可否认性和管理要求时，应以该参考资料约束方案表述。`,
    "模板来源只来自前端当前选中的 Word 文档模板；用户要新增模板时通过界面的“上传 Word 模板”入口完成，Agent 不注册聊天附件为模板。",
    "当前选中模板会以系统上下文单独提供；不得要求用户在正文里填写 profile_path/template_path/template_json_path。必须使用系统上下文中的 profile_path 调用 build_document_config。template_json_path 只用于 write_document_word 的模板结构 JSON，禁止把 template_json_path 当作 profile_path。",
    "如需按用户偏好微调 profile，只有 profile_path 以 document-templates/ 开头时才调用 update_document_profile；docs/... 的内置 profile 只读取使用，不要更新。",
    "整篇交付必须优先按 section-first 通用流程推进：先根据输入材料、项目档案和当前模板完成整体规划，并通过 build_document_config 的 generation_plan/section_group_plans 写入 document-config.json；再调用 list_document_sections 生成 data/output/document-sections/manifest.json；正文起草调用 draft_document_sections 按 section id 并行生成章节草稿，用户或 AI 后续微调用 update_document_section_draft 原地保存到 document-sections/<section>/draft.md；合并前调用 audit_document_sections 做章节级证据审查，存在风险时先调用 revise_document_sections_evidence 原地严格降级并复审；证据风险收口后可调用 polish_document_sections 做章节级语言优化，但不得新增事实或改变待补充/需确认；随后调用 assemble_document_sections 按模板章节顺序合并终稿 Markdown；写 Word 前再调用 audit_document_evidence 做终稿复核；如果审查存在风险，先调用 revise_document_evidence 严格降级为待补充/需确认并再次审查；需要 Word 交付时调用 plan_document_assets 补齐图表任务，再调用 write_document_word 读取终稿 Markdown 并按 document-config.wordTemplate 生成 docx。不要一次性让模型输出整篇长文。",
    "list_document_sections 用于建立章节索引；get_document_section 用于查询某节的模板属性、草稿、表格和图片上下文；draft_document_sections 用于并行生成章节草稿；update_document_section_draft 用于章节级原地微调；audit_document_sections 用于章节合并前逐节扫描疑似幻觉；revise_document_sections_evidence 用于把章节草稿中的风险确定表述原地降级为待补充/需确认；polish_document_sections 用于在不新增事实的前提下优化章节语言；assemble_document_sections 用于确定性合并章节终稿；audit_document_evidence 用于 Word 交付前扫描终稿疑似幻觉；revise_document_evidence 用于把终稿风险确定表述降级为待补充/需确认；write_document_word 用于把合并后的 Markdown 终稿按 profile/模板渲染成 Word。",
    "降低 AI 味：每一节只写与该节相关的项目事实、现状、风险、控制措施、算法/产品/部署位置/调用路径；避免泛泛而谈、套话开头、重复政策背景和空洞排比。没有事实就明确写“待补充/需确认”，不要把标准要求改写成项目现状。",
    "需要方案图示时由 image_generate 真实生成；生成后的图片文件作为本轮输出资源保留，并在正文中以待插图说明或 Markdown 图片引用承接，不要编造图片已经嵌入 Word 的结果。",
    "回复使用中文 Markdown，必要时给出缺失资料清单。",
    "不要编造用户未提供的关键事实；资料不足时可以说明“待补充/需确认”，并在后续获得信息后继续替换对应章节。不得说“已部署、已建设、已接入、已采用、已配置、已完成”等确定措辞，除非上下文中有明确证据。",
    "工具调用由你按任务需要自主决策：寒暄、普通问答和资料澄清阶段不要默认读取模板或附件；只有分析附件、查询最新资料、沉淀项目档案或交付文件确有需要时，才调用对应工具。",
    "当用户上传截图/图片并让你识别问题、指出修改点或根据截图调整方案时，如果当前 Chat 模型已配置支持图像输入，图片会随本轮用户消息直接提供给你；否则必须先调用 read_image 读取图片内容。不要只凭附件名或用户描述猜测截图里有什么。",
    "当进入交付阶段并需要 Word、PDF、Markdown 或图片文件时，必须通过 write_document_word/write_pdf/write_file/image_generate 等工具真实生成文件；生成 Word 前如果已有 Markdown 终稿，先调用 audit_document_evidence，发现高风险时先调用 revise_document_evidence 或向用户确认，不得声称最终版本完成。不要只在文本回复中声称已经生成。",
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

  messages.push({
    role: "system",
    content: renderSelectedDocumentTemplateContext()
  });

  const memory = formatSessionMemory(session.id);
  if (memory) {
    messages.push({
      role: "system",
      content: `以下是本会话已沉淀的项目档案、模板/附件上下文和工具结果。继续对话时优先基于这些信息更新项目档案；最终生成方案时必须优先参考。注意：只有“已确认事实”可以写成确定事实；“待补充信息”、模板提示、标准要求和工具规划任务都不能写成项目已经具备的现状。\n\n${memory}`
    });
  }

  const schemeProgressContext = renderSchemeProgressRuntimeContext(session.id);
  if (schemeProgressContext) {
    messages.push({
      role: "system",
      content: schemeProgressContext
    });
  }

  for (const item of session.items) {
    if (item.kind !== "message") continue;
    const content =
      item.role === "assistant" ? stripSyntheticSchemeCompletionNotice(item.content) : item.content.trim();
    if (!content) continue;
    messages.push({
      role: item.role === "user" ? "user" : "assistant",
      content
    });
  }
  return messages;
}

function buildAvailableResourceContext(session: ChatSession): string {
  const lines = [
    "以下是本会话可按需读取的文件资源。当前选中 Word 模板见单独系统上下文；这些文件尚未读取，只有在用户任务需要时才调用 read_word/read_pdf/read_file/read_image。",
    `内置默认 Word 模板：${BUILT_IN_TEMPLATE_DOCX_RELATIVE_PATH}`,
    `内置默认 Word 模板结构 JSON：${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH}（模板索引；Word 写入锚点优先从 .docx 动态解析）`,
    `标准参考资料：${BUILT_IN_STANDARD_REFERENCE_RELATIVE_PATH}`,
    "图片或截图附件：如果 Chat 模型已配置图像输入能力，可直接依据本轮图片分析；否则使用 read_image 识别。文本/Word/PDF 附件使用对应读取工具；Word 附件只作为普通资料读取，不能注册或切换为模板；模板以当前前端选择为准。"
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

function renderSelectedDocumentTemplateContext(): string {
  const template = getSelectedDocumentTemplate();
  const lines = [
    "当前选中 Word 文档模板（系统上下文，不属于用户输入正文）：",
    `- template_id: ${JSON.stringify(template.id)}`,
    `- template_name: ${JSON.stringify(template.name)}`,
    template.profilePath ? `- profile_path: ${JSON.stringify(template.profilePath)}` : "",
    template.templatePath ? `- template_path: ${JSON.stringify(template.templatePath)}` : "",
    template.templateJsonPath ? `- template_json_path: ${JSON.stringify(template.templateJsonPath)}` : "",
    `- render_mode: ${JSON.stringify(template.renderMode ?? "template_sections")}`,
    "生成流程：先基于上述 profile_path 和当前任务完成整体规划，并调用 build_document_config 将规划写入 document-config.json；再调用 list_document_sections 建立章节索引，调用 draft_document_sections 按 section id 并行起草章节，后续按 section id 保存/微调 document-sections/<section>/draft.md；合并前调用 audit_document_sections/revise_document_sections_evidence 做章节级证据审查和修订，需要提升表达时调用 polish_document_sections，最后 assemble_document_sections 合并终稿；写 Word 时使用上述 template_path/template_json_path/render_mode。",
    "注意：template_json_path 只用于 write_document_word 的模板结构 JSON，禁止把 template_json_path 当作 profile_path。",
    "profile 微调：只有 profile_path 以 document-templates/ 开头时，才可在用户明确要求微调生成规则后读取该 profile 并调用 update_document_profile；其他 profile 只读取使用，不要更新。",
    "不要要求用户再次上传当前模板；用户切换模板后，本系统上下文会自动更新。"
  ].filter(Boolean);
  return lines.join("\n");
}

function renderSchemeProgressRuntimeContext(sessionId: string): string {
  const progress = getLatestSchemeProgressItem(sessionId);
  if (!progress || progress.total === 0) return "";

  const unit = "章节";
  const failedSections = progress.sections.filter((section) => section.status === "failed");
  const draftedSections = progress.sections.filter((section) => section.status === "drafted");
  const nextFailed = failedSections[0];
  const nextDrafted = draftedSections[0];
  const nextPending = progress.sections.find((section) => section.status === "pending");
  const nextSection = nextFailed ?? nextDrafted ?? nextPending;
  const recentCompleted = progress.sections
    .filter((section) => section.status === "completed")
    .slice(-5)
    .map((section) => `${section.number} ${section.title}`);

  return [
    `当前${unit}生成进度：`,
    `- 当前文件：${progress.artifactName || "尚未生成最终文件"}`,
    `- 已起草待写入：${progress.drafted ?? draftedSections.length}`,
    `- 已完成：${progress.completed}/${progress.total}`,
    `- 失败：${progress.failed}`,
    nextSection
      ? `- 下一步优先处理：${nextSection.number} ${nextSection.title}（${
          nextSection.status === "failed" ? "失败重试" : nextSection.status === "drafted" ? "草稿待处理" : `下一待起草${unit}`
        }）`
      : "- 下一步优先处理：无",
    recentCompleted.length ? `- 最近完成：${recentCompleted.join("、")}` : "",
    "继续交付时按统一 section-first 流程 list_document_sections、draft_document_sections、assemble_document_sections、audit_document_evidence、write_document_word 推进。",
    "不得基于中间进度声称完整方案已完成；最终状态以本轮通用流程实际生成的文件为准。"
  ]
    .filter(Boolean)
    .join("\n");
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

async function buildPromptImages(
  attachments: AttachmentRef[] | undefined,
  settings: AppSettings
): Promise<ImageContent[] | undefined> {
  if (!settings.openai.chatImageInputEnabled || !attachments?.length) return undefined;

  const images: ImageContent[] = [];
  for (const attachment of attachments) {
    if (!isSupportedImageFile(attachment.path)) continue;
    const image = await readImageDataUrl(attachment.path, MAX_VISION_IMAGE_BYTES);
    images.push({
      type: "image",
      data: image.dataBase64,
      mimeType: image.mimeType
    });
  }

  return images.length ? images : undefined;
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
    getSessionInputDir: (sessionId) => getSessionWorkspaceDirs(sessionId).inputDir,
    getSessionOutputDir: (sessionId) => getSessionWorkspaceDirs(sessionId).outputDir,
    loadSettings: loadEnv,
    getSession: (sessionId) => sessions.get(sessionId),
    buildMessages,
    createAssistantMessage,
    updateItem,
    startToolCall,
    finishToolCall,
    addStage,
    ensureSchemeProgress,
    updateSchemeSectionProgress,
    settleSchemeProgress,
    appendSessionMemory,
    formatSessionMemory,
    getSessionReadableFiles,
    getBundledPythonRuntime: () => findBundledPythonRuntime({ rootDir: projectRootDir, resourcesDir }),
    createArtifact: (sessionId, filePath) => {
      createArtifact(sessionId, filePath);
    }
  };
}

async function runAgentResponse(
  sessionId: string,
  controller: AbortController,
  userPrompt: string,
  onAssistantCreated: (assistantItem: MessageStreamItem) => void,
  images?: ImageContent[]
): Promise<MessageStreamItem> {
  const runtime = createAgentRuntime(createAgentRuntimeHost());
  return runtime.runTurn({
    sessionId,
    userPrompt,
    images,
    controller,
    onAssistantCreated
  });
}

function applySchemeCompletionNotice(sessionId: string, assistantItem: MessageStreamItem): void {
  const progress = getLatestSchemeProgressItem(sessionId);
  if (!progress || progress.total === 0 || progress.status === "completed") return;
  if (progress.completed === progress.total && progress.failed === 0) return;
  if (assistantItem.content.includes("生成未完成：当前")) return;

  const unit = "章节";
  const nextFailed = progress.sections.find((section) => section.status === "failed");
  const nextPending = progress.sections.find((section) => section.status === "pending");
  const nextSection = nextFailed ?? nextPending;
  const nextText = nextSection ? `下一步应继续处理 ${nextSection.number} ${nextSection.title}。` : `下一步应继续补齐未完成${unit}。`;
  const notice = [
    `> ${unit}生成未完成：当前 ${progress.completed}/${progress.total}${progress.failed ? `，失败 ${progress.failed}` : ""}。`,
    "> 当前文件只能视为阶段性文件，不是完整方案。",
    `> ${nextText}`
  ].join("\n");

  assistantItem.content = [notice, assistantItem.content.trim()].filter(Boolean).join("\n\n");
}

async function handlePrompt(input: ChatPromptInput): Promise<{ accepted: true }> {
  const session = sessions.get(input.sessionId);
  if (!session) {
    throw new Error("Session not found");
  }
  ensureSessionWorkspaceDirs(input.sessionId);
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
      const promptImages = await buildPromptImages(input.attachments, loadEnv());
      assistantItem = await runAgentResponse(
        input.sessionId,
        controller,
        input.message,
        (createdItem) => {
          assistantItem = createdItem;
        },
        promptImages
      );
      applySchemeCompletionNotice(input.sessionId, assistantItem);
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
  const sessionId = input.sessionId;
  if (!sessionId || !sessions.has(sessionId)) {
    throw new Error("Session not found");
  }
  const workspace = ensureSessionWorkspaceDirs(sessionId);
  const filters = [
    { name: "Supported files", extensions: ["docx", "pdf", "md", "txt", "png", "jpg", "jpeg", "webp", "svg", "xlsx"] },
    { name: "All files", extensions: ["*"] }
  ];
  const result = await dialog.showOpenDialog({
    properties: input.multiple === false ? ["openFile"] : ["openFile", "multiSelections"],
    filters
  });
  if (result.canceled) return [];

  return result.filePaths.map((filePath) => {
    const id = createId("attachment");
    const safeName = sanitizeAttachmentName(basename(filePath));
    const workspacePath = join(workspace.inputDir, `${id}-${safeName}`);
    copyFileSync(filePath, workspacePath);
    const attachment: AttachmentRef = {
      id,
      sessionId,
      name: safeName,
      mimeType: "application/octet-stream",
      size: statSync(workspacePath).size,
      path: workspacePath,
      source: "picker",
      createdAt: now()
    };
    attachments.set(attachment.id, attachment);
    schedulePersistState();
    return attachment;
  });
}

function importClipboardAttachments(input: ClipboardAttachmentInput): AttachmentRef[] {
  const sessionId = input.sessionId;
  if (!sessionId || !sessions.has(sessionId)) {
    throw new Error("Session not found");
  }
  ensureDataDirs();
  const workspace = ensureSessionWorkspaceDirs(sessionId);
  return prepareImportedAttachments(input, {
    inputDir: workspace.inputDir,
    createId: () => createId("attachment")
  }).map((file) => {
    writeFileSync(file.outputPath, file.buffer);
    const attachment: AttachmentRef = {
      id: file.id,
      sessionId,
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
  ipcMain.handle("app:metadata", () => getAppMetadata());
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
  ipcMain.handle("document-template:list", () => listDocumentTemplates());
  ipcMain.handle("document-template:get-selected", () => getSelectedDocumentTemplate());
  ipcMain.handle("document-template:select", (_event, templateId: string) => selectDocumentTemplate(templateId));
  ipcMain.handle("document-template:upload", () => uploadDocumentTemplate());
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
    title: getAppMetadata().title,
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
