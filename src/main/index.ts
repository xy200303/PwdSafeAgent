import electron from "electron";
import type { BrowserWindow as BrowserWindowType } from "electron";
import { join, extname, basename, resolve } from "node:path";
import { existsSync, mkdirSync, statSync, writeFileSync } from "node:fs";
import type { ChatCompletionMessageParam } from "openai/resources/chat/completions";
import dotenv from "dotenv";
import { compactText, getCurrentTimeText } from "./agentTools";
import { createAgentRuntime, type AgentRuntimeHost, type MessageStreamItem } from "./agentRuntime";
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
import {
  applySchemeProgressUpdate,
  createSchemeProgressItem,
  settleSchemeProgressItem,
  type SchemeProgressUpdateInput
} from "./schemeProgress";
import { renderSchemeChapterGuide } from "./schemePlan";
import { loadPersistedState, savePersistedState, type PersistedStateSnapshot, type SessionMemoryEntry } from "./sessionPersistence";
import {
  IMAGE_GENERATION_REQUEST_TIMEOUT_DEFAULT_MS,
  clampDraftSectionParallelism,
  clampImageGenerationParallelism
} from "../shared/types";
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
  SchemeProgressItem,
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
      thinkingEnabled: parseBooleanEnv(process.env.OPENAI_THINKING_ENABLED, true),
      reasoningEffort: process.env.OPENAI_REASONING_EFFORT || "",
      requestTimeoutMs: Number(process.env.OPENAI_REQUEST_TIMEOUT_MS || 120000),
      imageRequestTimeoutMs: Number(
        process.env.OPENAI_IMAGE_REQUEST_TIMEOUT_MS || IMAGE_GENERATION_REQUEST_TIMEOUT_DEFAULT_MS
      ),
      maxOutputTokens: Number(process.env.OPENAI_MAX_OUTPUT_TOKENS || 16000),
      apiKeyConfigured: Boolean(process.env.OPENAI_API_KEY),
      imageApiKeyConfigured: Boolean(process.env.OPENAI_IMAGE_API_KEY)
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

function ensureSchemeProgress(sessionId: string, detail?: string, artifactName?: string): void {
  const existing = getLatestSchemeProgressItem(sessionId);
  const timestamp = now();
  if (!existing) {
    const item = createSchemeProgressItem({
      id: createId("scheme_progress"),
      docsDir,
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
    detail: detail ?? existing.detail,
    artifactName: artifactName ?? existing.artifactName,
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
    sendEvent({ id: createId("event"), type: "stream.item.updated", sessionId, payload: fileItem });
    sendEvent({ id: createId("event"), type: "session.updated", payload: session });
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
    "只有当用户明确说“生成/导出/输出/形成方案/出 Word/PDF/画图”等交付意图，或项目档案已经标记为生成就绪时，才调用 create_word、write_word、write_pdf、image_generate、send_file 等交付工具。",
    "Word 交付只使用 Word 模板和模板标注 JSON：先按 docs/密码应用方案.template.json 的章节写作提示、不可见 SDT 锚、表格锚点和图片锚点组织内容，再用 docs/密码应用方案.docx 作为最终 Word 样式模板。",
    "整篇交付必须按节推进：先 read_file 读取 docs/密码应用方案.template.json，再调用 plan_scheme_batches 生成稳定批次，然后 create_word 复制内置 Word 模板。正文生成必须按 plan_scheme_batches 返回的 first_draft_call/批次调用 draft_scheme_sections，一批尽量接近设置并行数，再把返回草稿按 JSON sections 顺序合并为 write_word.sections，一次批量写入同一个 path。section 优先传模板 JSON 的 id（如 sec_1_2_1），也可传章节编号和标题。不要把整篇方案合成一段长 Markdown 后一次性写入。",
    "章节写入必须按 docs/密码应用方案.template.json 中 sections 数组顺序推进；除非用户明确要求局部更新，不要跳章、不要抽样式填充多个章节。",
    "draft_scheme_sections 只用于并行起草正文，不写 Word、不生成表格、不生成图片；正文草稿完成后优先用 write_word.sections 批量写入，避免反复打开和保存同一个 docx。",
    "draft_scheme_sections 应使用 plan_scheme_batches 返回的批次参数，一次传入同一批待生成章节，优先接近设置中的章节并行数；每个数组项只对应一个模板章节或小节，并原样使用返回的 paragraph_tasks、writingHint、placeholders、relatedTables、relatedFigures。",
    "每个章节正文必须按 paragraph_tasks 分成多个自然段，段落之间承接上一段，首段承接上一节，末段自然引出下一节；不要把一整节写成单段长文。",
    "正文全部写入后，必须调用 plan_scheme_assets 规划 relatedTables/relatedFigures；表格按返回的 template_cells_plan 改写 value 后用 write_word.template_cells 精确写入。如果 next_plan_scheme_assets_call 不为空，必须继续规划和写入下一批，直到工具显示没有后续，避免表格只填一部分。图片按 image_generate_plan 并行生成后用 write_word.diagrams 传 figure_id、label、kind、path 精确嵌入。不要把【待填写】和【图片占位】留在最终交付版本里。",
    "局部更新同样使用 create_word 或已有 docx 路径；多章节更新优先传 write_word.sections，单章节更新才传 path、section 和该章节 content。",
    "write_word 不要求一次性完成所有章节；资料不足时可以先写已确认章节，后续继续增量替换。不要为了通过完整性检查而编造用户未提供的关键事实。",
    "只要章节进度未达到全部完成，或存在失败章节，最终回复必须称为阶段性文件/部分完成，不得说完整方案已生成、全部完成或已生成完整方案。",
    "降低 AI 味：每一节只写与该节相关的项目事实、现状、风险、控制措施、算法/产品/部署位置/调用路径；避免泛泛而谈、套话开头、重复政策背景和空洞排比。没有事实就明确写“待补充/需确认”。",
    "需要方案图示时由工具真实生成：正文章节写完后先调用 plan_scheme_assets 获取 figure_id 和 label，再并行调用多个 image_generate 生成架构图、拓扑图、流程图；全部图片生成完成后，再调用 write_word 并在 diagrams 中传入 figure_id 精确嵌入。",
    renderSchemeChapterGuide(),
    "回复使用中文 Markdown，必要时给出缺失资料清单。",
    "不要编造用户未提供的关键事实；资料不足时可以说明“待补充/需确认”，并在后续获得信息后继续替换对应章节。",
    "工具调用由你按任务需要自主决策：寒暄、普通问答和资料澄清阶段不要默认读取模板或附件；只有分析附件、查询最新资料、沉淀项目档案或交付文件确有需要时，才调用对应工具。",
    "当进入交付阶段并需要 Word、PDF、Markdown 或图片文件时，必须通过 create_word/write_word/write_pdf/write_file/image_generate 等工具真实生成文件；不要只在文本回复中声称已经生成。",
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

  const schemeProgressContext = renderSchemeProgressRuntimeContext(session.id);
  if (schemeProgressContext) {
    messages.push({
      role: "system",
      content: schemeProgressContext
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
    "Word 样式模板：docs/密码应用方案.docx",
    "Word 模板标注 JSON：docs/密码应用方案.template.json（read_file 会返回规范化章节规划任务清单，不返回原始 JSON）"
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

function renderSchemeProgressRuntimeContext(sessionId: string): string {
  const progress = getLatestSchemeProgressItem(sessionId);
  if (!progress || progress.total === 0) return "";

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
    "当前 Word 章节生成进度（必须遵守）：",
    `- 当前文件：${progress.artifactName || "已创建的模板副本，路径见最近一次 create_word/write_word 工具结果"}`,
    `- 已起草待写入：${progress.drafted ?? draftedSections.length}`,
    `- 已完成：${progress.completed}/${progress.total}`,
    `- 失败：${progress.failed}`,
    nextSection
      ? `- 下一步优先处理：${nextSection.number} ${nextSection.title}（${
          nextSection.status === "failed" ? "失败重试" : nextSection.status === "drafted" ? "草稿待写入 Word" : "下一待起草"
        }）`
      : "- 下一步优先处理：无",
    recentCompleted.length ? `- 最近完成：${recentCompleted.join("、")}` : "",
    draftedSections.length
      ? "继续生成时，先把已起草章节按模板 JSON 顺序合并到 write_word.sections 批量写入 Word，再起草新章节。"
      : "继续生成时，先调用 plan_scheme_batches 生成下一批待起草章节，再按 first_draft_call 调用 draft_scheme_sections；不要直接手写单个 section。",
    "表格 template_cells 和配图 image_generate/diagrams 放在所有正文章节完成后统一处理。",
    "不得跳到后续章节抽样填充。若本轮没有把 completed 写到 total，最终回复只能说阶段性文件/部分完成。"
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
  onAssistantCreated: (assistantItem: MessageStreamItem) => void
): Promise<MessageStreamItem> {
  const runtime = createAgentRuntime(createAgentRuntimeHost());
  return runtime.runTurn({
    sessionId,
    userPrompt,
    controller,
    onAssistantCreated
  });
}

function applySchemeCompletionNotice(sessionId: string, assistantItem: MessageStreamItem): void {
  const progress = getLatestSchemeProgressItem(sessionId);
  if (!progress || progress.total === 0 || progress.status === "completed") return;
  if (progress.completed === progress.total && progress.failed === 0) return;
  if (assistantItem.content.includes("章节生成未完成：当前")) return;

  const nextFailed = progress.sections.find((section) => section.status === "failed");
  const nextPending = progress.sections.find((section) => section.status === "pending");
  const nextSection = nextFailed ?? nextPending;
  const nextText = nextSection ? `下一步应继续处理 ${nextSection.number} ${nextSection.title}。` : "下一步应继续补齐未完成章节。";
  const notice = [
    `> 章节生成未完成：当前 ${progress.completed}/${progress.total}${progress.failed ? `，失败 ${progress.failed}` : ""}。`,
    "> 当前 Word 只能视为阶段性文件，不是完整方案。",
    `> ${nextText}`
  ].join("\n");

  assistantItem.content = [notice, assistantItem.content.trim()].filter(Boolean).join("\n\n");
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
