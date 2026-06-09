export type MessageRole = "user" | "assistant";

export type SessionStatus = "idle" | "running" | "completed" | "failed";

export type BundledPythonRuntimeSource = "resources" | "project";

export const DRAFT_SECTION_PARALLELISM_MIN = 1;
export const DRAFT_SECTION_PARALLELISM_MAX = 20;
export const DRAFT_SECTION_PARALLELISM_DEFAULT = 20;
export const IMAGE_GENERATION_REQUEST_TIMEOUT_DEFAULT_MS = 300000;
export const IMAGE_GENERATION_PARALLELISM_MIN = 1;
export const IMAGE_GENERATION_PARALLELISM_MAX = 10;
export const IMAGE_GENERATION_PARALLELISM_DEFAULT = 10;

export function clampDraftSectionParallelism(value: unknown): number {
  const numericValue =
    typeof value === "number" ? value : typeof value === "string" && value.trim() ? Number(value.trim()) : Number.NaN;
  if (!Number.isFinite(numericValue)) return DRAFT_SECTION_PARALLELISM_DEFAULT;
  return Math.min(
    Math.max(Math.trunc(numericValue), DRAFT_SECTION_PARALLELISM_MIN),
    DRAFT_SECTION_PARALLELISM_MAX
  );
}

export function clampImageGenerationParallelism(value: unknown): number {
  const numericValue =
    typeof value === "number" ? value : typeof value === "string" && value.trim() ? Number(value.trim()) : Number.NaN;
  if (!Number.isFinite(numericValue)) return IMAGE_GENERATION_PARALLELISM_DEFAULT;
  return Math.min(
    Math.max(Math.trunc(numericValue), IMAGE_GENERATION_PARALLELISM_MIN),
    IMAGE_GENERATION_PARALLELISM_MAX
  );
}

export interface BundledPythonRuntimeStatus {
  available: boolean;
  source?: BundledPythonRuntimeSource;
  homeDir?: string;
  pythonExePath?: string;
  scriptsDir?: string;
}

export interface RuntimeCommandCheck {
  ok: boolean;
  command: string;
  source: "bundled" | "system";
  output?: string;
  error?: string;
  durationMs: number;
}

export interface RuntimeCheckResult {
  checkedAt: string;
  bundledPython: BundledPythonRuntimeStatus;
  python: RuntimeCommandCheck;
  pip: RuntimeCommandCheck;
}

export interface AppMetadata {
  displayName: string;
  version: string;
  title: string;
}

export type SchemeSectionStatus = "pending" | "drafting" | "drafted" | "running" | "completed" | "failed" | "skipped";

export interface SchemeProgressSection {
  id: string;
  number: string;
  title: string;
  headingLevel: number;
  status: SchemeSectionStatus;
  writingHint?: string;
  detail?: string;
  relatedTables?: string[];
  relatedFigures?: string[];
  updatedAt?: string;
}

export interface SchemeProgressItem {
  id: string;
  kind: "scheme_progress";
  title: string;
  status: "pending" | "running" | "partial" | "completed" | "failed";
  total: number;
  drafted: number;
  completed: number;
  failed: number;
  activeSectionId?: string;
  artifactName?: string;
  detail?: string;
  sections: SchemeProgressSection[];
  createdAt: string;
  updatedAt?: string;
}

export type StreamItem =
  | {
      id: string;
      kind: "message";
      role: MessageRole;
      content: string;
      isFinished: boolean;
      createdAt: string;
      attachmentIds?: string[];
    }
  | {
      id: string;
      kind: "tool";
      toolCallId: string;
      toolName: string;
      status: "running" | "success" | "failed";
      summary?: string;
      inputPreview?: string;
      outputPreview?: string;
      errorPreview?: string;
      createdAt: string;
    }
  | {
      id: string;
      kind: "file";
      artifactId: string;
      name: string;
      fileKind: ArtifactKind;
      createdAt: string;
    }
  | {
      id: string;
      kind: "stage";
      title: string;
      detail?: string;
      createdAt: string;
    }
  | SchemeProgressItem;

export interface ChatSession {
  id: string;
  title: string;
  status: SessionStatus;
  createdAt: string;
  updatedAt: string;
  items: StreamItem[];
}

export type ArtifactKind = "docx" | "pdf" | "png" | "jpg" | "jpeg" | "webp" | "svg" | "json" | "md" | "txt" | "log" | "other";

export interface ArtifactSummary {
  id: string;
  sessionId?: string;
  name: string;
  kind: ArtifactKind;
  path: string;
  size: number;
  createdAt: string;
}

export interface ArtifactListInput {
  sessionId?: string;
}

export interface ArtifactPreview {
  artifactId: string;
  name: string;
  kind: ArtifactKind;
  mode: "text" | "image" | "pdf" | "docx" | "unsupported";
  size?: number;
  text?: string;
  dataUrl?: string;
  mimeType?: string;
  summary?: string;
}

export interface AttachmentRef {
  id: string;
  sessionId?: string;
  name: string;
  mimeType: string;
  size: number;
  path: string;
  source: "picker" | "clipboard" | "drop";
  createdAt: string;
}

export interface ClipboardAttachmentInput {
  sessionId?: string;
  source?: "clipboard" | "drop";
  files: Array<{
    name: string;
    mimeType: string;
    dataBase64: string;
  }>;
}

export interface PickAttachmentInput {
  sessionId?: string;
  multiple?: boolean;
}

export interface DocumentTemplateSummary {
  id: string;
  name: string;
  source: "builtin" | "uploaded";
  profilePath?: string;
  templatePath?: string;
  templateJsonPath?: string;
  renderMode?: "full_document" | "template_sections";
  sectionCount?: number;
  sectionGroupCount?: number;
  createdAt?: string;
}

export interface RenameSessionInput {
  sessionId: string;
  title: string;
}

export interface ChatPromptInput {
  sessionId: string;
  message: string;
  attachments?: AttachmentRef[];
}

export interface AppSettings {
  runtime: {
    envFilePath: string;
    configSource: ".env" | ".env.local" | "process";
    bundledPython: BundledPythonRuntimeStatus;
  };
  openai: {
    baseUrl: string;
    imageBaseUrl: string;
    visionBaseUrl: string;
    chatModel: string;
    chatImageInputEnabled: boolean;
    imageModel: string;
    visionModel: string;
    imageSize: string;
    imageQuality: string;
    autoImageGeneration: boolean;
    thinkingEnabled: boolean;
    reasoningEffort: string;
    requestTimeoutMs: number;
    imageRequestTimeoutMs: number;
    maxOutputTokens: number;
    apiKeyConfigured: boolean;
    imageApiKeyConfigured: boolean;
    visionApiKeyConfigured: boolean;
  };
  document: {
    autoPdfExport: boolean;
    libreOfficePath: string;
  };
  agent: {
    execBashEnabled: boolean;
    draftSectionParallelism: number;
    imageGenerationParallelism: number;
  };
}

export interface UpdateAppSettingsInput {
  openai: {
    baseUrl: string;
    imageBaseUrl: string;
    visionBaseUrl: string;
    chatModel: string;
    chatImageInputEnabled: boolean;
    imageModel: string;
    visionModel: string;
    imageSize: string;
    imageQuality: string;
    autoImageGeneration: boolean;
    thinkingEnabled: boolean;
    reasoningEffort: string;
    requestTimeoutMs: number;
    imageRequestTimeoutMs: number;
    maxOutputTokens: number;
    apiKey?: string;
    imageApiKey?: string;
    visionApiKey?: string;
  };
  document: {
    autoPdfExport: boolean;
    libreOfficePath: string;
  };
  agent: {
    execBashEnabled: boolean;
    draftSectionParallelism: number;
    imageGenerationParallelism: number;
  };
}

export type RendererEvent =
  | {
      id: string;
      type: "session.created";
      payload: ChatSession;
    }
  | {
      id: string;
      type: "session.updated";
      payload: ChatSession;
    }
  | {
      id: string;
      type: "session.deleted";
      sessionId: string;
    }
  | {
      id: string;
      type: "stream.item.added";
      sessionId: string;
      payload: StreamItem;
    }
  | {
      id: string;
      type: "stream.item.updated";
      sessionId: string;
      payload: StreamItem;
    }
  | {
      id: string;
      type: "artifact.created";
      sessionId: string;
      payload: ArtifactSummary;
    };

export interface PwdSafeAgentApi {
  app: {
    getMetadata(): Promise<AppMetadata>;
  };
  session: {
    list(): Promise<ChatSession[]>;
    create(): Promise<ChatSession>;
    rename(input: RenameSessionInput): Promise<ChatSession>;
    delete(sessionId: string): Promise<ChatSession[]>;
  };
  chat: {
    prompt(input: ChatPromptInput): Promise<{ accepted: true }>;
    abort(sessionId: string): Promise<void>;
  };
  attachment: {
    pick(input: PickAttachmentInput): Promise<AttachmentRef[]>;
    importClipboard(input: ClipboardAttachmentInput): Promise<AttachmentRef[]>;
    remove(attachmentId: string): Promise<void>;
  };
  documentTemplate: {
    list(): Promise<DocumentTemplateSummary[]>;
    getSelected(): Promise<DocumentTemplateSummary>;
    select(templateId: string): Promise<DocumentTemplateSummary>;
    upload(): Promise<DocumentTemplateSummary | undefined>;
  };
  artifact: {
    list(input?: ArtifactListInput): Promise<ArtifactSummary[]>;
    open(artifactId: string): Promise<void>;
    reveal(artifactId: string): Promise<void>;
    preview(artifactId: string): Promise<ArtifactPreview>;
  };
  settings: {
    get(): Promise<AppSettings>;
    save(input: UpdateAppSettingsInput): Promise<AppSettings>;
    checkRuntime(): Promise<RuntimeCheckResult>;
  };
  events: {
    subscribe(listener: (event: RendererEvent) => void): () => void;
  };
}
