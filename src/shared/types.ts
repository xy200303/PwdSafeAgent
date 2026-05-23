export type MessageRole = "user" | "assistant";

export type SessionStatus = "idle" | "running" | "completed" | "failed";

export type AgentRuntimeKind = "openai-chat" | "pi-agent";

export type BundledPythonRuntimeSource = "resources" | "project";

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
    };

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
    chatModel: string;
    imageModel: string;
    imageSize: string;
    imageQuality: string;
    autoImageGeneration: boolean;
    requestTimeoutMs: number;
    maxOutputTokens: number;
    apiKeyConfigured: boolean;
    imageApiKeyConfigured: boolean;
  };
  document: {
    autoPdfExport: boolean;
    libreOfficePath: string;
  };
  agent: {
    runtime: AgentRuntimeKind;
    execBashEnabled: boolean;
    piAgentPackage: string;
    piAgentExport: string;
  };
}

export interface UpdateAppSettingsInput {
  openai: {
    baseUrl: string;
    imageBaseUrl: string;
    chatModel: string;
    imageModel: string;
    imageSize: string;
    imageQuality: string;
    autoImageGeneration: boolean;
    requestTimeoutMs: number;
    maxOutputTokens: number;
    apiKey?: string;
    imageApiKey?: string;
  };
  document: {
    autoPdfExport: boolean;
    libreOfficePath: string;
  };
  agent: {
    runtime: AgentRuntimeKind;
    execBashEnabled: boolean;
    piAgentPackage: string;
    piAgentExport: string;
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
