import { useEffect, useRef, useState, type ClipboardEvent, type FormEvent, type JSX } from "react";
import * as AlertDialog from "@radix-ui/react-alert-dialog";
import * as Dialog from "@radix-ui/react-dialog";
import * as DropdownMenu from "@radix-ui/react-dropdown-menu";
import * as ScrollArea from "@radix-ui/react-scroll-area";
import * as Switch from "@radix-ui/react-switch";
import * as Tabs from "@radix-ui/react-tabs";
import * as Tooltip from "@radix-ui/react-tooltip";
import { IncremarkContent } from "@incremark/react";
import {
  Bot,
  CheckCircle2,
  Eye,
  File,
  Files,
  FolderOpen,
  Image,
  Loader2,
  MoreHorizontal,
  Paperclip,
  PencilLine,
  Play,
  Plus,
  Search,
  Settings,
  Square,
  Trash2,
  Upload,
  User,
  X,
  XCircle
} from "lucide-react";
import {
  addAttachments,
  applyRendererEvent,
  clearComposer,
  removeAttachment,
  setArtifacts,
  setArtifactsOpen,
  setComposer,
  setCurrentSession,
  setSessions,
  setSettings,
  setSettingsOpen,
  upsertSession,
  type RootState
} from "../store";
import { useDispatch, useSelector } from "react-redux";
import type { AppDispatch } from "../store";
import type { ReactNode } from "react";
import type {
  AppSettings,
  ArtifactKind,
  ArtifactPreview,
  ArtifactSummary,
  AttachmentRef,
  ChatSession,
  PwdSafeAgentApi,
  StreamItem
} from "../../../shared/types";
import { orderStreamItemsForDisplay } from "../streamOrdering";

export function App(): JSX.Element {
  const dispatch = useDispatch<AppDispatch>();
  const { sessions, artifacts, currentSessionId, composer, pendingAttachments, settings, settingsOpen, artifactsOpen } = useSelector(
    (state: RootState) => state.chat
  );
  const [preview, setPreview] = useState<ArtifactPreview | undefined>();
  const [previewError, setPreviewError] = useState("");
  const [renameTarget, setRenameTarget] = useState<ChatSession | undefined>();
  const [deleteTarget, setDeleteTarget] = useState<ChatSession | undefined>();
  const currentSession = sessions.find((session) => session.id === currentSessionId);
  const api = window.pwdSafeAgent;

  useEffect(() => {
    if (!api) return;
    void bootstrap(dispatch, api);
    const unsubscribe = api.events.subscribe((event) => {
      dispatch(applyRendererEvent(event));
    });
    return unsubscribe;
  }, [api, dispatch]);

  async function createConversation(): Promise<void> {
    if (!api) return;
    const session = await api.session.create();
    dispatch(upsertSession(session));
    dispatch(setCurrentSession(session.id));
  }

  async function renameConversation(sessionId: string, title: string): Promise<void> {
    if (!api) return;
    const session = await api.session.rename({ sessionId, title });
    dispatch(upsertSession(session));
  }

  async function deleteConversation(sessionId: string): Promise<void> {
    if (!api) return;
    const nextSessions = await api.session.delete(sessionId);
    dispatch(setSessions(nextSessions));
    dispatch(setArtifacts(await api.artifact.list()));
  }

  async function previewArtifact(artifactId: string): Promise<void> {
    if (!api) return;
    setPreviewError("");
    try {
      setPreview(await api.artifact.preview(artifactId));
    } catch (error) {
      setPreview(undefined);
      setPreviewError(error instanceof Error ? error.message : String(error));
    }
  }

  async function openArtifactHistory(): Promise<void> {
    if (!api) return;
    dispatch(setArtifacts(await api.artifact.list()));
    dispatch(setArtifactsOpen(true));
  }

  if (!api) {
    return <MissingBridge />;
  }

  return (
    <Tooltip.Provider delayDuration={350}>
      <div className="app-shell">
        <Sidebar
          sessions={sessions}
          artifactCount={artifacts.length}
          currentSessionId={currentSessionId}
          onCreateConversation={createConversation}
          onSelectSession={(sessionId) => dispatch(setCurrentSession(sessionId))}
          onRenameSession={(session) => setRenameTarget(session)}
          onDeleteSession={(session) => setDeleteTarget(session)}
          onOpenArtifacts={() => void openArtifactHistory()}
          onOpenSettings={() => dispatch(setSettingsOpen(true))}
        />
        <main className="content-area">
          <Header session={currentSession} settings={settings} />
          <MessagePane session={currentSession} onPreviewArtifact={(artifactId) => void previewArtifact(artifactId)} />
          <Composer
            api={api}
            session={currentSession}
            value={composer}
            attachments={pendingAttachments}
            onChange={(value) => dispatch(setComposer(value))}
            onAddAttachments={(items) => dispatch(addAttachments(items))}
            onRemoveAttachment={(id) => {
              dispatch(removeAttachment(id));
              void api.attachment.remove(id);
            }}
            onSent={() => dispatch(clearComposer())}
          />
        </main>
        {settingsOpen ? (
          <SettingsPanel
            api={api}
            settings={settings}
            onClose={() => dispatch(setSettingsOpen(false))}
            onSaved={(next) => dispatch(setSettings(next))}
          />
        ) : null}
        {preview || previewError ? (
          <ArtifactPreviewPanel
            preview={preview}
            error={previewError}
            onClose={() => {
              setPreview(undefined);
              setPreviewError("");
            }}
          />
        ) : null}
        {artifactsOpen ? (
          <ArtifactHistoryPanel
            artifacts={artifacts}
            sessions={sessions}
            onClose={() => dispatch(setArtifactsOpen(false))}
            onPreviewArtifact={(artifactId) => {
              dispatch(setArtifactsOpen(false));
              void previewArtifact(artifactId);
            }}
          />
        ) : null}
        {renameTarget ? (
          <RenameSessionDialog
            session={renameTarget}
            onClose={() => setRenameTarget(undefined)}
            onRename={(sessionId, title) => renameConversation(sessionId, title)}
          />
        ) : null}
        {deleteTarget ? (
          <DeleteSessionDialog
            session={deleteTarget}
            onClose={() => setDeleteTarget(undefined)}
            onDelete={(sessionId) => deleteConversation(sessionId)}
          />
        ) : null}
      </div>
    </Tooltip.Provider>
  );
}

function MissingBridge(): JSX.Element {
  return (
    <div className="bridge-missing">
      <strong>Electron 预加载桥未就绪</strong>
      <span>请通过 `npm run dev` 或打包后的 Electron 应用启动，不要直接在浏览器打开渲染页。</span>
      <small>如果仍然出现该提示，请重启开发服务；preload 会注入 `window.pwdSafeAgent`。</small>
    </div>
  );
}

async function bootstrap(dispatch: AppDispatch, api: PwdSafeAgentApi): Promise<void> {
  const [sessions, settings, artifacts] = await Promise.all([api.session.list(), api.settings.get(), api.artifact.list()]);
  dispatch(setSessions(sessions));
  dispatch(setSettings(settings));
  dispatch(setArtifacts(artifacts));
  if (sessions.length === 0) {
    const session = await api.session.create();
    dispatch(upsertSession(session));
    dispatch(setCurrentSession(session.id));
  }
}

function Sidebar(props: {
  sessions: ChatSession[];
  artifactCount: number;
  currentSessionId?: string;
  onCreateConversation: () => void;
  onSelectSession: (sessionId: string) => void;
  onRenameSession: (session: ChatSession) => void;
  onDeleteSession: (session: ChatSession) => void;
  onOpenArtifacts: () => void;
  onOpenSettings: () => void;
}): JSX.Element {
  const [query, setQuery] = useState("");
  const visibleSessions = props.sessions.filter((session) => session.title.toLowerCase().includes(query.toLowerCase()));

  return (
    <aside className="sidebar">
      <div className="brand">
        <div className="brand-mark">P</div>
        <div>
          <div className="brand-name">PwdSafeAgent</div>
          <div className="brand-subtitle">密码方案生成</div>
        </div>
      </div>

      <button className="primary-action" onClick={props.onCreateConversation}>
        <Plus size={16} />
        新建对话
      </button>

      <label className="search-box">
        <Search size={15} />
        <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="搜索对话" />
      </label>

      <ScrollArea.Root className="session-scroll">
        <ScrollArea.Viewport className="session-list">
          {visibleSessions.map((session) => (
            <div
              key={session.id}
              className={session.id === props.currentSessionId ? "session-row active" : "session-row"}
            >
              <button className="session-row-main" onClick={() => props.onSelectSession(session.id)}>
                <span className="session-title">{session.title}</span>
                <span className="session-meta">{statusLabel(session.status)}</span>
              </button>
              <DropdownMenu.Root>
                <DropdownMenu.Trigger asChild>
                  <button className="session-menu-button" aria-label={`打开 ${session.title} 的操作菜单`}>
                    <MoreHorizontal size={15} />
                  </button>
                </DropdownMenu.Trigger>
                <DropdownMenu.Portal>
                  <DropdownMenu.Content className="dropdown-content" side="right" align="start" sideOffset={6}>
                    <DropdownMenu.Item className="dropdown-item" onSelect={() => props.onRenameSession(session)}>
                      <PencilLine size={14} />
                      重命名
                    </DropdownMenu.Item>
                    <DropdownMenu.Item
                      className="dropdown-item danger"
                      onSelect={() => props.onDeleteSession(session)}
                    >
                      <Trash2 size={14} />
                      删除对话
                    </DropdownMenu.Item>
                  </DropdownMenu.Content>
                </DropdownMenu.Portal>
              </DropdownMenu.Root>
            </div>
          ))}
        </ScrollArea.Viewport>
        <ScrollArea.Scrollbar className="scrollbar" orientation="vertical">
          <ScrollArea.Thumb className="scrollbar-thumb" />
        </ScrollArea.Scrollbar>
      </ScrollArea.Root>

      <TooltipButton label="查看已生成的方案文件" className="settings-button" onClick={props.onOpenArtifacts}>
        <Files size={16} />
        交付文件
        {props.artifactCount ? <span className="nav-count">{props.artifactCount}</span> : null}
      </TooltipButton>

      <TooltipButton label="打开模型与导出设置" className="settings-button" onClick={props.onOpenSettings}>
        <Settings size={16} />
        设置
      </TooltipButton>
    </aside>
  );
}

function Header({ session, settings }: { session?: ChatSession; settings?: AppSettings }): JSX.Element {
  return (
    <header className="topbar">
      <div>
        <h1>{session?.title || "新的密码方案对话"}</h1>
        <p>
          模板：密码应用方案.docx · 模型：{settings?.openai.chatModel || "未加载"} ·
          {settings?.openai.apiKeyConfigured ? " API Key 已配置" : " 本地演示模式"}
        </p>
      </div>
      <div className={`status-pill ${session?.status === "running" ? "running" : ""}`}>
        {session?.status === "running" ? <Loader2 size={14} className="spin" /> : <CheckCircle2 size={14} />}
        {statusLabel(session?.status || "idle")}
      </div>
    </header>
  );
}

function MessagePane({
  session,
  onPreviewArtifact
}: {
  session?: ChatSession;
  onPreviewArtifact: (artifactId: string) => void;
}): JSX.Element {
  const scrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight, behavior: "smooth" });
  }, [session?.items.length, session?.items.at(-1)]);

  return (
    <ScrollArea.Root className="message-scroll">
      <ScrollArea.Viewport ref={scrollRef} className="message-pane">
        {!session || session.items.length === 0 ? (
          <div className="empty-state">
            <Bot size={28} />
            <strong>开始编制密码应用方案</strong>
            <span>输入系统名称、建设单位、等保级别或直接粘贴现有资料。</span>
          </div>
        ) : (
          orderStreamItemsForDisplay(session.items).map((item) => (
            <StreamRow key={item.id} item={item} onPreviewArtifact={onPreviewArtifact} />
          ))
        )}
      </ScrollArea.Viewport>
      <ScrollArea.Scrollbar className="scrollbar" orientation="vertical">
        <ScrollArea.Thumb className="scrollbar-thumb" />
      </ScrollArea.Scrollbar>
    </ScrollArea.Root>
  );
}

function StreamRow({
  item,
  onPreviewArtifact
}: {
  item: StreamItem;
  onPreviewArtifact: (artifactId: string) => void;
}): JSX.Element {
  if (item.kind === "message") {
    const isUser = item.role === "user";
    return (
      <article className={`message-row ${item.role}`}>
        {!isUser ? <div className="avatar"><Bot size={15} /></div> : null}
        <div className="message-body">
          {item.role === "assistant" ? (
            <IncremarkContent content={item.content} isFinished={item.isFinished} />
          ) : (
            <p>{item.content}</p>
          )}
        </div>
        {isUser ? <div className="avatar"><User size={15} /></div> : null}
      </article>
    );
  }

  if (item.kind === "tool") {
    return (
      <div className={`tool-row ${item.status}`}>
        {item.status === "running" ? (
          <Loader2 size={14} className="spin" />
        ) : item.status === "failed" ? (
          <XCircle size={14} />
        ) : (
          <CheckCircle2 size={14} />
        )}
        <span className="tool-name">{item.toolName}</span>
        <span className="tool-summary">{item.summary}</span>
      </div>
    );
  }

  if (item.kind === "file") {
    return (
      <div className="file-row">
        {isImageKind(item.fileKind) ? <Image size={16} /> : <File size={16} />}
        <span>{item.name}</span>
        <span className="file-kind">{item.fileKind}</span>
        <ArtifactActions artifactId={item.artifactId} onPreview={() => onPreviewArtifact(item.artifactId)} />
      </div>
    );
  }

  return (
    <div className="stage-row">
      <span>{item.title}</span>
      {item.detail ? <small>{item.detail}</small> : null}
    </div>
  );
}

function ArtifactActions({ artifactId, onPreview }: { artifactId: string; onPreview: () => void }): JSX.Element {
  return (
    <div className="file-actions">
      <button type="button" className="file-quick-action" onClick={onPreview}>
        <Eye size={14} />
        查看
      </button>
      <DropdownMenu.Root>
        <DropdownMenu.Trigger asChild>
          <button type="button" className="file-menu-button" aria-label="打开文件操作菜单">
            <MoreHorizontal size={15} />
          </button>
        </DropdownMenu.Trigger>
        <DropdownMenu.Portal>
          <DropdownMenu.Content className="dropdown-content" align="end" sideOffset={6}>
            <DropdownMenu.Item
              className="dropdown-item"
              onSelect={() => void window.pwdSafeAgent?.artifact.open(artifactId)}
            >
              <File size={14} />
              系统打开
            </DropdownMenu.Item>
            <DropdownMenu.Item
              className="dropdown-item"
              onSelect={() => void window.pwdSafeAgent?.artifact.reveal(artifactId)}
            >
              <FolderOpen size={14} />
              在文件夹中定位
            </DropdownMenu.Item>
          </DropdownMenu.Content>
        </DropdownMenu.Portal>
      </DropdownMenu.Root>
    </div>
  );
}

function ArtifactPreviewPanel({
  preview,
  error,
  onClose
}: {
  preview?: ArtifactPreview;
  error: string;
  onClose: () => void;
}): JSX.Element {
  return (
    <Dialog.Root
      open
      onOpenChange={(open) => {
        if (!open) onClose();
      }}
    >
      <Dialog.Portal>
        <Dialog.Overlay className="preview-backdrop" />
        <Dialog.Content className="preview-panel">
        <header>
          <div>
            <Dialog.Title asChild>
              <strong>{preview?.name || "预览失败"}</strong>
            </Dialog.Title>
            <span>{preview ? preview.kind.toUpperCase() : error}</span>
          </div>
          <Dialog.Close asChild>
            <button type="button" aria-label="关闭预览">
              <X size={16} />
            </button>
          </Dialog.Close>
        </header>
        {preview?.summary ? (
          <Dialog.Description className="preview-summary">{preview.summary}</Dialog.Description>
        ) : (
          <Dialog.Description className="sr-only">文件预览内容</Dialog.Description>
        )}
        <ScrollArea.Root className="preview-scroll">
        <ScrollArea.Viewport className="preview-body">
          {error ? <p className="preview-empty">{error}</p> : null}
          {preview?.mode === "image" && preview.dataUrl ? <img src={preview.dataUrl} alt={preview.name} /> : null}
          {preview?.mode === "pdf" && preview.dataUrl ? <iframe title={preview.name} src={preview.dataUrl} /> : null}
          {preview?.mode === "text" && preview.kind === "md" && preview.text ? (
            <div className="preview-markdown">
              <IncremarkContent content={preview.text} isFinished />
            </div>
          ) : null}
          {preview?.mode === "text" && preview.kind !== "md" && preview.text ? <pre>{preview.text}</pre> : null}
          {preview?.mode === "unsupported" ? <p className="preview-empty">{preview.summary}</p> : null}
        </ScrollArea.Viewport>
        <ScrollArea.Scrollbar className="scrollbar" orientation="vertical">
          <ScrollArea.Thumb className="scrollbar-thumb" />
        </ScrollArea.Scrollbar>
        </ScrollArea.Root>
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}

function ArtifactHistoryPanel({
  artifacts,
  sessions,
  onClose,
  onPreviewArtifact
}: {
  artifacts: ArtifactSummary[];
  sessions: ChatSession[];
  onClose: () => void;
  onPreviewArtifact: (artifactId: string) => void;
}): JSX.Element {
  return (
    <Dialog.Root
      open
      onOpenChange={(open) => {
        if (!open) onClose();
      }}
    >
      <Dialog.Portal>
        <Dialog.Overlay className="preview-backdrop" />
        <Dialog.Content className="artifact-history-panel">
          <header>
            <div>
              <Dialog.Title asChild>
                <strong>交付文件</strong>
              </Dialog.Title>
              <Dialog.Description>{artifacts.length ? `共 ${artifacts.length} 个生成产物` : "暂无生成产物"}</Dialog.Description>
            </div>
            <Dialog.Close asChild>
              <button type="button" aria-label="关闭交付文件">
                <X size={16} />
              </button>
            </Dialog.Close>
          </header>
          <ScrollArea.Root className="artifact-history-scroll">
            <ScrollArea.Viewport className="artifact-history-body">
              {artifacts.length === 0 ? (
                <div className="artifact-empty">
                  <Files size={26} />
                  <strong>还没有交付文件</strong>
                  <span>生成方案后，Markdown、Word、PDF 和图示会出现在这里。</span>
                </div>
              ) : (
                artifacts.map((artifact) => (
                  <div className="artifact-history-row" key={artifact.id}>
                    {isImageKind(artifact.kind) ? <Image size={16} /> : <File size={16} />}
                    <div className="artifact-history-main">
                      <strong>{artifact.name}</strong>
                      <span>
                        {artifact.kind.toUpperCase()} · {formatBytes(artifact.size)} ·{" "}
                        {resolveSessionTitle(sessions, artifact.sessionId)}
                      </span>
                    </div>
                    <time>{formatDateTime(artifact.createdAt)}</time>
                    <ArtifactActions artifactId={artifact.id} onPreview={() => onPreviewArtifact(artifact.id)} />
                  </div>
                ))
              )}
            </ScrollArea.Viewport>
            <ScrollArea.Scrollbar className="scrollbar" orientation="vertical">
              <ScrollArea.Thumb className="scrollbar-thumb" />
            </ScrollArea.Scrollbar>
          </ScrollArea.Root>
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}

function RenameSessionDialog({
  session,
  onClose,
  onRename
}: {
  session: ChatSession;
  onClose: () => void;
  onRename: (sessionId: string, title: string) => Promise<void>;
}): JSX.Element {
  const [title, setTitle] = useState(session.title);
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    setTitle(session.title);
    setError("");
  }, [session]);

  async function submit(event: FormEvent): Promise<void> {
    event.preventDefault();
    const nextTitle = title.trim();
    if (!nextTitle) {
      setError("对话名称不能为空");
      return;
    }

    setSaving(true);
    setError("");
    try {
      await onRename(session.id, nextTitle);
      onClose();
    } catch (renameError) {
      setError(renameError instanceof Error ? renameError.message : String(renameError));
    } finally {
      setSaving(false);
    }
  }

  return (
    <Dialog.Root
      open
      onOpenChange={(open) => {
        if (!open) onClose();
      }}
    >
      <Dialog.Portal>
        <Dialog.Overlay className="modal-backdrop" />
        <Dialog.Content className="modal-panel rename-panel">
          <form onSubmit={(event) => void submit(event)}>
            <Dialog.Title asChild>
              <h2>重命名对话</h2>
            </Dialog.Title>
            <Dialog.Description>给这次方案编制取一个更容易识别的名称。</Dialog.Description>
            <label>
              对话名称
              <input value={title} onChange={(event) => setTitle(event.target.value)} autoFocus />
            </label>
            {error ? <p className="form-error">{error}</p> : null}
            <footer>
              <Dialog.Close asChild>
                <button type="button" className="secondary-action">
                  取消
                </button>
              </Dialog.Close>
              <button type="submit" className="send-action" disabled={saving}>
                {saving ? <Loader2 size={14} className="spin" /> : <PencilLine size={14} />}
                保存
              </button>
            </footer>
          </form>
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}

function DeleteSessionDialog({
  session,
  onClose,
  onDelete
}: {
  session: ChatSession;
  onClose: () => void;
  onDelete: (sessionId: string) => Promise<void>;
}): JSX.Element {
  const [deleting, setDeleting] = useState(false);
  const [error, setError] = useState("");

  async function confirmDelete(): Promise<void> {
    setDeleting(true);
    setError("");
    try {
      await onDelete(session.id);
      onClose();
    } catch (deleteError) {
      setError(deleteError instanceof Error ? deleteError.message : String(deleteError));
      setDeleting(false);
    }
  }

  return (
    <AlertDialog.Root
      open
      onOpenChange={(open) => {
        if (!open && !deleting) onClose();
      }}
    >
      <AlertDialog.Portal>
        <AlertDialog.Overlay className="modal-backdrop" />
        <AlertDialog.Content className="modal-panel delete-panel">
          <AlertDialog.Title asChild>
            <h2>删除这个对话？</h2>
          </AlertDialog.Title>
          <AlertDialog.Description>
            将从侧边栏移除“{session.title}”及其聊天记录。已生成到磁盘的方案文件不会被删除。
          </AlertDialog.Description>
          {error ? <p className="form-error">{error}</p> : null}
          <footer>
            <AlertDialog.Cancel asChild>
              <button type="button" className="secondary-action" disabled={deleting}>
                取消
              </button>
            </AlertDialog.Cancel>
            <AlertDialog.Action asChild>
              <button
                type="button"
                className="danger-action"
                disabled={deleting}
                onClick={(event) => {
                  event.preventDefault();
                  void confirmDelete();
                }}
              >
                {deleting ? <Loader2 size={14} className="spin" /> : <Trash2 size={14} />}
                删除
              </button>
            </AlertDialog.Action>
          </footer>
        </AlertDialog.Content>
      </AlertDialog.Portal>
    </AlertDialog.Root>
  );
}

function Composer(props: {
  api: PwdSafeAgentApi;
  session?: ChatSession;
  value: string;
  attachments: AttachmentRef[];
  onChange: (value: string) => void;
  onAddAttachments: (attachments: AttachmentRef[]) => void;
  onRemoveAttachment: (id: string) => void;
  onSent: () => void;
}): JSX.Element {
  const canSend = Boolean(props.session && (props.value.trim() || props.attachments.length));
  const running = props.session?.status === "running";

  async function submit(event?: FormEvent): Promise<void> {
    event?.preventDefault();
    if (!props.session || !canSend) return;
    await props.api.chat.prompt({
      sessionId: props.session.id,
      message: props.value.trim(),
      attachments: props.attachments
    });
    props.onSent();
  }

  async function chooseFiles(): Promise<void> {
    if (!props.session) return;
    const picked = await props.api.attachment.pick({ sessionId: props.session.id, multiple: true });
    props.onAddAttachments(picked);
  }

  async function onPaste(event: ClipboardEvent<HTMLTextAreaElement>): Promise<void> {
    if (!props.session) return;
    const files = Array.from(event.clipboardData.files);
    if (files.length === 0) return;
    event.preventDefault();
    const payload = await Promise.all(
      files.map(async (file) => ({
        name: file.name || "clipboard-file",
        mimeType: file.type || "application/octet-stream",
        dataBase64: await fileToBase64(file)
      }))
    );
    const imported = await props.api.attachment.importClipboard({
      sessionId: props.session.id,
      files: payload
    });
    props.onAddAttachments(imported);
  }

  return (
    <form className="composer" onSubmit={(event) => void submit(event)}>
      {props.attachments.length ? (
        <div className="attachment-strip">
          {props.attachments.map((attachment) => (
            <span className="attachment-chip" key={attachment.id}>
              <File size={14} />
              {attachment.name}
              <button type="button" onClick={() => props.onRemoveAttachment(attachment.id)} aria-label="移除附件">
                <Trash2 size={13} />
              </button>
            </span>
          ))}
        </div>
      ) : null}
      <textarea
        value={props.value}
        onChange={(event) => props.onChange(event.target.value)}
        onPaste={(event) => void onPaste(event)}
        placeholder="输入系统背景、建设单位、网络拓扑、关键数据，或粘贴/选择文件..."
      />
      <div className="composer-actions">
        <TooltipButton label="选择 Word、PDF、图片或文本附件" className="icon-action" onClick={chooseFiles}>
          <Paperclip size={16} />
        </TooltipButton>
        <div className="paste-hint">
          <Upload size={14} />
          支持粘贴 Word、PDF、图片和文本资料
        </div>
        <div className="spacer" />
        {running ? (
          <TooltipButton
            label="停止当前 Agent 生成"
            className="secondary-action"
            onClick={() => void props.api.chat.abort(props.session!.id)}
          >
            <Square size={15} />
            停止
          </TooltipButton>
        ) : null}
        <button type="submit" className="send-action" disabled={!canSend || running}>
          <Play size={15} />
          发送
        </button>
      </div>
    </form>
  );
}

function SettingsPanel({
  api,
  settings,
  onClose,
  onSaved
}: {
  api: PwdSafeAgentApi;
  settings?: AppSettings;
  onClose: () => void;
  onSaved: (settings: AppSettings) => void;
}): JSX.Element {
  const [baseUrl, setBaseUrl] = useState(settings?.openai.baseUrl || "https://api.openai.com/v1");
  const [imageBaseUrl, setImageBaseUrl] = useState(settings?.openai.imageBaseUrl || "");
  const [chatModel, setChatModel] = useState(settings?.openai.chatModel || "gpt-5.5");
  const [imageModel, setImageModel] = useState(settings?.openai.imageModel || "gpt-image-2");
  const [imageSize, setImageSize] = useState(settings?.openai.imageSize || "1536x1024");
  const [imageQuality, setImageQuality] = useState(settings?.openai.imageQuality || "high");
  const [autoImageGeneration, setAutoImageGeneration] = useState(settings?.openai.autoImageGeneration ?? true);
  const [autoPdfExport, setAutoPdfExport] = useState(settings?.document.autoPdfExport ?? false);
  const [libreOfficePath, setLibreOfficePath] = useState(settings?.document.libreOfficePath || "");
  const [timeout, setTimeoutValue] = useState(settings?.openai.requestTimeoutMs || 120000);
  const [maxTokens, setMaxTokens] = useState(settings?.openai.maxOutputTokens || 16000);
  const [execBashEnabled, setExecBashEnabled] = useState(settings?.agent.execBashEnabled ?? false);
  const [apiKey, setApiKey] = useState("");
  const [imageApiKey, setImageApiKey] = useState("");

  async function save(): Promise<void> {
    const next = await api.settings.save({
      openai: {
        baseUrl,
        imageBaseUrl,
        chatModel,
        imageModel,
        imageSize,
        imageQuality,
        autoImageGeneration,
        requestTimeoutMs: timeout,
        maxOutputTokens: maxTokens,
        apiKey: apiKey.trim() || undefined,
        imageApiKey: imageApiKey.trim() || undefined
      },
      document: {
        autoPdfExport,
        libreOfficePath
      },
      agent: {
        execBashEnabled
      }
    });
    onSaved(next);
    onClose();
  }

  return (
    <Dialog.Root
      open
      onOpenChange={(open) => {
        if (!open) onClose();
      }}
    >
      <Dialog.Portal>
        <Dialog.Overlay className="settings-backdrop" />
        <Dialog.Content className="settings-panel">
        <header>
          <Dialog.Title asChild>
            <h2>设置</h2>
          </Dialog.Title>
          <Dialog.Close asChild>
            <button type="button">关闭</button>
          </Dialog.Close>
        </header>
        <Dialog.Description className="settings-description">
          模型配置会写入本地 `.env.local`，API Key 只保存在 Electron 后端可读取的环境文件中。
        </Dialog.Description>
        <Tabs.Root className="settings-tabs" defaultValue="model">
          <Tabs.List className="tabs-list" aria-label="设置分组">
            <Tabs.Trigger className="tabs-trigger" value="model">
              模型
            </Tabs.Trigger>
            <Tabs.Trigger className="tabs-trigger" value="document">
              文档
            </Tabs.Trigger>
            <Tabs.Trigger className="tabs-trigger" value="advanced">
              高级
            </Tabs.Trigger>
          </Tabs.List>
          <ScrollArea.Root className="settings-scroll">
            <ScrollArea.Viewport className="settings-viewport">
              <Tabs.Content className="settings-form" value="model">
                <label>
                  API Key
                  <input
                    value={apiKey}
                    onChange={(event) => setApiKey(event.target.value)}
                    placeholder={settings?.openai.apiKeyConfigured ? "已配置，留空保持不变" : "sk-..."}
                  />
                </label>
                <label>
                  生图 API Key
                  <input
                    value={imageApiKey}
                    onChange={(event) => setImageApiKey(event.target.value)}
                    placeholder={settings?.openai.imageApiKeyConfigured ? "已配置，留空保持不变" : "留空沿用 API Key"}
                  />
                </label>
                <label>
                  Base URL
                  <input value={baseUrl} onChange={(event) => setBaseUrl(event.target.value)} />
                </label>
                <label>
                  生图 Base URL
                  <input
                    value={imageBaseUrl}
                    onChange={(event) => setImageBaseUrl(event.target.value)}
                    placeholder="留空沿用 Base URL"
                  />
                </label>
                <label>
                  Chat 模型
                  <input value={chatModel} onChange={(event) => setChatModel(event.target.value)} />
                </label>
                <label>
                  生图模型
                  <input value={imageModel} onChange={(event) => setImageModel(event.target.value)} />
                </label>
                <label>
                  生图尺寸
                  <input value={imageSize} onChange={(event) => setImageSize(event.target.value)} />
                </label>
                <label>
                  生图质量
                  <input value={imageQuality} onChange={(event) => setImageQuality(event.target.value)} />
                </label>
              </Tabs.Content>

              <Tabs.Content className="settings-form" value="document">
                <SwitchRow checked={autoImageGeneration} onCheckedChange={setAutoImageGeneration}>
                  生成方案时自动生成架构图和流程图
                </SwitchRow>
                <SwitchRow checked={autoPdfExport} onCheckedChange={setAutoPdfExport}>
                  生成 Word 后自动导出 PDF
                </SwitchRow>
                <label>
                  LibreOffice 路径
                  <input
                    value={libreOfficePath}
                    onChange={(event) => setLibreOfficePath(event.target.value)}
                    placeholder="留空自动查找 soffice/libreoffice"
                  />
                </label>
                <p className="settings-note">
                  PDF 导出依赖本机 LibreOffice。未配置或未安装时，Agent 会保留 Word 方案并在工具调用中说明原因。
                </p>
              </Tabs.Content>

              <Tabs.Content className="settings-form" value="advanced">
                <label>
                  请求超时 ms
                  <input
                    type="number"
                    value={timeout}
                    onChange={(event) => setTimeoutValue(Number(event.target.value))}
                  />
                </label>
                <label>
                  最大输出 tokens
                  <input
                    type="number"
                    value={maxTokens}
                    onChange={(event) => setMaxTokens(Number(event.target.value))}
                  />
                </label>
                <SwitchRow checked={execBashEnabled} onCheckedChange={setExecBashEnabled}>
                  允许 Agent 使用 exec_bash 执行命令
                </SwitchRow>
                <p className="settings-note">
                  开启后模型可请求命令执行工具。建议仅在可信任务中使用，并保持对工具调用过程的人工观察。
                </p>
              </Tabs.Content>
            </ScrollArea.Viewport>
            <ScrollArea.Scrollbar className="scrollbar" orientation="vertical">
              <ScrollArea.Thumb className="scrollbar-thumb" />
            </ScrollArea.Scrollbar>
          </ScrollArea.Root>
        </Tabs.Root>
        <footer>
          <span>{settings?.runtime.envFilePath || ".env.local"}</span>
          <button className="send-action" onClick={() => void save()}>
            保存
          </button>
        </footer>
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}

function TooltipButton({
  label,
  className,
  onClick,
  children
}: {
  label: string;
  className: string;
  onClick: () => void;
  children: ReactNode;
}): JSX.Element {
  return (
    <Tooltip.Root>
      <Tooltip.Trigger asChild>
        <button type="button" className={className} onClick={onClick}>
          {children}
        </button>
      </Tooltip.Trigger>
      <Tooltip.Portal>
        <Tooltip.Content className="tooltip-content" sideOffset={7}>
          {label}
          <Tooltip.Arrow className="tooltip-arrow" />
        </Tooltip.Content>
      </Tooltip.Portal>
    </Tooltip.Root>
  );
}

function SwitchRow({
  checked,
  onCheckedChange,
  children
}: {
  checked: boolean;
  onCheckedChange: (checked: boolean) => void;
  children: string;
}): JSX.Element {
  return (
    <div className="switch-row">
      <span>{children}</span>
      <Switch.Root className="switch-control" checked={checked} onCheckedChange={onCheckedChange} aria-label={children}>
        <Switch.Thumb className="switch-thumb" />
      </Switch.Root>
    </div>
  );
}

function statusLabel(status: ChatSession["status"]): string {
  const map = {
    idle: "空闲",
    running: "生成中",
    completed: "已完成",
    failed: "失败"
  };
  return map[status];
}

function isImageKind(kind: ArtifactKind): boolean {
  return kind === "png" || kind === "jpg" || kind === "jpeg" || kind === "webp" || kind === "svg";
}

function resolveSessionTitle(sessions: ChatSession[], sessionId?: string): string {
  if (!sessionId) return "历史会话";
  return sessions.find((session) => session.id === sessionId)?.title || "历史会话";
}

function formatBytes(size: number): string {
  if (!Number.isFinite(size) || size <= 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  let value = size;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value >= 10 || unitIndex === 0 ? value.toFixed(0) : value.toFixed(1)} ${units[unitIndex]}`;
}

function formatDateTime(value: string): string {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return date.toLocaleString("zh-CN", {
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  });
}

async function fileToBase64(file: File): Promise<string> {
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }
  return btoa(binary);
}
