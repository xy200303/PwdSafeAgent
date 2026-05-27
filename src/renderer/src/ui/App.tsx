import {
  lazy,
  Suspense,
  useEffect,
  useRef,
  useState,
  type ClipboardEvent,
  type DragEvent,
  type FormEvent,
  type KeyboardEvent,
  type JSX
} from "react";
import * as AlertDialog from "@radix-ui/react-alert-dialog";
import * as Collapsible from "@radix-ui/react-collapsible";
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
  ChevronDown,
  Circle,
  Eye,
  File,
  Files,
  FolderOpen,
  Image,
  ListChecks,
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
import {
  DRAFT_SECTION_PARALLELISM_DEFAULT,
  DRAFT_SECTION_PARALLELISM_MAX,
  DRAFT_SECTION_PARALLELISM_MIN,
  IMAGE_GENERATION_REQUEST_TIMEOUT_DEFAULT_MS,
  IMAGE_GENERATION_PARALLELISM_DEFAULT,
  IMAGE_GENERATION_PARALLELISM_MAX,
  IMAGE_GENERATION_PARALLELISM_MIN,
  clampDraftSectionParallelism,
  clampImageGenerationParallelism
} from "../../../shared/types";
import type {
  AppSettings,
  ArtifactKind,
  ArtifactPreview,
  ArtifactSummary,
  AttachmentRef,
  ChatSession,
  PwdSafeAgentApi,
  RuntimeCheckResult,
  RuntimeCommandCheck,
  SchemeProgressItem,
  SchemeProgressSection,
  SchemeSectionStatus,
  StreamItem
} from "../../../shared/types";
import { getErrorMessage, resolvePwdSafeAgentApi } from "../bridge";
import { groupStreamItemsForDisplay, orderStreamItemsForDisplay, type ProcessStreamItem } from "../streamOrdering";
import { filesToAttachmentPayload } from "./attachmentPayload";

const LazyArtifactPreviewPanel = lazy(() => import("./ArtifactPreviewPanel"));

export function App(): JSX.Element {
  const dispatch = useDispatch<AppDispatch>();
  const { sessions, artifacts, currentSessionId, composer, pendingAttachments, settings, settingsOpen, artifactsOpen } = useSelector(
    (state: RootState) => state.chat
  );
  const [preview, setPreview] = useState<ArtifactPreview | undefined>();
  const [previewError, setPreviewError] = useState("");
  const [appError, setAppError] = useState("");
  const [renameTarget, setRenameTarget] = useState<ChatSession | undefined>();
  const [deleteTarget, setDeleteTarget] = useState<ChatSession | undefined>();
  const currentSession = sessions.find((session) => session.id === currentSessionId);
  const schemeProgress = findLatestSchemeProgressItem(currentSession);
  const bridge = resolvePwdSafeAgentApi(window.pwdSafeAgent);
  const api = bridge.api;

  useEffect(() => {
    if (!api) return;
    let unsubscribe: (() => void) | undefined;
    void bootstrap(dispatch, api).catch((error: unknown) => setAppError(`启动初始化失败：${getErrorMessage(error)}`));
    try {
      unsubscribe = api.events.subscribe((event) => {
        dispatch(applyRendererEvent(event));
      });
    } catch (error) {
      setAppError(`事件订阅失败：${getErrorMessage(error)}`);
    }
    return unsubscribe;
  }, [api, dispatch]);

  async function createConversation(): Promise<void> {
    if (!api) return;
    try {
      const session = await api.session.create();
      dispatch(upsertSession(session));
      dispatch(setCurrentSession(session.id));
    } catch (error) {
      setAppError(`新建对话失败：${getErrorMessage(error)}`);
    }
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
      setPreviewError(getErrorMessage(error));
    }
  }

  async function openArtifactHistory(): Promise<void> {
    if (!api) return;
    dispatch(setArtifacts(await api.artifact.list()));
    dispatch(setArtifactsOpen(true));
  }

  if (!api) {
    return <MissingBridge detail={bridge.error} />;
  }

  return (
    <Tooltip.Provider delayDuration={350}>
      <div className={schemeProgress ? "app-shell with-progress" : "app-shell"}>
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
          <MessagePane
            api={api}
            session={currentSession}
            onPreviewArtifact={(artifactId) => void previewArtifact(artifactId)}
            onError={(message) => setAppError(message)}
          />
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
            onError={(message) => setAppError(message)}
          />
        </main>
        {schemeProgress ? <SchemeProgressPanel item={schemeProgress} /> : null}
        {settingsOpen ? (
          <SettingsPanel
            api={api}
            settings={settings}
            onClose={() => dispatch(setSettingsOpen(false))}
            onSaved={(next) => dispatch(setSettings(next))}
            onError={(message) => setAppError(message)}
          />
        ) : null}
        {preview || previewError ? (
          <Suspense fallback={<PreviewLoadingPanel />}>
            <LazyArtifactPreviewPanel
              preview={preview}
              error={previewError}
              onClose={() => {
                setPreview(undefined);
                setPreviewError("");
              }}
            />
          </Suspense>
        ) : null}
        {artifactsOpen ? (
          <ArtifactHistoryPanel
            artifacts={artifacts}
            sessions={sessions}
            onClose={() => dispatch(setArtifactsOpen(false))}
            api={api}
            onError={(message) => setAppError(message)}
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
        {appError ? <AppNotice message={appError} onClose={() => setAppError("")} /> : null}
      </div>
    </Tooltip.Provider>
  );
}

function MissingBridge({ detail }: { detail?: string }): JSX.Element {
  return (
    <div className="bridge-missing">
      <strong>Electron 预加载桥未就绪</strong>
      <span>{detail || "请通过 `npm run dev` 或打包后的 Electron 应用启动，不要直接在浏览器打开渲染页。"}</span>
      <small>如果仍然出现该提示，请重启开发服务；preload 会注入 `window.pwdSafeAgent`。</small>
    </div>
  );
}

function AppNotice({ message, onClose }: { message: string; onClose: () => void }): JSX.Element {
  return (
    <div className="app-notice" role="status">
      <span>{message}</span>
      <button type="button" onClick={onClose} aria-label="关闭提示">
        <X size={14} />
      </button>
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
          模板：密码应用方案.docx · Agent：内置 Pi Agent · 模型：{settings?.openai.chatModel || "未加载"}
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
  api,
  session,
  onPreviewArtifact,
  onError
}: {
  api: PwdSafeAgentApi;
  session?: ChatSession;
  onPreviewArtifact: (artifactId: string) => void;
  onError: (message: string) => void;
}): JSX.Element {
  const scrollRef = useRef<HTMLDivElement>(null);
  const orderedItems = session ? orderStreamItemsForDisplay(session.items) : [];
  const displayBlocks = session ? groupStreamItemsForDisplay(session.items) : [];
  const shouldShowPendingThinking = Boolean(
    session?.status === "running" && !orderedItems.some(isUnfinishedAssistantMessage)
  );

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
          <>
            {displayBlocks.map((block) => (
              block.kind === "message" ? (
                <StreamRow
                  key={block.id}
                  api={api}
                  item={block.item}
                  onPreviewArtifact={onPreviewArtifact}
                  onError={onError}
                />
              ) : (
                <ProcessCard
                  key={block.id}
                  api={api}
                  items={block.items}
                  onPreviewArtifact={onPreviewArtifact}
                  onError={onError}
                />
              )
            ))}
            {shouldShowPendingThinking ? <PendingAssistantRow /> : null}
          </>
        )}
      </ScrollArea.Viewport>
      <ScrollArea.Scrollbar className="scrollbar" orientation="vertical">
        <ScrollArea.Thumb className="scrollbar-thumb" />
      </ScrollArea.Scrollbar>
    </ScrollArea.Root>
  );
}

interface SchemeProgressChapter {
  root: SchemeProgressSection;
  sections: SchemeProgressSection[];
  total: number;
  completed: number;
  failed: number;
  status: SchemeSectionStatus;
}

function SchemeProgressPanel({ item }: { item: SchemeProgressItem }): JSX.Element {
  const chapters = buildSchemeProgressChapters(item.sections);
  const percent = item.total > 0 ? Math.round((item.completed / item.total) * 100) : 0;
  const drafted = item.drafted ?? item.sections.filter((section) => section.status === "drafted").length;

  return (
    <aside className="scheme-progress-panel">
      <header className="scheme-progress-header">
        <ListChecks size={17} />
        <div>
          <span>章节进度</span>
          <strong>{item.artifactName || item.title}</strong>
        </div>
        <span className={`scheme-progress-pill ${item.status}`}>
          <SchemeStatusIcon status={item.status} />
          {schemeProgressStatusLabel(item.status)}
        </span>
      </header>

      <div className="scheme-progress-summary">
        <div className="scheme-progress-numbers">
          <strong>
            {item.completed}/{item.total}
          </strong>
          <span>{percent}%</span>
        </div>
        <div className="scheme-progress-track" aria-label={`章节完成 ${percent}%`}>
          <span style={{ width: `${percent}%` }} />
        </div>
        <p>
          {item.detail || "等待按模板章节写入"}
          {drafted ? ` · 已起草 ${drafted}` : ""}
          {item.failed ? ` · 失败 ${item.failed}` : ""}
        </p>
      </div>

      <ScrollArea.Root className="scheme-progress-scroll">
        <ScrollArea.Viewport className="scheme-progress-body">
          {chapters.map((chapter, index) => (
            <SchemeChapterProgress
              key={chapter.root.id}
              chapter={chapter}
              defaultOpen={chapter.status === "running" || chapter.failed > 0 || chapter.completed > 0 || index === 0}
            />
          ))}
        </ScrollArea.Viewport>
        <ScrollArea.Scrollbar className="scrollbar" orientation="vertical">
          <ScrollArea.Thumb className="scrollbar-thumb" />
        </ScrollArea.Scrollbar>
      </ScrollArea.Root>
    </aside>
  );
}

function SchemeChapterProgress({
  chapter,
  defaultOpen
}: {
  chapter: SchemeProgressChapter;
  defaultOpen: boolean;
}): JSX.Element {
  const childSections = chapter.sections.filter((section) => section.id !== chapter.root.id);
  const percent = chapter.total > 0 ? Math.round((chapter.completed / chapter.total) * 100) : 0;

  return (
    <Collapsible.Root className={`scheme-chapter ${chapter.status}`} defaultOpen={defaultOpen}>
      <Collapsible.Trigger asChild>
        <button type="button" className="scheme-chapter-trigger">
          <SchemeStatusIcon status={chapter.status} />
          <span className="scheme-chapter-main">
            <strong>
              {chapter.root.number} {chapter.root.title}
            </strong>
            <span>
              {chapter.completed}/{chapter.total} · {percent}%
              {chapter.failed ? ` · 失败 ${chapter.failed}` : ""}
            </span>
          </span>
          <ChevronDown size={14} />
        </button>
      </Collapsible.Trigger>
      <Collapsible.Content className="scheme-section-list">
        {childSections.length ? (
          childSections.map((section) => <SchemeSectionProgressRow key={section.id} section={section} />)
        ) : (
          <SchemeSectionProgressRow section={chapter.root} />
        )}
      </Collapsible.Content>
    </Collapsible.Root>
  );
}

function SchemeSectionProgressRow({ section }: { section: SchemeProgressSection }): JSX.Element {
  const depth = Math.min(Math.max(section.headingLevel - 2, 0), 4);
  return (
    <div className={`scheme-section-row ${section.status} depth-${depth}`}>
      <SchemeStatusIcon status={section.status} />
      <span className="scheme-section-main">
        <strong>
          {section.number} {section.title}
        </strong>
        <small>{formatSchemeSectionDetail(section)}</small>
      </span>
    </div>
  );
}

function SchemeStatusIcon({ status }: { status: SchemeSectionStatus | SchemeProgressItem["status"] }): JSX.Element {
  if (status === "running" || status === "drafting") return <Loader2 size={13} className="spin" />;
  if (status === "drafted") return <PencilLine size={13} />;
  if (status === "completed") return <CheckCircle2 size={13} />;
  if (status === "failed") return <XCircle size={13} />;
  return <Circle size={13} />;
}

function buildSchemeProgressChapters(sections: SchemeProgressSection[]): SchemeProgressChapter[] {
  const roots = sections.filter((section) => section.headingLevel === 1);
  const chapterRoots = roots.length ? roots : sections.slice(0, 1);
  return chapterRoots.map((root) => {
    const members = sections.filter((section) => section.id === root.id || section.number.startsWith(`${root.number}.`));
    const completed = members.filter((section) => section.status === "completed").length;
    const failed = members.filter((section) => section.status === "failed").length;
    return {
      root,
      sections: members,
      total: members.length,
      completed,
      failed,
      status: resolveSchemeChapterStatus(members, completed, failed)
    };
  });
}

function resolveSchemeChapterStatus(
  sections: SchemeProgressSection[],
  completed: number,
  failed: number
): SchemeSectionStatus {
  if (sections.some((section) => section.status === "running" || section.status === "drafting")) return "running";
  if (failed > 0) return "failed";
  if (sections.length > 0 && completed === sections.length) return "completed";
  if (sections.some((section) => section.status === "drafted") || completed > 0) return "running";
  return "pending";
}

function formatSchemeSectionDetail(section: SchemeProgressSection): string {
  if (section.detail) return section.detail;
  const refs = [
    section.relatedTables?.length ? `表格 ${section.relatedTables.length}` : "",
    section.relatedFigures?.length ? `图示 ${section.relatedFigures.length}` : ""
  ].filter(Boolean);
  if (refs.length) return refs.join(" · ");
  return section.writingHint || schemeSectionStatusLabel(section.status);
}

function isUnfinishedAssistantMessage(item: StreamItem): boolean {
  return item.kind === "message" && item.role === "assistant" && !item.isFinished;
}

function PendingAssistantRow(): JSX.Element {
  return (
    <article className="message-row assistant thinking-row">
      <div className="avatar">
        <Bot size={15} />
      </div>
      <div className="message-body">
        <ThinkingIndicator />
      </div>
    </article>
  );
}

function ProcessCard({
  api,
  items,
  onPreviewArtifact,
  onError
}: {
  api: PwdSafeAgentApi;
  items: ProcessStreamItem[];
  onPreviewArtifact: (artifactId: string) => void;
  onError: (message: string) => void;
}): JSX.Element {
  const status = resolveProcessCardStatus(items);

  return (
    <Collapsible.Root className={`process-card ${status}`} defaultOpen>
      <div className="process-card-toolbar">
        <ProcessCardStatusIcon status={status} />
        <span className="process-card-title">过程记录</span>
        <span className="process-card-count">{formatProcessCardCounts(items)}</span>
        <span className="process-card-summary">{formatProcessCardSummary(items)}</span>
        <Collapsible.Trigger asChild>
          <button type="button" className="process-card-toggle" aria-label="展开或折叠过程记录">
            <ChevronDown size={14} />
          </button>
        </Collapsible.Trigger>
      </div>
      <Collapsible.Content className="process-card-body">
        {items.map((processItem) => (
          <StreamRow
            key={processItem.id}
            api={api}
            item={processItem}
            onPreviewArtifact={onPreviewArtifact}
            onError={onError}
            embedded
          />
        ))}
      </Collapsible.Content>
    </Collapsible.Root>
  );
}

function ProcessCardStatusIcon({ status }: { status: "running" | "success" | "failed" }): JSX.Element {
  if (status === "running") return <Loader2 size={14} className="spin" />;
  if (status === "failed") return <XCircle size={14} />;
  return <CheckCircle2 size={14} />;
}

function resolveProcessCardStatus(items: ProcessStreamItem[]): "running" | "success" | "failed" {
  if (items.some((item) => item.kind === "tool" && item.status === "failed")) return "failed";
  if (items.some((item) => item.kind === "tool" && item.status === "running")) return "running";
  return "success";
}

function formatProcessCardCounts(items: ProcessStreamItem[]): string {
  const toolCount = items.filter((item) => item.kind === "tool").length;
  const fileCount = items.filter((item) => item.kind === "file").length;
  const stageCount = items.filter((item) => item.kind === "stage").length;
  return [
    toolCount ? `${toolCount} 个工具` : "",
    stageCount ? `${stageCount} 条过程` : "",
    fileCount ? `${fileCount} 个文件` : ""
  ].filter(Boolean).join(" · ");
}

function formatProcessCardSummary(items: ProcessStreamItem[]): string {
  const latest = items.at(-1);
  if (!latest) return "等待执行";
  if (latest.kind === "file") return `已生成 ${latest.name}`;
  if (latest.kind === "stage") return [latest.title, latest.detail].filter(Boolean).join(" · ");
  return latest.summary || `工具 ${latest.toolName}`;
}

function StreamRow({
  api,
  item,
  onPreviewArtifact,
  onError,
  embedded = false
}: {
  api: PwdSafeAgentApi;
  item: StreamItem;
  onPreviewArtifact: (artifactId: string) => void;
  onError: (message: string) => void;
  embedded?: boolean;
}): JSX.Element {
  if (item.kind === "message") {
    const isUser = item.role === "user";
    return (
      <article className={`message-row ${item.role}`}>
        {!isUser ? <div className="avatar"><Bot size={15} /></div> : null}
        <div className="message-body">
          {item.role === "assistant" ? (
            item.content.trim() ? (
              <IncremarkContent content={item.content} isFinished={item.isFinished} />
            ) : item.isFinished ? (
              null
            ) : (
              <ThinkingIndicator />
            )
          ) : (
            <p>{item.content}</p>
          )}
        </div>
        {isUser ? <div className="avatar"><User size={15} /></div> : null}
      </article>
    );
  }

  if (item.kind === "tool") {
    return <ToolRow item={item} embedded={embedded} />;
  }

  if (item.kind === "file") {
    return (
      <div className={embedded ? "file-row embedded" : "file-row"}>
        {isImageKind(item.fileKind) ? <Image size={16} /> : <File size={16} />}
        <span>{item.name}</span>
        <span className="file-kind">{item.fileKind}</span>
        <ArtifactActions
          api={api}
          artifactId={item.artifactId}
          onPreview={() => onPreviewArtifact(item.artifactId)}
          onError={onError}
        />
      </div>
    );
  }

  return (
    <div className={embedded ? "stage-row embedded" : "stage-row"}>
      <span>{item.title}</span>
      {item.detail ? <small>{item.detail}</small> : null}
    </div>
  );
}

function ThinkingIndicator(): JSX.Element {
  return (
    <div className="thinking-indicator" role="status" aria-live="polite" aria-label="Agent 正在思考">
      <span className="thinking-label">正在思考</span>
      <span className="thinking-dots" aria-hidden="true">
        <span />
        <span />
        <span />
      </span>
    </div>
  );
}

function ToolRow({ item, embedded = false }: { item: Extract<StreamItem, { kind: "tool" }>; embedded?: boolean }): JSX.Element {
  const hasDetails = Boolean(item.inputPreview || item.outputPreview || item.errorPreview);

  return (
    <Collapsible.Root className={`tool-row ${item.status}${embedded ? " embedded" : ""}`}>
      <div className="tool-row-main">
        {item.status === "running" ? (
          <Loader2 size={14} className="spin" />
        ) : item.status === "failed" ? (
          <XCircle size={14} />
        ) : (
          <CheckCircle2 size={14} />
        )}
        <span className="tool-name">{item.toolName}</span>
        <span className="tool-summary">{item.summary}</span>
        {hasDetails ? (
          <Collapsible.Trigger asChild>
            <button type="button" className="tool-detail-trigger">
              详情
              <ChevronDown size={13} />
            </button>
          </Collapsible.Trigger>
        ) : null}
      </div>
      {hasDetails ? (
        <Collapsible.Content className="tool-detail">
          {item.inputPreview ? <ToolDetailBlock title="输入" content={item.inputPreview} /> : null}
          {item.outputPreview ? <ToolDetailBlock title="输出" content={item.outputPreview} /> : null}
          {item.errorPreview ? <ToolDetailBlock title="错误" content={item.errorPreview} /> : null}
        </Collapsible.Content>
      ) : null}
    </Collapsible.Root>
  );
}

function ToolDetailBlock({ title, content }: { title: string; content: string }): JSX.Element {
  return (
    <section className="tool-detail-block">
      <span>{title}</span>
      <pre>{formatToolPreview(content)}</pre>
    </section>
  );
}

function ArtifactActions({
  api,
  artifactId,
  onPreview,
  onError
}: {
  api: PwdSafeAgentApi;
  artifactId: string;
  onPreview: () => void;
  onError: (message: string) => void;
}): JSX.Element {
  async function openArtifact(): Promise<void> {
    try {
      await api.artifact.open(artifactId);
    } catch (error) {
      onError(`打开文件失败：${getErrorMessage(error)}`);
    }
  }

  async function revealArtifact(): Promise<void> {
    try {
      await api.artifact.reveal(artifactId);
    } catch (error) {
      onError(`定位文件失败：${getErrorMessage(error)}`);
    }
  }

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
              onSelect={() => void openArtifact()}
            >
              <File size={14} />
              系统打开
            </DropdownMenu.Item>
            <DropdownMenu.Item
              className="dropdown-item"
              onSelect={() => void revealArtifact()}
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

function PreviewLoadingPanel(): JSX.Element {
  return (
    <div className="preview-loading-panel" role="status">
      <Loader2 size={16} className="spin" />
      正在加载文件预览...
    </div>
  );
}

function ArtifactHistoryPanel({
  artifacts,
  sessions,
  onClose,
  api,
  onError,
  onPreviewArtifact
}: {
  artifacts: ArtifactSummary[];
  sessions: ChatSession[];
  onClose: () => void;
  api: PwdSafeAgentApi;
  onError: (message: string) => void;
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
                    <ArtifactActions
                      api={api}
                      artifactId={artifact.id}
                      onPreview={() => onPreviewArtifact(artifact.id)}
                      onError={onError}
                    />
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
  onError: (message: string) => void;
}): JSX.Element {
  const canSend = Boolean(props.session && (props.value.trim() || props.attachments.length));
  const running = props.session?.status === "running";
  const [draggingFiles, setDraggingFiles] = useState(false);

  async function submit(event?: FormEvent): Promise<void> {
    event?.preventDefault();
    if (!props.session || !canSend) return;
    try {
      await props.api.chat.prompt({
        sessionId: props.session.id,
        message: props.value.trim(),
        attachments: props.attachments
      });
      props.onSent();
    } catch (error) {
      props.onError(`发送失败：${getErrorMessage(error)}`);
    }
  }

  async function chooseFiles(): Promise<void> {
    if (!props.session) return;
    try {
      const picked = await props.api.attachment.pick({ sessionId: props.session.id, multiple: true });
      props.onAddAttachments(picked);
    } catch (error) {
      props.onError(`选择附件失败：${getErrorMessage(error)}`);
    }
  }

  async function importFiles(files: File[], actionLabel: string, source: "clipboard" | "drop"): Promise<void> {
    if (!props.session) return;
    if (files.length === 0) return;
    try {
      const payload = await filesToAttachmentPayload(files);
      const imported = await props.api.attachment.importClipboard({
        sessionId: props.session.id,
        source,
        files: payload
      });
      props.onAddAttachments(imported);
    } catch (error) {
      props.onError(`${actionLabel}附件失败：${getErrorMessage(error)}`);
    }
  }

  async function onPaste(event: ClipboardEvent<HTMLTextAreaElement>): Promise<void> {
    const files = Array.from(event.clipboardData.files);
    if (files.length === 0) return;
    event.preventDefault();
    await importFiles(files, "粘贴", "clipboard");
  }

  function onDragOver(event: DragEvent<HTMLFormElement>): void {
    if (!event.dataTransfer.types.includes("Files")) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = props.session ? "copy" : "none";
    setDraggingFiles(true);
  }

  function onDragLeave(event: DragEvent<HTMLFormElement>): void {
    if (event.currentTarget.contains(event.relatedTarget as Node | null)) return;
    setDraggingFiles(false);
  }

  async function onDrop(event: DragEvent<HTMLFormElement>): Promise<void> {
    if (!event.dataTransfer.types.includes("Files")) return;
    event.preventDefault();
    setDraggingFiles(false);
    await importFiles(Array.from(event.dataTransfer.files), "拖入", "drop");
  }

  function onKeyDown(event: KeyboardEvent<HTMLTextAreaElement>): void {
    if (event.key !== "Enter" || event.shiftKey || event.nativeEvent.isComposing) return;
    event.preventDefault();
    void submit();
  }

  return (
    <form
      className={draggingFiles ? "composer dragging-files" : "composer"}
      onSubmit={(event) => void submit(event)}
      onDragOver={onDragOver}
      onDragLeave={onDragLeave}
      onDrop={(event) => void onDrop(event)}
    >
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
        onKeyDown={onKeyDown}
        onPaste={(event) => void onPaste(event)}
        placeholder="输入系统背景、建设单位、网络拓扑、关键数据，或粘贴/选择文件..."
      />
      {draggingFiles ? (
        <div className="composer-drop-indicator">
          <Upload size={18} />
          松开即可添加附件
        </div>
      ) : null}
      <div className="composer-actions">
        <TooltipButton label="选择 Word、PDF、图片或文本附件" className="icon-action" onClick={chooseFiles}>
          <Paperclip size={16} />
        </TooltipButton>
        <div className="paste-hint">
          <Upload size={14} />
          支持拖拽、粘贴 Word、PDF、图片和文本资料
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
  onSaved,
  onError
}: {
  api: PwdSafeAgentApi;
  settings?: AppSettings;
  onClose: () => void;
  onSaved: (settings: AppSettings) => void;
  onError: (message: string) => void;
}): JSX.Element {
  const [baseUrl, setBaseUrl] = useState(settings?.openai.baseUrl || "https://api.openai.com/v1");
  const [imageBaseUrl, setImageBaseUrl] = useState(settings?.openai.imageBaseUrl || "");
  const [chatModel, setChatModel] = useState(settings?.openai.chatModel || "gpt-5.5");
  const [imageModel, setImageModel] = useState(settings?.openai.imageModel || "gpt-image-2");
  const [imageSize, setImageSize] = useState(settings?.openai.imageSize || "1536x1024");
  const [imageQuality, setImageQuality] = useState(settings?.openai.imageQuality || "high");
  const [autoImageGeneration, setAutoImageGeneration] = useState(settings?.openai.autoImageGeneration ?? true);
  const [thinkingEnabled, setThinkingEnabled] = useState(settings?.openai.thinkingEnabled ?? true);
  const [reasoningEffort, setReasoningEffort] = useState(settings?.openai.reasoningEffort || "");
  const [autoPdfExport, setAutoPdfExport] = useState(settings?.document.autoPdfExport ?? false);
  const [libreOfficePath, setLibreOfficePath] = useState(settings?.document.libreOfficePath || "");
  const [timeout, setTimeoutValue] = useState(settings?.openai.requestTimeoutMs || 120000);
  const [imageTimeout, setImageTimeoutValue] = useState(
    settings?.openai.imageRequestTimeoutMs || IMAGE_GENERATION_REQUEST_TIMEOUT_DEFAULT_MS
  );
  const [maxTokens, setMaxTokens] = useState(settings?.openai.maxOutputTokens || 16000);
  const [execBashEnabled, setExecBashEnabled] = useState(settings?.agent.execBashEnabled ?? true);
  const [draftSectionParallelism, setDraftSectionParallelism] = useState(
    settings ? clampDraftSectionParallelism(settings.agent.draftSectionParallelism) : DRAFT_SECTION_PARALLELISM_DEFAULT
  );
  const [imageGenerationParallelism, setImageGenerationParallelism] = useState(
    settings
      ? clampImageGenerationParallelism(settings.agent.imageGenerationParallelism)
      : IMAGE_GENERATION_PARALLELISM_DEFAULT
  );
  const [apiKey, setApiKey] = useState("");
  const [imageApiKey, setImageApiKey] = useState("");
  const [runtimeCheck, setRuntimeCheck] = useState<RuntimeCheckResult | undefined>();
  const [checkingRuntime, setCheckingRuntime] = useState(false);
  const [saving, setSaving] = useState(false);

  async function save(): Promise<void> {
    setSaving(true);
    try {
      const next = await api.settings.save({
        openai: {
          baseUrl,
          imageBaseUrl,
          chatModel,
          imageModel,
          imageSize,
          imageQuality,
          autoImageGeneration,
          thinkingEnabled,
          reasoningEffort,
          requestTimeoutMs: timeout,
          imageRequestTimeoutMs: imageTimeout,
          maxOutputTokens: maxTokens,
          apiKey: apiKey.trim() || undefined,
          imageApiKey: imageApiKey.trim() || undefined
        },
        document: {
          autoPdfExport,
          libreOfficePath
        },
        agent: {
          execBashEnabled,
          draftSectionParallelism: clampDraftSectionParallelism(draftSectionParallelism),
          imageGenerationParallelism: clampImageGenerationParallelism(imageGenerationParallelism)
        }
      });
      onSaved(next);
      onClose();
    } catch (error) {
      onError(`保存设置失败：${getErrorMessage(error)}`);
    } finally {
      setSaving(false);
    }
  }

  async function runRuntimeCheck(): Promise<void> {
    setCheckingRuntime(true);
    try {
      setRuntimeCheck(await api.settings.checkRuntime());
    } catch (error) {
      onError(`运行时自检失败：${getErrorMessage(error)}`);
    } finally {
      setCheckingRuntime(false);
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
                <SwitchRow checked={thinkingEnabled} onCheckedChange={setThinkingEnabled}>
                  启用模型思考模式
                </SwitchRow>
                <label>
                  推理强度
                  <select value={reasoningEffort} onChange={(event) => setReasoningEffort(event.target.value)}>
                    <option value="">默认</option>
                    <option value="none">none（关闭）</option>
                    <option value="minimal">minimal</option>
                    <option value="low">low</option>
                    <option value="medium">medium</option>
                    <option value="high">high</option>
                    <option value="xhigh">xhigh</option>
                  </select>
                </label>
                <p className="settings-note">
                  关闭后会写入 `OPENAI_THINKING_ENABLED=false`。Qwen/DashScope 兼容模型会使用 qwen 思考参数格式。
                </p>
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
                  允许 Agent 使用 image_generate 生成架构图和流程图
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
                  Agent 引擎
                  <input value="内置 @mariozechner/pi-coding-agent" readOnly />
                </label>
                <RuntimeStatusPanel settings={settings} check={runtimeCheck} />
                <p className="settings-note">
                  后端直接调用 Pi Agent SDK，并通过桥接层接入桌面端消息流、工具事件和文件产物。
                </p>
                <div className="runtime-check-actions">
                  <button
                    type="button"
                    className="secondary-action"
                    onClick={() => void runRuntimeCheck()}
                    disabled={checkingRuntime}
                  >
                    {checkingRuntime ? <Loader2 size={14} className="spin" /> : <Play size={14} />}
                    运行 Python 自检
                  </button>
                </div>
                <label>
                  请求超时 ms
                  <input
                    type="number"
                    value={timeout}
                    onChange={(event) => setTimeoutValue(Number(event.target.value))}
                  />
                </label>
                <label>
                  生图请求超时 ms
                  <input
                    type="number"
                    min={1000}
                    step={1000}
                    value={imageTimeout}
                    onChange={(event) => setImageTimeoutValue(Number(event.target.value))}
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
                <label>
                  方案章节并行数
                  <input
                    type="number"
                    min={DRAFT_SECTION_PARALLELISM_MIN}
                    max={DRAFT_SECTION_PARALLELISM_MAX}
                    step={1}
                    value={draftSectionParallelism}
                    onChange={(event) => setDraftSectionParallelism(clampDraftSectionParallelism(event.target.value))}
                  />
                </label>
                <p className="settings-note">
                  控制 draft_scheme_sections 默认同时起草的章节数；工具调用中的 max_parallel 仅作为单次覆盖。
                </p>
                <label>
                  生图并行数
                  <input
                    type="number"
                    min={IMAGE_GENERATION_PARALLELISM_MIN}
                    max={IMAGE_GENERATION_PARALLELISM_MAX}
                    step={1}
                    value={imageGenerationParallelism}
                    onChange={(event) =>
                      setImageGenerationParallelism(clampImageGenerationParallelism(event.target.value))
                    }
                  />
                </label>
                <p className="settings-note">
                  控制多个 image_generate 同时执行的数量；默认 10，可按生图接口配额下调。
                </p>
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
          <button className="send-action" onClick={() => void save()} disabled={saving}>
            {saving ? <Loader2 size={14} className="spin" /> : null}
            保存
          </button>
        </footer>
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}

function RuntimeStatusPanel({
  settings,
  check
}: {
  settings?: AppSettings;
  check?: RuntimeCheckResult;
}): JSX.Element {
  const bundledPython = settings?.runtime.bundledPython;
  const available = bundledPython?.available ?? false;

  return (
    <div className={`runtime-status ${available ? "available" : "missing"}`}>
      <div>
        <strong>{available ? "内置 Python 已启用" : "未发现内置 Python"}</strong>
        <span>
          {available
            ? `${runtimeSourceLabel(bundledPython?.source)} · exec_bash 会优先使用随包 Python`
            : "exec_bash 将回退到用户系统 PATH 中的 Python"}
        </span>
      </div>
      {available ? <code title={bundledPython?.pythonExePath}>{bundledPython?.pythonExePath}</code> : null}
      {check ? (
        <div className="runtime-check-result">
          <RuntimeCommandLine label="Python" check={check.python} />
          <RuntimeCommandLine label="pip" check={check.pip} />
          <small>检测时间：{formatDateTime(check.checkedAt)}</small>
        </div>
      ) : null}
    </div>
  );
}

function RuntimeCommandLine({ label, check }: { label: string; check: RuntimeCommandCheck }): JSX.Element {
  return (
    <div className={`runtime-command-line ${check.ok ? "ok" : "failed"}`}>
      <span>{label}</span>
      <strong>{check.ok ? "可用" : "失败"}</strong>
      <code title={check.output || check.error}>{check.output || check.error}</code>
    </div>
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

function findLatestSchemeProgressItem(session?: ChatSession): SchemeProgressItem | undefined {
  if (!session) return undefined;
  for (let index = session.items.length - 1; index >= 0; index -= 1) {
    const item = session.items[index];
    if (item.kind === "scheme_progress") return item;
  }
  return undefined;
}

function schemeProgressStatusLabel(status: SchemeProgressItem["status"]): string {
  const map = {
    pending: "待开始",
    running: "生成中",
    partial: "部分完成",
    completed: "已完成",
    failed: "失败"
  };
  return map[status];
}

function schemeSectionStatusLabel(status: SchemeSectionStatus): string {
  const map = {
    pending: "待生成",
    drafting: "起草中",
    drafted: "已起草",
    running: "生成中",
    completed: "已完成",
    failed: "失败",
    skipped: "跳过"
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

function runtimeSourceLabel(source?: AppSettings["runtime"]["bundledPython"]["source"]): string {
  if (source === "resources") return "安装包资源";
  if (source === "project") return "项目目录";
  return "运行时资源";
}

function formatToolPreview(content: string): string {
  const trimmed = content.trim();
  if (!trimmed) return "";
  try {
    return JSON.stringify(JSON.parse(trimmed), null, 2);
  } catch {
    return content;
  }
}
