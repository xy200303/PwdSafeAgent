import { configureStore, createSlice, type PayloadAction } from "@reduxjs/toolkit";
import type {
  AppSettings,
  ArtifactSummary,
  AttachmentRef,
  ChatSession,
  DocumentTemplateSummary,
  RendererEvent,
  StreamItem
} from "../../shared/types";

interface SessionDraft {
  composer: string;
  pendingAttachments: AttachmentRef[];
}

interface ChatState {
  sessions: ChatSession[];
  artifacts: ArtifactSummary[];
  documentTemplates: DocumentTemplateSummary[];
  selectedDocumentTemplateId: string;
  currentSessionId?: string;
  composer: string;
  pendingAttachments: AttachmentRef[];
  drafts: Record<string, SessionDraft>;
  settings?: AppSettings;
  settingsOpen: boolean;
  artifactsOpen: boolean;
}

const initialState: ChatState = {
  sessions: [],
  artifacts: [],
  documentTemplates: [],
  selectedDocumentTemplateId: "default",
  composer: "",
  pendingAttachments: [],
  drafts: {},
  settingsOpen: false,
  artifactsOpen: false
};

const chatSlice = createSlice({
  name: "chat",
  initialState,
  reducers: {
    setSessions(state, action: PayloadAction<ChatSession[]>) {
      state.sessions = action.payload.map((session) =>
        mergeSessionPreservingFileItems(
          state.sessions.find((existing) => existing.id === session.id),
          session
        )
      );
      pruneDrafts(state);
      if (!state.currentSessionId || !action.payload.some((session) => session.id === state.currentSessionId)) {
        activateSession(state, action.payload[0]?.id);
      }
    },
    setCurrentSession(state, action: PayloadAction<string>) {
      activateSession(state, action.payload);
    },
    upsertSession(state, action: PayloadAction<ChatSession>) {
      const index = state.sessions.findIndex((session) => session.id === action.payload.id);
      const nextSession = mergeSessionPreservingFileItems(state.sessions[index], action.payload);
      if (index >= 0) {
        state.sessions[index] = nextSession;
      } else {
        state.sessions.unshift(nextSession);
      }
      if (!state.currentSessionId) {
        activateSession(state, action.payload.id);
      }
    },
    addStreamItem(state, action: PayloadAction<{ sessionId: string; item: StreamItem }>) {
      const session = state.sessions.find((existing) => existing.id === action.payload.sessionId);
      if (!session) return;
      if (!hasMatchingStreamItem(session.items, action.payload.item)) {
        session.items.push(action.payload.item);
      }
    },
    updateStreamItem(state, action: PayloadAction<{ sessionId: string; item: StreamItem }>) {
      const session = state.sessions.find((existing) => existing.id === action.payload.sessionId);
      if (!session) return;
      const index = session.items.findIndex((item) => item.id === action.payload.item.id);
      if (index >= 0) {
        session.items[index] = action.payload.item;
      }
    },
    setComposer(state, action: PayloadAction<string>) {
      state.composer = action.payload;
      syncActiveDraft(state);
    },
    addAttachments(state, action: PayloadAction<{ attachments: AttachmentRef[]; sessionId?: string }>) {
      const targetSessionId =
        action.payload.sessionId ?? action.payload.attachments[0]?.sessionId ?? state.currentSessionId;
      if (!targetSessionId || action.payload.attachments.length === 0) return;
      const draft = readDraft(state, targetSessionId);
      writeDraft(state, targetSessionId, {
        composer: draft.composer,
        pendingAttachments: [...draft.pendingAttachments, ...action.payload.attachments]
      });
    },
    removeAttachment(state, action: PayloadAction<string>) {
      state.pendingAttachments = state.pendingAttachments.filter((attachment) => attachment.id !== action.payload);
      syncActiveDraft(state);
    },
    removeSession(state, action: PayloadAction<string>) {
      const deletedArtifactIds = getSessionArtifactIds(
        state.sessions.find((session) => session.id === action.payload)
      );
      state.sessions = state.sessions.filter((session) => session.id !== action.payload);
      state.artifacts = state.artifacts.filter(
        (artifact) => artifact.sessionId !== action.payload && !deletedArtifactIds.has(artifact.id)
      );
      delete state.drafts[action.payload];
      if (state.currentSessionId === action.payload) {
        activateSession(state, state.sessions[0]?.id);
      }
    },
    setArtifacts(state, action: PayloadAction<ArtifactSummary[]>) {
      state.artifacts = action.payload;
    },
    setDocumentTemplates(state, action: PayloadAction<DocumentTemplateSummary[]>) {
      state.documentTemplates = action.payload;
      if (!state.documentTemplates.some((template) => template.id === state.selectedDocumentTemplateId)) {
        state.selectedDocumentTemplateId = "default";
      }
    },
    upsertDocumentTemplate(state, action: PayloadAction<DocumentTemplateSummary>) {
      const index = state.documentTemplates.findIndex((template) => template.id === action.payload.id);
      if (index >= 0) {
        state.documentTemplates[index] = action.payload;
      } else {
        state.documentTemplates.push(action.payload);
      }
      state.selectedDocumentTemplateId = action.payload.id;
    },
    setSelectedDocumentTemplate(state, action: PayloadAction<string>) {
      state.selectedDocumentTemplateId = state.documentTemplates.some((template) => template.id === action.payload)
        ? action.payload
        : "default";
    },
    upsertArtifact(state, action: PayloadAction<ArtifactSummary>) {
      const index = state.artifacts.findIndex((artifact) => artifact.id === action.payload.id);
      if (index >= 0) {
        state.artifacts[index] = action.payload;
      } else {
        state.artifacts.unshift(action.payload);
      }
    },
    clearComposer(state, action: PayloadAction<string>) {
      writeDraft(state, action.payload, {
        composer: "",
        pendingAttachments: []
      });
    },
    setSettings(state, action: PayloadAction<AppSettings>) {
      state.settings = action.payload;
    },
    setSettingsOpen(state, action: PayloadAction<boolean>) {
      state.settingsOpen = action.payload;
    },
    setArtifactsOpen(state, action: PayloadAction<boolean>) {
      state.artifactsOpen = action.payload;
    },
    applyRendererEvent(state, action: PayloadAction<RendererEvent>) {
      const event = action.payload;
      if (event.type === "session.created" || event.type === "session.updated") {
        const index = state.sessions.findIndex((session) => session.id === event.payload.id);
        const nextSession = mergeSessionPreservingFileItems(state.sessions[index], event.payload);
        if (index >= 0) {
          state.sessions[index] = nextSession;
        } else {
          state.sessions.unshift(nextSession);
        }
        if (!state.currentSessionId) {
          activateSession(state, event.payload.id);
        }
        return;
      }
      if (event.type === "session.deleted") {
        const deletedArtifactIds = getSessionArtifactIds(
          state.sessions.find((session) => session.id === event.sessionId)
        );
        state.sessions = state.sessions.filter((session) => session.id !== event.sessionId);
        state.artifacts = state.artifacts.filter(
          (artifact) => artifact.sessionId !== event.sessionId && !deletedArtifactIds.has(artifact.id)
        );
        delete state.drafts[event.sessionId];
        if (state.currentSessionId === event.sessionId) {
          activateSession(state, state.sessions[0]?.id);
        }
        return;
      }
      if (event.type === "stream.item.added") {
        const session = state.sessions.find((existing) => existing.id === event.sessionId);
        if (session && !hasMatchingStreamItem(session.items, event.payload)) {
          session.items.push(event.payload);
        }
        return;
      }
      if (event.type === "stream.item.updated") {
        const session = state.sessions.find((existing) => existing.id === event.sessionId);
        const index = session?.items.findIndex((item) => item.id === event.payload.id) ?? -1;
        if (session && index >= 0) {
          session.items[index] = event.payload;
        }
        return;
      }
      if (event.type === "artifact.created") {
        const index = state.artifacts.findIndex((artifact) => artifact.id === event.payload.id);
        const nextArtifact = { ...event.payload, sessionId: event.payload.sessionId ?? event.sessionId };
        if (index >= 0) {
          state.artifacts[index] = nextArtifact;
        } else {
          state.artifacts.unshift(nextArtifact);
        }
        const session = state.sessions.find((existing) => existing.id === event.sessionId);
        if (session) {
          const fileItem: StreamItem = {
            id: `file_${nextArtifact.id}`,
            kind: "file",
            artifactId: nextArtifact.id,
            name: nextArtifact.name,
            fileKind: nextArtifact.kind,
            createdAt: nextArtifact.createdAt
          };
          if (!hasMatchingStreamItem(session.items, fileItem)) {
            session.items.push(fileItem);
          }
        }
      }
    }
  }
});

export const {
  addAttachments,
  applyRendererEvent,
  clearComposer,
  removeAttachment,
  removeSession,
  setArtifacts,
  setArtifactsOpen,
  setComposer,
  setCurrentSession,
  setDocumentTemplates,
  setSessions,
  setSettings,
  setSettingsOpen,
  setSelectedDocumentTemplate,
  upsertArtifact,
  upsertDocumentTemplate,
  upsertSession
} = chatSlice.actions;

export const store = configureStore({
  reducer: {
    chat: chatSlice.reducer
  }
});

export type RootState = ReturnType<typeof store.getState>;
export type AppDispatch = typeof store.dispatch;

const EMPTY_DRAFT: SessionDraft = {
  composer: "",
  pendingAttachments: []
};

function getSessionArtifactIds(session?: ChatSession): Set<string> {
  return new Set(
    session?.items
      .filter((item) => item.kind === "file")
      .map((item) => item.artifactId) ?? []
  );
}

function hasMatchingStreamItem(items: StreamItem[], item: StreamItem): boolean {
  if (items.some((existing) => existing.id === item.id)) return true;
  if (item.kind === "file") {
    return items.some((existing) => existing.kind === "file" && existing.artifactId === item.artifactId);
  }
  return false;
}

function mergeSessionPreservingFileItems(existing: ChatSession | undefined, incoming: ChatSession): ChatSession {
  if (!existing) return incoming;

  const preservedFileItems = existing.items.filter(
    (item) => item.kind === "file" && !incoming.items.some((incomingItem) => incomingItem.kind === "file" && incomingItem.artifactId === item.artifactId)
  );
  if (!preservedFileItems.length) return incoming;

  return {
    ...incoming,
    items: [...incoming.items, ...preservedFileItems].sort((left, right) => left.createdAt.localeCompare(right.createdAt))
  };
}

function activateSession(state: ChatState, sessionId?: string): void {
  state.currentSessionId = sessionId;
  const draft = sessionId ? readDraft(state, sessionId) : EMPTY_DRAFT;
  state.composer = draft.composer;
  state.pendingAttachments = draft.pendingAttachments;
}

function syncActiveDraft(state: ChatState): void {
  if (!state.currentSessionId) return;
  state.drafts[state.currentSessionId] = {
    composer: state.composer,
    pendingAttachments: [...state.pendingAttachments]
  };
}

function readDraft(state: ChatState, sessionId: string): SessionDraft {
  const draft = state.drafts[sessionId];
  return {
    composer: draft?.composer || "",
    pendingAttachments: [...(draft?.pendingAttachments ?? [])]
  };
}

function writeDraft(state: ChatState, sessionId: string, draft: SessionDraft): void {
  state.drafts[sessionId] = {
    composer: draft.composer,
    pendingAttachments: [...draft.pendingAttachments]
  };
  if (state.currentSessionId === sessionId) {
    state.composer = draft.composer;
    state.pendingAttachments = [...draft.pendingAttachments];
  }
}

function pruneDrafts(state: ChatState): void {
  const validSessionIds = new Set(state.sessions.map((session) => session.id));
  for (const sessionId of Object.keys(state.drafts)) {
    if (!validSessionIds.has(sessionId)) {
      delete state.drafts[sessionId];
    }
  }
}
