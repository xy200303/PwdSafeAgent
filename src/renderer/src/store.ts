import { configureStore, createSlice, type PayloadAction } from "@reduxjs/toolkit";
import type { AppSettings, ArtifactSummary, AttachmentRef, ChatSession, RendererEvent, StreamItem } from "../../shared/types";

interface ChatState {
  sessions: ChatSession[];
  artifacts: ArtifactSummary[];
  currentSessionId?: string;
  composer: string;
  pendingAttachments: AttachmentRef[];
  settings?: AppSettings;
  settingsOpen: boolean;
  artifactsOpen: boolean;
}

const initialState: ChatState = {
  sessions: [],
  artifacts: [],
  composer: "",
  pendingAttachments: [],
  settingsOpen: false,
  artifactsOpen: false
};

const chatSlice = createSlice({
  name: "chat",
  initialState,
  reducers: {
    setSessions(state, action: PayloadAction<ChatSession[]>) {
      state.sessions = action.payload;
      if (!state.currentSessionId || !action.payload.some((session) => session.id === state.currentSessionId)) {
        state.currentSessionId = action.payload[0]?.id;
      }
    },
    setCurrentSession(state, action: PayloadAction<string>) {
      state.currentSessionId = action.payload;
    },
    upsertSession(state, action: PayloadAction<ChatSession>) {
      const index = state.sessions.findIndex((session) => session.id === action.payload.id);
      if (index >= 0) {
        state.sessions[index] = action.payload;
      } else {
        state.sessions.unshift(action.payload);
      }
      state.currentSessionId ??= action.payload.id;
    },
    addStreamItem(state, action: PayloadAction<{ sessionId: string; item: StreamItem }>) {
      const session = state.sessions.find((existing) => existing.id === action.payload.sessionId);
      if (!session) return;
      if (!session.items.some((item) => item.id === action.payload.item.id)) {
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
    },
    addAttachments(state, action: PayloadAction<AttachmentRef[]>) {
      state.pendingAttachments.push(...action.payload);
    },
    removeAttachment(state, action: PayloadAction<string>) {
      state.pendingAttachments = state.pendingAttachments.filter((attachment) => attachment.id !== action.payload);
    },
    removeSession(state, action: PayloadAction<string>) {
      const deletedArtifactIds = getSessionArtifactIds(
        state.sessions.find((session) => session.id === action.payload)
      );
      state.sessions = state.sessions.filter((session) => session.id !== action.payload);
      state.artifacts = state.artifacts.filter(
        (artifact) => artifact.sessionId !== action.payload && !deletedArtifactIds.has(artifact.id)
      );
      if (state.currentSessionId === action.payload) {
        state.currentSessionId = state.sessions[0]?.id;
      }
    },
    setArtifacts(state, action: PayloadAction<ArtifactSummary[]>) {
      state.artifacts = action.payload;
    },
    upsertArtifact(state, action: PayloadAction<ArtifactSummary>) {
      const index = state.artifacts.findIndex((artifact) => artifact.id === action.payload.id);
      if (index >= 0) {
        state.artifacts[index] = action.payload;
      } else {
        state.artifacts.unshift(action.payload);
      }
    },
    clearComposer(state) {
      state.composer = "";
      state.pendingAttachments = [];
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
        if (index >= 0) {
          state.sessions[index] = event.payload;
        } else {
          state.sessions.unshift(event.payload);
        }
        state.currentSessionId ??= event.payload.id;
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
        if (state.currentSessionId === event.sessionId) {
          state.currentSessionId = state.sessions[0]?.id;
        }
        return;
      }
      if (event.type === "stream.item.added") {
        const session = state.sessions.find((existing) => existing.id === event.sessionId);
        if (session && !session.items.some((item) => item.id === event.payload.id)) {
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
  setSessions,
  setSettings,
  setSettingsOpen,
  upsertArtifact,
  upsertSession
} = chatSlice.actions;

export const store = configureStore({
  reducer: {
    chat: chatSlice.reducer
  }
});

export type RootState = ReturnType<typeof store.getState>;
export type AppDispatch = typeof store.dispatch;

function getSessionArtifactIds(session?: ChatSession): Set<string> {
  return new Set(
    session?.items
      .filter((item) => item.kind === "file")
      .map((item) => item.artifactId) ?? []
  );
}
