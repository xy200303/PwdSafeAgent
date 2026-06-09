import { contextBridge, ipcRenderer } from "electron";
import type {
  ArtifactListInput,
  ChatPromptInput,
  ClipboardAttachmentInput,
  PickAttachmentInput,
  PwdSafeAgentApi,
  RenameSessionInput,
  RendererEvent,
  UpdateAppSettingsInput
} from "../shared/types";

const api: PwdSafeAgentApi = {
  app: {
    getMetadata: () => ipcRenderer.invoke("app:metadata")
  },
  session: {
    list: () => ipcRenderer.invoke("session:list"),
    create: () => ipcRenderer.invoke("session:create"),
    rename: (input: RenameSessionInput) => ipcRenderer.invoke("session:rename", input),
    delete: (sessionId: string) => ipcRenderer.invoke("session:delete", sessionId)
  },
  chat: {
    prompt: (input: ChatPromptInput) => ipcRenderer.invoke("chat:prompt", input),
    abort: (sessionId: string) => ipcRenderer.invoke("chat:abort", sessionId)
  },
  attachment: {
    pick: (input: PickAttachmentInput) => ipcRenderer.invoke("attachment:pick", input),
    importClipboard: (input: ClipboardAttachmentInput) => ipcRenderer.invoke("attachment:import-clipboard", input),
    remove: (attachmentId: string) => ipcRenderer.invoke("attachment:remove", attachmentId)
  },
  artifact: {
    list: (input?: ArtifactListInput) => ipcRenderer.invoke("artifact:list", input),
    open: (artifactId: string) => ipcRenderer.invoke("artifact:open", artifactId),
    reveal: (artifactId: string) => ipcRenderer.invoke("artifact:reveal", artifactId),
    preview: (artifactId: string) => ipcRenderer.invoke("artifact:preview", artifactId)
  },
  documentTemplate: {
    list: () => ipcRenderer.invoke("document-template:list"),
    getSelected: () => ipcRenderer.invoke("document-template:get-selected"),
    select: (templateId: string) => ipcRenderer.invoke("document-template:select", templateId),
    upload: () => ipcRenderer.invoke("document-template:upload")
  },
  settings: {
    get: () => ipcRenderer.invoke("settings:get"),
    save: (input: UpdateAppSettingsInput) => ipcRenderer.invoke("settings:save", input),
    checkRuntime: () => ipcRenderer.invoke("settings:check-runtime")
  },
  events: {
    subscribe: (listener: (event: RendererEvent) => void) => {
      const handler = (_event: Electron.IpcRendererEvent, payload: RendererEvent) => listener(payload);
      ipcRenderer.on("app:event", handler);
      return () => ipcRenderer.off("app:event", handler);
    }
  }
};

contextBridge.exposeInMainWorld("pwdSafeAgent", api);
