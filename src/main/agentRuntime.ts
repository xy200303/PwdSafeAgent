import type { ChatCompletionMessageParam } from "openai/resources/chat/completions";
import type { AppSettings, ChatSession, StreamItem } from "../shared/types";
import type { BundledPythonRuntime } from "./bundledRuntime";
import { createPiAgentBridge } from "./piAgentBridge";

export type MessageStreamItem = Extract<StreamItem, { kind: "message" }>;

export interface AgentRuntimeTurnInput {
  sessionId: string;
  userPrompt: string;
  controller: AbortController;
  onAssistantCreated: (assistantItem: MessageStreamItem) => void;
}

export interface AgentRuntime {
  runTurn(input: AgentRuntimeTurnInput): Promise<MessageStreamItem>;
}

export interface AgentRuntimeHost {
  rootDir: string;
  docsDir: string;
  inputDir: string;
  outputDir: string;
  loadSettings(): AppSettings;
  getSession(sessionId: string): ChatSession | undefined;
  buildMessages(session: ChatSession): ChatCompletionMessageParam[];
  createAssistantMessage(sessionId: string): MessageStreamItem;
  updateItem(sessionId: string, item: StreamItem): void;
  startToolCall(sessionId: string, toolName: string, summary: string): StreamItem;
  finishToolCall(sessionId: string, item: StreamItem, status: "success" | "failed", summary: string): void;
  addStage(sessionId: string, title: string, detail?: string): void;
  appendSessionMemory(sessionId: string, source: string, content: string): void;
  formatSessionMemory(sessionId: string): string;
  getSessionReadableFiles(sessionId: string): string[];
  getBundledPythonRuntime(): BundledPythonRuntime | undefined;
  createArtifact(sessionId: string, filePath: string): void;
}

export function createAgentRuntime(host: AgentRuntimeHost): AgentRuntime {
  return createPiAgentBridge(host);
}
