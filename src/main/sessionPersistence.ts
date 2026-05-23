import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import type { ArtifactSummary, AttachmentRef, ChatSession } from "../shared/types";

export interface SessionMemoryEntry {
  source: string;
  content: string;
}

export interface PersistedStateSnapshot {
  sessions: ChatSession[];
  attachments: AttachmentRef[];
  artifacts: ArtifactSummary[];
  sessionMemories: Record<string, SessionMemoryEntry[]>;
}

interface PersistedStateFile extends PersistedStateSnapshot {
  version: 1;
  savedAt: string;
}

export function loadPersistedState(filePath: string): PersistedStateSnapshot | null {
  if (!existsSync(filePath)) return null;

  const parsed = JSON.parse(readFileSync(filePath, "utf-8")) as Partial<PersistedStateFile>;
  return {
    sessions: Array.isArray(parsed.sessions) ? parsed.sessions : [],
    attachments: Array.isArray(parsed.attachments) ? parsed.attachments : [],
    artifacts: Array.isArray(parsed.artifacts) ? parsed.artifacts : [],
    sessionMemories: isMemoryRecord(parsed.sessionMemories) ? parsed.sessionMemories : {}
  };
}

export function savePersistedState(filePath: string, snapshot: PersistedStateSnapshot): void {
  mkdirSync(dirname(filePath), { recursive: true });
  const payload: PersistedStateFile = {
    version: 1,
    savedAt: new Date().toISOString(),
    sessions: snapshot.sessions,
    attachments: snapshot.attachments,
    artifacts: snapshot.artifacts,
    sessionMemories: snapshot.sessionMemories
  };
  writeFileSync(filePath, `${JSON.stringify(payload, null, 2)}\n`, "utf-8");
}

function isMemoryRecord(value: unknown): value is Record<string, SessionMemoryEntry[]> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  return Object.values(value).every(
    (entries) =>
      Array.isArray(entries) &&
      entries.every(
        (entry) =>
          entry &&
          typeof entry === "object" &&
          typeof (entry as SessionMemoryEntry).source === "string" &&
          typeof (entry as SessionMemoryEntry).content === "string"
      )
  );
}
