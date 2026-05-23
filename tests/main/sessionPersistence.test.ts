import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { loadPersistedState, savePersistedState, type PersistedStateSnapshot } from "../../src/main/sessionPersistence";

describe("sessionPersistence", () => {
  it("saves and restores sessions, artifacts, attachments and memory", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-state-"));
    const filePath = join(dir, "state.json");
    const snapshot: PersistedStateSnapshot = {
      sessions: [
        {
          id: "session_1",
          title: "统一身份认证系统",
          status: "completed",
          createdAt: "2026-05-23T00:00:00.000Z",
          updatedAt: "2026-05-23T00:01:00.000Z",
          items: [
            {
              id: "msg_1",
              kind: "message",
              role: "assistant",
              content: "已生成方案。",
              isFinished: true,
              createdAt: "2026-05-23T00:01:00.000Z"
            }
          ]
        }
      ],
      attachments: [
        {
          id: "attachment_1",
          sessionId: "session_1",
          name: "需求说明.docx",
          mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          size: 1024,
          path: "C:/tmp/需求说明.docx",
          source: "picker",
          createdAt: "2026-05-23T00:00:10.000Z"
        }
      ],
      artifacts: [
        {
          id: "artifact_1",
          name: "方案.docx",
          kind: "docx",
          path: "C:/tmp/方案.docx",
          size: 2048,
          createdAt: "2026-05-23T00:01:00.000Z"
        }
      ],
      sessionMemories: {
        session_1: [{ source: "Word 模板", content: "模板内容" }]
      }
    };

    try {
      savePersistedState(filePath, snapshot);
      const restored = loadPersistedState(filePath);

      expect(restored?.sessions[0]?.title).toBe("统一身份认证系统");
      expect(restored?.attachments[0]?.name).toBe("需求说明.docx");
      expect(restored?.artifacts[0]?.kind).toBe("docx");
      expect(restored?.sessionMemories.session_1[0]?.content).toBe("模板内容");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("returns null when no state file exists", () => {
    expect(loadPersistedState(join(tmpdir(), "missing-pwd-safe-agent-state.json"))).toBeNull();
  });

  it("ignores legacy template preload state while restoring", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-legacy-state-"));
    const filePath = join(dir, "state.json");

    try {
      await writeFile(
        filePath,
        JSON.stringify(
          {
            version: 1,
            savedAt: "2026-05-23T00:00:00.000Z",
            sessions: [],
            attachments: [],
            artifacts: [],
            templateLoadedSessionIds: ["legacy_session"],
            sessionMemories: {}
          },
          null,
          2
        ),
        "utf-8"
      );

      expect(loadPersistedState(filePath)).toEqual({
        sessions: [],
        attachments: [],
        artifacts: [],
        sessionMemories: {}
      });
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
