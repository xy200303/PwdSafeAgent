import { beforeEach, describe, expect, it } from "vitest";
import { applyRendererEvent, setArtifacts, setCurrentSession, setSessions, store } from "../../src/renderer/src/store";
import type { ChatSession, RendererEvent } from "../../src/shared/types";

describe("renderer store", () => {
  beforeEach(() => {
    store.dispatch(setSessions([]));
    store.dispatch(setArtifacts([]));
  });

  it("selects the first available session when the current session disappears", () => {
    const first = makeSession("session_a", "A");
    const second = makeSession("session_b", "B");

    store.dispatch(setSessions([first, second]));
    store.dispatch(setCurrentSession(second.id));
    store.dispatch(setSessions([first]));

    expect(store.getState().chat.currentSessionId).toBe(first.id);
  });

  it("applies session.deleted events and keeps the sidebar selection valid", () => {
    const first = makeSession("session_a", "A");
    const second = {
      ...makeSession("session_b", "B"),
      items: [
        {
          id: "file_1",
          kind: "file",
          artifactId: "artifact_1",
          name: "方案.docx",
          fileKind: "docx",
          createdAt: "2026-05-23T00:01:00.000Z"
        } as const
      ]
    };
    const event: RendererEvent = {
      id: "event_1",
      type: "session.deleted",
      sessionId: second.id
    };

    store.dispatch(setSessions([first, second]));
    store.dispatch(
      setArtifacts([
        {
          id: "artifact_1",
          name: "方案.docx",
          kind: "docx",
          path: "C:/tmp/方案.docx",
          size: 2048,
          createdAt: "2026-05-23T00:01:00.000Z"
        }
      ])
    );
    store.dispatch(setCurrentSession(second.id));
    store.dispatch(applyRendererEvent(event));

    expect(store.getState().chat.sessions.map((session) => session.id)).toEqual([first.id]);
    expect(store.getState().chat.artifacts).toEqual([]);
    expect(store.getState().chat.currentSessionId).toBe(first.id);
  });

  it("records created artifacts with their owning session", () => {
    const session = makeSession("session_a", "A");
    const event: RendererEvent = {
      id: "event_2",
      type: "artifact.created",
      sessionId: "session_a",
      payload: {
        id: "artifact_1",
        name: "方案.docx",
        kind: "docx",
        path: "C:/tmp/方案.docx",
        size: 2048,
        createdAt: "2026-05-23T00:01:00.000Z"
      }
    };

    store.dispatch(setSessions([session]));
    store.dispatch(applyRendererEvent(event));

    expect(store.getState().chat.artifacts).toEqual([{ ...event.payload, sessionId: "session_a" }]);
    expect(store.getState().chat.sessions[0]?.items).toEqual([
      {
        id: "file_artifact_1",
        kind: "file",
        artifactId: "artifact_1",
        name: "方案.docx",
        fileKind: "docx",
        createdAt: "2026-05-23T00:01:00.000Z"
      }
    ]);
  });

  it("deduplicates file stream items when artifact and stream events arrive out of order", () => {
    const session = makeSession("session_a", "A");
    const artifactEvent: RendererEvent = {
      id: "event_2",
      type: "artifact.created",
      sessionId: session.id,
      payload: {
        id: "artifact_1",
        name: "方案.docx",
        kind: "docx",
        path: "C:/tmp/方案.docx",
        size: 2048,
        createdAt: "2026-05-23T00:01:00.000Z"
      }
    };
    const streamEvent: RendererEvent = {
      id: "event_3",
      type: "stream.item.added",
      sessionId: session.id,
      payload: {
        id: "file_1",
        kind: "file",
        artifactId: "artifact_1",
        name: "方案.docx",
        fileKind: "docx",
        createdAt: "2026-05-23T00:01:01.000Z"
      }
    };

    store.dispatch(setSessions([session]));
    store.dispatch(applyRendererEvent(artifactEvent));
    store.dispatch(applyRendererEvent(streamEvent));

    expect(store.getState().chat.sessions[0]?.items.filter((item) => item.kind === "file")).toHaveLength(1);
  });

  it("keeps file cards when a later session update is missing them", () => {
    const session = makeSession("session_a", "A");
    const artifactEvent: RendererEvent = {
      id: "event_2",
      type: "artifact.created",
      sessionId: session.id,
      payload: {
        id: "artifact_1",
        name: "方案.docx",
        kind: "docx",
        path: "C:/tmp/方案.docx",
        size: 2048,
        createdAt: "2026-05-23T00:01:00.000Z"
      }
    };
    const staleSessionUpdate: RendererEvent = {
      id: "event_4",
      type: "session.updated",
      payload: {
        ...session,
        items: [
          {
            id: "tool_1",
            kind: "tool",
            toolCallId: "call_1",
            toolName: "send_file",
            status: "success",
            summary: "已发送 方案.docx",
            createdAt: "2026-05-23T00:00:59.000Z"
          }
        ]
      }
    };

    store.dispatch(setSessions([session]));
    store.dispatch(applyRendererEvent(artifactEvent));
    store.dispatch(applyRendererEvent(staleSessionUpdate));

    expect(store.getState().chat.sessions[0]?.items).toEqual([
      staleSessionUpdate.payload.items[0],
      {
        id: "file_artifact_1",
        kind: "file",
        artifactId: "artifact_1",
        name: "方案.docx",
        fileKind: "docx",
        createdAt: "2026-05-23T00:01:00.000Z"
      }
    ]);
  });

  it("moves a refreshed file card to the latest session position", () => {
    const session = {
      ...makeSession("session_a", "A"),
      items: [
        {
          id: "file_1",
          kind: "file",
          artifactId: "artifact_1",
          name: "旧方案.docx",
          fileKind: "docx",
          createdAt: "2026-05-23T00:01:00.000Z"
        },
        {
          id: "tool_1",
          kind: "tool",
          toolCallId: "call_1",
          toolName: "write_word",
          status: "success",
          summary: "已生成旧方案.docx",
          createdAt: "2026-05-23T00:01:01.000Z"
        }
      ] satisfies ChatSession["items"]
    };
    const refreshedFile = {
      id: "file_1",
      kind: "file",
      artifactId: "artifact_1",
      name: "新方案.docx",
      fileKind: "docx",
      createdAt: "2026-05-23T00:02:01.000Z"
    } satisfies ChatSession["items"][number];
    const sendTool = {
      id: "tool_2",
      kind: "tool",
      toolCallId: "call_2",
      toolName: "send_file",
      status: "success",
      summary: "已发送 新方案.docx",
      createdAt: "2026-05-23T00:02:00.000Z"
    } satisfies ChatSession["items"][number];

    store.dispatch(setSessions([session]));
    store.dispatch(
      applyRendererEvent({
        id: "event_5",
        type: "session.updated",
        payload: {
          ...session,
          items: [session.items[1], sendTool, refreshedFile]
        }
      })
    );

    expect(store.getState().chat.sessions[0]?.items.map((item) => item.id)).toEqual(["tool_1", "tool_2", "file_1"]);
    expect(store.getState().chat.sessions[0]?.items.at(-1)).toMatchObject({
      kind: "file",
      name: "新方案.docx"
    });
  });

  it("preserves tool detail previews from stream events", () => {
    const session = makeSession("session_a", "A");
    const event: RendererEvent = {
      id: "event_3",
      type: "stream.item.added",
      sessionId: session.id,
      payload: {
        id: "tool_1",
        kind: "tool",
        toolCallId: "call_1",
        toolName: "read_word",
        status: "success",
        summary: "已读取模板",
        inputPreview: "{\"path\":\"docs/templates/密码应用方案.docx\"}",
        outputPreview: "模板正文摘要",
        createdAt: "2026-05-23T00:01:00.000Z"
      }
    };

    store.dispatch(setSessions([session]));
    store.dispatch(applyRendererEvent(event));

    expect(store.getState().chat.sessions[0]?.items[0]).toMatchObject({
      kind: "tool",
      inputPreview: "{\"path\":\"docs/templates/密码应用方案.docx\"}",
      outputPreview: "模板正文摘要"
    });
  });
});

function makeSession(id: string, title: string): ChatSession {
  return {
    id,
    title,
    status: "idle",
    createdAt: "2026-05-23T00:00:00.000Z",
    updatedAt: "2026-05-23T00:00:00.000Z",
    items: []
  };
}
