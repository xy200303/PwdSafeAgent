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

    store.dispatch(applyRendererEvent(event));

    expect(store.getState().chat.artifacts).toEqual([{ ...event.payload, sessionId: "session_a" }]);
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
