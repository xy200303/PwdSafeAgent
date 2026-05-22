import { describe, expect, it } from "vitest";
import { orderStreamItemsForDisplay } from "../../src/renderer/src/streamOrdering";
import type { StreamItem } from "../../src/shared/types";

describe("streamOrdering", () => {
  it("keeps tool outputs above the assistant reply within a turn", () => {
    const items: StreamItem[] = [
      message("user_1", "user", "生成方案"),
      tool("tool_1", "time"),
      message("assistant_1", "assistant", "正在生成回复"),
      tool("tool_2", "write_word"),
      file("file_1")
    ];

    expect(orderStreamItemsForDisplay(items).map((item) => item.id)).toEqual([
      "user_1",
      "tool_1",
      "tool_2",
      "file_1",
      "assistant_1"
    ]);
  });

  it("does not move tool outputs across user turns", () => {
    const items: StreamItem[] = [
      message("user_1", "user", "你好"),
      tool("tool_1", "time"),
      message("assistant_1", "assistant", "你好"),
      message("user_2", "user", "生成方案"),
      message("assistant_2", "assistant", "方案正文"),
      tool("tool_2", "write_word")
    ];

    expect(orderStreamItemsForDisplay(items).map((item) => item.id)).toEqual([
      "user_1",
      "tool_1",
      "assistant_1",
      "user_2",
      "tool_2",
      "assistant_2"
    ]);
  });
});

function message(id: string, role: "user" | "assistant", content: string): StreamItem {
  return {
    id,
    kind: "message",
    role,
    content,
    isFinished: true,
    createdAt: "2026-05-23T00:00:00.000Z"
  };
}

function tool(id: string, toolName: string): StreamItem {
  return {
    id,
    kind: "tool",
    toolCallId: `${id}_call`,
    toolName,
    status: "success",
    summary: "完成",
    createdAt: "2026-05-23T00:00:00.000Z"
  };
}

function file(id: string): StreamItem {
  return {
    id,
    kind: "file",
    artifactId: `${id}_artifact`,
    name: "方案.docx",
    fileKind: "docx",
    createdAt: "2026-05-23T00:00:00.000Z"
  };
}
