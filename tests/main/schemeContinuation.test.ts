import { describe, expect, it } from "vitest";
import { stripSyntheticSchemeCompletionNotice } from "../../src/main/schemeContinuation";

describe("schemeContinuation", () => {
  it("strips synthetic incomplete-progress notice from assistant history", () => {
    const content = [
      "> 章节生成未完成：当前 25/158。",
      "> 当前 Word 只能视为阶段性文件，不是完整方案。",
      "> 下一步应继续处理 3.4.2 业务应用需求分析。",
      "",
      "我已经完成本批章节写入，接下来会继续推进。"
    ].join("\n");

    expect(stripSyntheticSchemeCompletionNotice(content)).toBe("我已经完成本批章节写入，接下来会继续推进。");
  });
});
