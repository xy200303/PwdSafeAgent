import { describe, expect, it } from "vitest";
import type { SchemeProgressItem } from "../../src/shared/types";
import {
  rewriteContinuationPrompt,
  stripSyntheticSchemeCompletionNotice
} from "../../src/main/schemeContinuation";

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

  it("rewrites generic continue prompts into actionable continuation requests", () => {
    const progress: SchemeProgressItem = {
      id: "scheme_progress_1",
      kind: "scheme_progress",
      title: "方案章节生成",
      status: "partial",
      total: 158,
      drafted: 0,
      completed: 25,
      failed: 0,
      sections: [
        { id: "sec_3_4_1", number: "3.4.1", title: "上一节", headingLevel: 3, status: "completed" },
        { id: "sec_3_4_2", number: "3.4.2", title: "业务应用需求分析", headingLevel: 3, status: "pending" }
      ],
      createdAt: "2026-06-03T00:00:00.000Z"
    };

    const rewritten = rewriteContinuationPrompt("继续啊", progress);
    expect(rewritten).toContain("不要重复总结当前进度");
    expect(rewritten).toContain("当前章节进度：25/158。");
    expect(rewritten).toContain("当前优先处理：3.4.2 业务应用需求分析。");
  });

  it("keeps non-generic prompts untouched", () => {
    const progress: SchemeProgressItem = {
      id: "scheme_progress_1",
      kind: "scheme_progress",
      title: "方案章节生成",
      status: "partial",
      total: 10,
      drafted: 1,
      completed: 2,
      failed: 0,
      sections: [],
      createdAt: "2026-06-03T00:00:00.000Z"
    };

    expect(rewriteContinuationPrompt("继续把 3.4.2 改成更偏实施方案的写法", progress)).toBe(
      "继续把 3.4.2 改成更偏实施方案的写法"
    );
  });
});
