import { copyFileSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import {
  applySchemeProgressUpdate,
  createSchemeProgressItem,
  extractTemplateAnchorIds
} from "../../src/main/schemeProgress";

const BUILT_IN_TEMPLATE_DOCX_PATH = join(process.cwd(), "resources", "docs", "templates", "密码应用方案.docx");

describe("schemeProgress", () => {
  it("creates progress from template JSON and updates sections by number or anchor", () => {
    const docsDir = mkdtempSync(join(tmpdir(), "scheme-progress-"));
    const templatesDir = join(docsDir, "templates");
    try {
      mkdirSync(templatesDir, { recursive: true });
      writeFileSync(
        join(templatesDir, "密码应用方案.template.json"),
        JSON.stringify({
          sections: [
            { id: "sec_2", number: "2", title: "系统概述", headingLevel: 1, writingHint: "概述系统" },
            { id: "sec_2_2_2", number: "2.2.2", title: "网络环境", headingLevel: 3, relatedTables: ["table_1"] }
          ]
        }),
        "utf-8"
      );

      const item = createSchemeProgressItem({
        id: "scheme_progress_1",
        docsDir,
        createdAt: "2026-05-25T00:00:00.000Z"
      });

      expect(item.total).toBe(2);
      expect(item.sections[1]).toMatchObject({
        id: "sec_2_2_2",
        number: "2.2.2",
        status: "pending"
      });

      const running = applySchemeProgressUpdate(
        item,
        { section: "2.2.2 网络环境", status: "drafting", detail: "正在起草 2.2.2" },
        "2026-05-25T00:01:00.000Z"
      );
      expect(running.activeSectionId).toBe("sec_2_2_2");
      expect(running.sections[1]?.status).toBe("drafting");

      const drafted = applySchemeProgressUpdate(
        running,
        { anchorIds: ["sec_2_2_2"], status: "drafted", detail: "已起草" },
        "2026-05-25T00:02:00.000Z"
      );
      expect(drafted.drafted).toBe(1);
      expect(drafted.completed).toBe(0);
      expect(drafted.activeSectionId).toBeUndefined();

      const completed = applySchemeProgressUpdate(
        drafted,
        { anchorIds: ["sec_2_2_2"], status: "completed", detail: "已写入模板" },
        "2026-05-25T00:03:00.000Z"
      );
      expect(completed.drafted).toBe(0);
      expect(completed.completed).toBe(1);
      expect(completed.sections[1]).toMatchObject({
        status: "completed",
        detail: "已写入模板"
      });
    } finally {
      rmSync(docsDir, { recursive: true, force: true });
    }
  });

  it("creates progress from docx anchors when template JSON is missing", () => {
    const docsDir = mkdtempSync(join(tmpdir(), "scheme-progress-docx-"));
    const templatesDir = join(docsDir, "templates");
    try {
      mkdirSync(templatesDir, { recursive: true });
      copyFileSync(BUILT_IN_TEMPLATE_DOCX_PATH, join(templatesDir, "密码应用方案.docx"));

      const item = createSchemeProgressItem({
        id: "scheme_progress_docx",
        docsDir,
        createdAt: "2026-05-25T00:00:00.000Z"
      });

      expect(item.total).toBeGreaterThan(0);
      expect(item.sections.find((section) => section.id === "sec_5_4_9_4")).toMatchObject({
        number: "5.4.9.4",
        status: "pending"
      });
    } finally {
      rmSync(docsDir, { recursive: true, force: true });
    }
  });

  it("creates progress from custom document sections", () => {
    const item = createSchemeProgressItem({
      id: "scheme_progress_sections",
      docsDir: process.cwd(),
      title: "文档章节进度",
      sections: [
        {
          id: "sec_intro",
          number: "1",
          title: "项目概述",
          headingLevel: 1,
          status: "pending"
        },
        {
          id: "sec_design",
          number: "2",
          title: "密码应用设计",
          headingLevel: 1,
          status: "pending"
        }
      ],
      createdAt: "2026-05-25T00:00:00.000Z"
    });

    expect(item.title).toBe("文档章节进度");
    expect(item.total).toBe(2);
    expect(item.sections.map((section) => section.id)).toEqual(["sec_intro", "sec_design"]);

    const completed = applySchemeProgressUpdate(
      item,
      { section: "sec_design", status: "completed", detail: "密码应用设计 章节草稿已生成" },
      "2026-05-25T00:01:00.000Z"
    );

    expect(completed.completed).toBe(1);
    expect(completed.sections[1]).toMatchObject({
      status: "completed",
      detail: "密码应用设计 章节草稿已生成"
    });
  });

  it("extracts template section anchors from tool output", () => {
    expect(extractTemplateAnchorIds("模板锚点：sec_2_1、sec_2_2_2、sec_2_1")).toEqual(["sec_2_1", "sec_2_2_2"]);
  });
});
