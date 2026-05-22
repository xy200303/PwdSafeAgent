import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import {
  compactText,
  getReadToolName,
  readDocumentText,
  sanitizeFileName,
  writeUtf8File
} from "../../src/main/agentTools";

describe("agentTools", () => {
  it("maps file extensions to built-in read tools", () => {
    expect(getReadToolName("template.docx")).toBe("read_word");
    expect(getReadToolName("report.pdf")).toBe("read_pdf");
    expect(getReadToolName("notes.md")).toBe("read_file");
  });

  it("sanitizes unsafe file names", () => {
    expect(sanitizeFileName('  密码<应用>:"方案"?.md  ')).toBe("密码_应用___方案__.md");
    expect(sanitizeFileName("   ")).toBe("artifact");
  });

  it("compacts long text and keeps a truncation hint", () => {
    const compacted = compactText("第一段\n\n\n\n第二段" + "x".repeat(20), 10);

    expect(compacted).toContain("第一段\n\n第二");
    expect(compacted).toContain("内容过长");
  });

  it("reads supported text documents", async () => {
    const result = await readDocumentText(join(process.cwd(), "tests", "fixtures", "sample.txt"), 200);

    expect(result.toolName).toBe("read_file");
    expect(result.sourceName).toBe("sample.txt");
    expect(result.content).toContain("示例密码应用系统");
    expect(result.summary).toContain("已读取 sample.txt");
  });

  it("writes utf-8 files and creates parent directories", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-"));
    const output = join(dir, "nested", "方案草稿.md");

    try {
      await writeUtf8File(output, "# 密码应用方案\n\n测试内容");
      await expect(readFile(output, "utf-8")).resolves.toContain("测试内容");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
