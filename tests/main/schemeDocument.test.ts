import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import mammoth from "mammoth";
import PizZip from "pizzip";
import { describe, expect, it } from "vitest";
import {
  buildSchemeTemplateData,
  createWordDocxFromTemplate,
  extractSchemeFacts,
  replaceTemplatePlaceholders,
  replaceWordSectionContent,
  replaceWordSectionsContent,
  renderFactSummaryMarkdown,
  updateWordTemplateContent,
  writeSchemeDocxFromTemplate
} from "../../src/main/schemeDocument";

const TEMPLATE_BODY_STYLE_ID = "30";
const HEADING_OR_CAPTION_STYLE_PATTERN = /<w:pStyle\b[^>]*\bw:val="(?:2|3|4|12)"/;

async function readDocumentXml(filePath: string): Promise<string> {
  const zip = new PizZip(await readFile(filePath, "binary"));
  return zip.file("word/document.xml")?.asText() ?? "";
}

function findParagraphXmlContaining(documentXml: string, text: string): string {
  const markerIndex = documentXml.indexOf(text);
  expect(markerIndex).toBeGreaterThan(0);
  const paragraphStart = documentXml.lastIndexOf("<w:p", markerIndex);
  const paragraphEnd = documentXml.indexOf("</w:p>", markerIndex) + "</w:p>".length;
  expect(paragraphStart).toBeGreaterThanOrEqual(0);
  expect(paragraphEnd).toBeGreaterThan(paragraphStart);
  return documentXml.slice(paragraphStart, paragraphEnd);
}

function findSdtXmlContaining(documentXml: string, tag: string): string {
  const markerIndex = documentXml.indexOf(tag);
  expect(markerIndex).toBeGreaterThan(0);
  const sdtStart = documentXml.lastIndexOf("<w:sdt", markerIndex);
  const sdtEnd = documentXml.indexOf("</w:sdt>", markerIndex) + "</w:sdt>".length;
  expect(sdtStart).toBeGreaterThanOrEqual(0);
  expect(sdtEnd).toBeGreaterThan(sdtStart);
  return documentXml.slice(sdtStart, sdtEnd);
}

function expectTemplateBodyParagraphStyle(paragraphXml: string): void {
  expect(paragraphXml).toContain(`<w:pStyle w:val="${TEMPLATE_BODY_STYLE_ID}"`);
  expect(paragraphXml).not.toMatch(HEADING_OR_CAPTION_STYLE_PATTERN);
  expect(paragraphXml).not.toMatch(/<w:b(?:\s|\/|>)/);
}

describe("schemeDocument", () => {
  it("copies the official template byte-for-byte when creating a Word file without fields", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-create-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const outputPath = join(dir, "模板副本.docx");

    try {
      const result = await createWordDocxFromTemplate(templatePath, outputPath);
      const templateContent = await readFile(templatePath);
      const outputContent = await readFile(outputPath);

      expect(result.fileName).toBe("模板副本.docx");
      expect(result.templateReplacementCount).toBe(0);
      expect(outputContent.equals(templateContent)).toBe(true);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("uses the template JSON to replace a precise table cell while preserving the template document", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-cell-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const outputPath = join(dir, "表格单元格替换.docx");

    try {
      const result = await createWordDocxFromTemplate(templatePath, outputPath, {
        templateCells: [
          {
            tableId: "table_4_2_2_1",
            rowIndex: 1,
            columnIndex: 1,
            value: "杭州市西湖区核心数据中心"
          }
        ]
      });
      const zip = new PizZip(await readFile(outputPath, "binary"));
      const documentXml = zip.file("word/document.xml")?.asText() ?? "";
      const extracted = await mammoth.extractRawText({ path: outputPath });

      expect(result.templateCellReplacementCount).toBe(1);
      expect(result.templateReplacementCount).toBe(0);
      expect(extracted.value).toContain("杭州市西湖区核心数据中心");
      expect(documentXml).toContain("杭州市西湖区核心数据中心");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("keeps cleaned table cell formatting when replacing table placeholders", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-clean-cell-style-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const outputPath = join(dir, "清洁表格样式替换.docx");

    try {
      await createWordDocxFromTemplate(templatePath, outputPath, {
        templateCells: [
          {
            tableId: "table_36_7",
            rowIndex: 1,
            columnIndex: 4,
            value: "自定义合规措施样式标记"
          }
        ]
      });
      const documentXml = await readDocumentXml(outputPath);
      const paragraphXml = findParagraphXmlContaining(documentXml, "自定义合规措施样式标记");
      const markerIndex = paragraphXml.indexOf("自定义合规措施样式标记");
      const runXml = paragraphXml.slice(
        paragraphXml.lastIndexOf("<w:r", markerIndex),
        paragraphXml.indexOf("</w:r>", markerIndex) + "</w:r>".length
      );

      expect(paragraphXml).toContain('<w:pStyle w:val="54"');
      expect(paragraphXml).toContain('<w:ind w:firstLine="422"');
      expect(runXml).toContain('w:ascii="Times New Roman"');
      expect(runXml).toContain('<w:sz w:val="21"');
      expect(runXml).not.toMatch(/<w:b(?:\s|\/|>)/);
      expect(paragraphXml).not.toContain("【待填写】");
      expect(documentXml).not.toContain("密码应用措施");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("resolves STD content control tags through the template mapping while preserving the outer Word control", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-content-control-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const outputPath = join(dir, "控件局部替换.docx");

    try {
      const result = await updateWordTemplateContent(templatePath, outputPath, {
        contentControls: [
          {
            tag: "STD_SEC_1_1_BODY",
            value: [
              "统一身份认证系统通过 Content Control 完成局部替换。",
              "| 控件 | 说明 |",
              "| --- | --- |",
              "| STD | 兼容旧标签 |"
            ].join("\n")
          }
        ]
      });
      const documentXml = await readDocumentXml(outputPath);
      const sdtXml = findSdtXmlContaining(documentXml, "ps:section:sec_1_1:body");
      const paragraphXml = findParagraphXmlContaining(documentXml, "统一身份认证系统通过 Content Control 完成局部替换");

      expect(result.contentControlReplacementCount).toBe(1);
      expect(result.templateReplacementCount).toBe(0);
      expect(result.templateCellReplacementCount).toBe(0);
      expectTemplateBodyParagraphStyle(paragraphXml);
      expect(sdtXml).toContain("ps:section:sec_1_1:body");
      expect(sdtXml).toContain("统一身份认证系统通过 Content Control 完成局部替换");
      expect(sdtXml).toContain("<w:tbl>");
      expect(sdtXml).toContain("兼容旧标签");
      expect(documentXml).toContain("法律法规要求");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("resolves template field block ids through the template mapping for precise front-matter edits", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-field-block-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const outputPath = join(dir, "模板块局部替换.docx");

    try {
      const result = await updateWordTemplateContent(templatePath, outputPath, {
        contentControls: [
          {
            tag: "field_block_front_9",
            value: "统一身份认证系统商用密码应用方案"
          }
        ]
      });
      const documentXml = await readDocumentXml(outputPath);
      const sdtXml = findSdtXmlContaining(documentXml, "ps:field-block:field_block_front_9");

      expect(result.contentControlReplacementCount).toBe(1);
      expect(sdtXml).toContain("ps:field-block:field_block_front_9");
      expect(sdtXml).toContain("统一身份认证系统商用密码应用方案");
      expect(documentXml).toContain("统一身份认证系统商用密码应用方案");
      expect(sdtXml).not.toContain("{应用系统}密码应用方案");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("updates a section text block without removing the section's nested figure anchors", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-text-block-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const outputPath = join(dir, "章节正文块局部替换.docx");

    try {
      const result = await updateWordTemplateContent(templatePath, outputPath, {
        contentControls: [
          {
            tag: "sec_2_2_2_text_1",
            value: "本节局部正文块已单独更新，保留网络框架图和其他嵌套控件。"
          }
        ]
      });
      const documentXml = await readDocumentXml(outputPath);
      const sdtXml = findSdtXmlContaining(documentXml, "ps:section:sec_2_2_2:text:1");
      const paragraphXml = findParagraphXmlContaining(documentXml, "本节局部正文块已单独更新");

      expect(result.contentControlReplacementCount).toBe(1);
      expectTemplateBodyParagraphStyle(paragraphXml);
      expect(sdtXml).toContain("ps:section:sec_2_2_2:text:1");
      expect(sdtXml).toContain("本节局部正文块已单独更新");
      expect(documentXml).toContain('w:val="ps:figure:fig_1_2_2_2_1:image"');
      expect(documentXml).toContain('w:val="ps:section:sec_2_2_2:body"');
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("replaces one template section while keeping the surrounding document", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-section-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const draftPath = join(dir, "增量方案.docx");

    try {
      await createWordDocxFromTemplate(templatePath, draftPath);
      const result = await replaceWordSectionContent(draftPath, draftPath, {
        section: "1.1",
        content: [
          "统一身份认证系统按照商用密码应用要求开展建设。",
          "本节内容由增量写入工具替换，模板中的其他章节保持不变。"
        ].join("\n")
      });
      const extracted = await mammoth.extractRawText({ path: result.outputPath });

      expect(result.matchedHeading).toContain("1.1");
      expect(result.matchedHeading).toContain("系统建设规划");
      expect(result.templateAnchorId).toBe("sec_1_1");
      expect(result.replacementCount).toBeGreaterThan(0);
      expect(extracted.value).toContain("本节内容由增量写入工具替换");
      expect(extracted.value).toContain("法律法规要求");
      expect(extracted.value).not.toContain("密码是保障网络与信息安全的核心技术和基础支撑");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("keeps child headings when replacing a parent template section", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-parent-section-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const draftPath = join(dir, "父子章节增量方案.docx");

    try {
      await createWordDocxFromTemplate(templatePath, draftPath);
      await replaceWordSectionContent(draftPath, draftPath, {
        section: "1.2",
        content: "本节概述密码应用方案编制所依据的法律法规和政策标准。"
      });
      const result = await replaceWordSectionContent(draftPath, draftPath, {
        section: "1.2.1",
        content: "本系统建设运行应依据网络安全法落实网络安全保护义务。"
      });
      const extracted = await mammoth.extractRawText({ path: result.outputPath });

      expect(result.matchedHeading).toContain("1.2.1");
      expect(result.templateAnchorId).toBe("sec_1_2_1");
      expect(extracted.value).toContain("本节概述密码应用方案编制所依据的法律法规和政策标准。");
      expect(extracted.value).toContain("本系统建设运行应依据网络安全法落实网络安全保护义务。");
      expect(extracted.value).toContain("《网络安全等级保护条例》");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("uses the template body style when replacing an empty parent section anchor", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-parent-style-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const draftPath = join(dir, "父章节正文样式.docx");

    try {
      await createWordDocxFromTemplate(templatePath, draftPath);
      await replaceWordSectionContent(draftPath, draftPath, {
        section: "sec_1_2",
        content: "- 标准规范类：GB/T 39786-2021《信息安全技术 信息系统密码应用基本要求》。"
      });
      const documentXml = await readDocumentXml(draftPath);
      const paragraphXml = findParagraphXmlContaining(documentXml, "标准规范类");

      expectTemplateBodyParagraphStyle(paragraphXml);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("keeps every generated list and normal paragraph in the template body style", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-list-style-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const draftPath = join(dir, "列表正文样式.docx");

    try {
      await createWordDocxFromTemplate(templatePath, draftPath);
      await replaceWordSectionContent(draftPath, draftPath, {
        section: "sec_1_2",
        content: [
          "- 标准规范类：GB/T 39786-2021《信息安全技术 信息系统密码应用基本要求》。",
          "- 项目技术文件类：《某市智慧政务服务平台网络安全等级保护定级报告》。",
          "普通正文段落也应保持模板正文样式。"
        ].join("\n")
      });
      const documentXml = await readDocumentXml(draftPath);

      for (const marker of ["标准规范类", "项目技术文件类", "普通正文段落也应保持模板正文样式"]) {
        expectTemplateBodyParagraphStyle(findParagraphXmlContaining(documentXml, marker));
      }
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("does not preserve a previous bad heading style when rewriting a section anchor", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-bad-style-rewrite-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const draftPath = join(dir, "错误样式重写.docx");

    try {
      await createWordDocxFromTemplate(templatePath, draftPath);
      const zip = new PizZip(await readFile(draftPath, "binary"));
      const documentXml = zip.file("word/document.xml")?.asText() ?? "";
      const badParagraph = [
        "<w:p>",
        '<w:pPr><w:pStyle w:val="3"/></w:pPr>',
        '<w:r><w:t>旧错误样式</w:t></w:r>',
        "</w:p>"
      ].join("");
      zip.file(
        "word/document.xml",
        documentXml.replace(
          /(<w:tag w:val="ps:section:sec_1_2:body"\/><\/w:sdtPr><w:sdtContent>)[\s\S]*?(<\/w:sdtContent><\/w:sdt>)/,
          `$1${badParagraph}$2`
        )
      );
      await writeFile(draftPath, zip.generate({ type: "nodebuffer", compression: "DEFLATE" }));

      await replaceWordSectionContent(draftPath, draftPath, {
        section: "sec_1_2",
        content: "重写后应恢复模板正文样式。"
      });
      const rewrittenXml = await readDocumentXml(draftPath);
      const paragraphXml = findParagraphXmlContaining(rewrittenXml, "重写后应恢复模板正文样式");

      expectTemplateBodyParagraphStyle(paragraphXml);
      expect(rewrittenXml).not.toContain("旧错误样式");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("keeps body style for every generated section in a batch replacement", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-batch-style-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const draftPath = join(dir, "批量正文样式.docx");

    try {
      await createWordDocxFromTemplate(templatePath, draftPath);
      await replaceWordSectionsContent(draftPath, draftPath, {
        sections: [
          {
            section: "sec_1_2",
            content: "- 批量父章节列表正文样式标记。"
          },
          {
            section: "sec_8_1",
            content: "实施内容正文样式标记。"
          },
          {
            section: "sec_7",
            content: "第七章正文样式标记。"
          }
        ]
      });
      const documentXml = await readDocumentXml(draftPath);

      for (const marker of ["批量父章节列表正文样式标记", "实施内容正文样式标记", "第七章正文样式标记"]) {
        expectTemplateBodyParagraphStyle(findParagraphXmlContaining(documentXml, marker));
      }
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("uses heading style only for markdown subheadings and body style for following text", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-heading-body-style-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const draftPath = join(dir, "标题正文样式分离.docx");

    try {
      await createWordDocxFromTemplate(templatePath, draftPath);
      await replaceWordSectionContent(draftPath, draftPath, {
        section: "sec_1_2",
        content: ["### 1.2.1 子标题样式标记", "标题后的正文样式标记。"].join("\n")
      });
      const documentXml = await readDocumentXml(draftPath);
      const headingParagraphXml = findParagraphXmlContaining(documentXml, "子标题样式标记");
      const bodyParagraphXml = findParagraphXmlContaining(documentXml, "标题后的正文样式标记");

      expect(headingParagraphXml).not.toContain(`<w:pStyle w:val="${TEMPLATE_BODY_STYLE_ID}"`);
      expectTemplateBodyParagraphStyle(bodyParagraphXml);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("replaces multiple template sections in one docx pass", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-batch-section-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const draftPath = join(dir, "批量章节增量方案.docx");

    try {
      await createWordDocxFromTemplate(templatePath, draftPath);
      const result = await replaceWordSectionsContent(draftPath, draftPath, {
        sections: [
          {
            section: "sec_1_1",
            content: "第一节通过批量章节写入完成。"
          },
          {
            section: "sec_1_2_1",
            content: "第二个章节也在同一次 Word 打开保存中完成。"
          },
          {
            section: "sec_7",
            content: "第七章按真实模板章节 sec_7 写入，不使用虚拟 7.2。"
          }
        ]
      });
      const extracted = await mammoth.extractRawText({ path: result.outputPath });

      expect(result.sections.map((section) => section.templateAnchorId)).toEqual(["sec_1_1", "sec_1_2_1", "sec_7"]);
      expect(extracted.value).toContain("第一节通过批量章节写入完成。");
      expect(extracted.value).toContain("第二个章节也在同一次 Word 打开保存中完成。");
      expect(extracted.value).toContain("第七章按真实模板章节 sec_7 写入");
      expect(extracted.value).toContain("《网络安全等级保护条例》");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("uses invisible template anchors when JSON block indexes and titles drift", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-invisible-section-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const builtInTemplateJsonPath = join(process.cwd(), "docs", "密码应用方案.template.json");
    const driftedTemplateJsonPath = join(dir, "drifted-template.json");
    const draftPath = join(dir, "锚点增量方案.docx");

    try {
      const templateJson = JSON.parse(await readFile(builtInTemplateJsonPath, "utf-8"));
      const section = templateJson.sections.find((item: { id: string }) => item.id === "sec_1_2_1");
      expect(section?.anchors?.body?.tag).toBe("ps:section:sec_1_2_1:body");
      Object.assign(section, {
        number: "9.9.9",
        title: "错误标题",
        headingBlock: 0,
        bodyRange: [0, 0]
      });
      await writeFile(driftedTemplateJsonPath, `${JSON.stringify(templateJson, null, 2)}\n`, "utf-8");

      await createWordDocxFromTemplate(templatePath, draftPath);
      const result = await replaceWordSectionContent(draftPath, draftPath, {
        section: "sec_1_2_1",
        content: "本节通过 Word 不可见模板锚精确定位并完成替换。",
        templateJsonPath: driftedTemplateJsonPath
      });
      const extracted = await mammoth.extractRawText({ path: result.outputPath });

      expect(result.matchedHeading).toContain("1.2.1");
      expect(result.templateAnchorId).toBe("sec_1_2_1");
      expect(extracted.value).toContain("本节通过 Word 不可见模板锚精确定位并完成替换。");
      expect(extracted.value).toContain("《网络安全等级保护条例》");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("uses invisible template anchors even when Word heading text changes", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-heading-drift-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const draftPath = join(dir, "标题漂移增量方案.docx");

    try {
      await createWordDocxFromTemplate(templatePath, draftPath);
      const zip = new PizZip(await readFile(draftPath, "binary"));
      const documentXml = zip.file("word/document.xml")?.asText() ?? "";
      zip.file("word/document.xml", documentXml.replaceAll("《中国人民共和国网络安全法》", "用户手工改过的标题"));
      await writeFile(draftPath, zip.generate({ type: "nodebuffer", compression: "DEFLATE" }));

      const result = await replaceWordSectionContent(draftPath, draftPath, {
        section: "sec_1_2_1",
        content: "即使标题文字被用户改动，也通过不可见模板锚点完成替换。"
      });
      const extracted = await mammoth.extractRawText({ path: result.outputPath });

      expect(result.matchedHeading).toContain("1.2.1");
      expect(result.templateAnchorId).toBe("sec_1_2_1");
      expect(extracted.value).toContain("用户手工改过的标题");
      expect(extracted.value).toContain("即使标题文字被用户改动，也通过不可见模板锚点完成替换。");
      expect(extracted.value).not.toContain("网络运营者开展经营和服务活动，必须遵守法律、行政法规");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("rejects section replacement when the invisible body anchor is missing", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-missing-anchor-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const draftPath = join(dir, "锚点缺失方案.docx");

    try {
      await createWordDocxFromTemplate(templatePath, draftPath);
      const zip = new PizZip(await readFile(draftPath, "binary"));
      const documentXml = zip.file("word/document.xml")?.asText() ?? "";
      expect(documentXml).toContain('w:val="ps:section:sec_1_2_1:body"');
      zip.file(
        "word/document.xml",
        documentXml.replace('w:val="ps:section:sec_1_2_1:body"', 'w:val="ps:section:removed_sec_1_2_1:body"')
      );
      await writeFile(draftPath, zip.generate({ type: "nodebuffer", compression: "DEFLATE" }));

      await expect(
        replaceWordSectionContent(draftPath, draftPath, {
          section: "sec_1_2_1",
          content: "锚点缺失时不应退回标题编号搜索。"
        })
      ).rejects.toThrow("未找到 Word 章节：sec_1_2_1");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("does not update template table cells when the invisible table anchor is missing", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-missing-table-anchor-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const draftPath = join(dir, "表格锚点缺失模板.docx");
    const outputPath = join(dir, "表格锚点缺失输出.docx");

    try {
      await createWordDocxFromTemplate(templatePath, draftPath);
      const zip = new PizZip(await readFile(draftPath, "binary"));
      const documentXml = zip.file("word/document.xml")?.asText() ?? "";
      expect(documentXml).toContain('w:val="ps:table:table_4_2_2_1"');
      zip.file(
        "word/document.xml",
        documentXml.replace('w:val="ps:table:table_4_2_2_1"', 'w:val="ps:table:removed_table_4_2_2_1"')
      );
      await writeFile(draftPath, zip.generate({ type: "nodebuffer", compression: "DEFLATE" }));

      const result = await createWordDocxFromTemplate(draftPath, outputPath, {
        templateCells: [
          {
            tableId: "table_4_2_2_1",
            rowIndex: 1,
            columnIndex: 1,
            value: "锚点缺失时不应按题注或序号兜底写入"
          }
        ]
      });
      const outputXml = await readDocumentXml(outputPath);

      expect(result.templateCellReplacementCount).toBe(0);
      expect(outputXml).not.toContain("锚点缺失时不应按题注或序号兜底写入");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("rejects sections that do not exist in the template JSON", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-missing-section-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const draftPath = join(dir, "不存在章节.docx");

    try {
      await createWordDocxFromTemplate(templatePath, draftPath);

      await expect(
        replaceWordSectionContent(draftPath, draftPath, {
          section: "7.2",
          content: "不存在的第七章子节不应被兜底写入。"
        })
      ).rejects.toThrow("未找到 Word 章节：7.2");

      await expect(
        replaceWordSectionContent(draftPath, draftPath, {
          section: "sec_7_2",
          content: "不存在的第七章子节不应被兜底写入。"
        })
      ).rejects.toThrow("未找到 Word 章节：sec_7_2");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("splits numbered markdown and replaces only matching template sections", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-precise-section-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const draftPath = join(dir, "精准增量方案.docx");

    try {
      await createWordDocxFromTemplate(templatePath, draftPath);
      await replaceWordSectionContent(draftPath, draftPath, {
        section: "1",
        content: [
          "# 1. 背景",
          "## 1.1. 项目背景",
          "这是精准写入的项目背景正文。",
          "## 1.2. 编制依据",
          "这是精准写入的编制依据正文。",
          "# 2. 系统概述",
          "## 2.1. 业务场景描述",
          "这段第二章内容不应写入第一章替换结果。"
        ].join("\n")
      });
      const extracted = await mammoth.extractRawText({ path: draftPath });

      expect(extracted.value).toContain("这是精准写入的项目背景正文。");
      expect(extracted.value).toContain("这是精准写入的编制依据正文。");
      expect(extracted.value).toContain("系统概述");
      expect(extracted.value).not.toContain("这段第二章内容不应写入第一章替换结果。");
      expect(extracted.value).not.toContain("项目背景\n这是精准写入的项目背景正文。");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("renders markdown tables as real Word tables during section replacement", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-section-table-"));
    const templatePath = join(process.cwd(), "docs", "密码应用方案.docx");
    const draftPath = join(dir, "表格增量方案.docx");

    try {
      await createWordDocxFromTemplate(templatePath, draftPath);
      await replaceWordSectionContent(draftPath, draftPath, {
        section: "1.1",
        content: [
          "表格前正文样式标记。",
          "| 字段 | 内容 |",
          "| --- | --- |",
          "| 系统名称 | 智慧医疗大数据共享平台 |",
          "| 等保级别 | 三级 |",
          "表格后正文样式标记。"
        ].join("\n")
      });
      const documentXml = await readDocumentXml(draftPath);
      const extracted = await mammoth.extractRawText({ path: draftPath });

      expect(documentXml).toMatch(/<w:tbl[\s\S]*字段[\s\S]*智慧医疗大数据共享平台[\s\S]*<\/w:tbl>/);
      expect(extracted.value).toContain("智慧医疗大数据共享平台");
      expect(extracted.value).not.toContain("| 字段 | 内容 |");
      expectTemplateBodyParagraphStyle(findParagraphXmlContaining(documentXml, "表格前正文样式标记"));
      expectTemplateBodyParagraphStyle(findParagraphXmlContaining(documentXml, "表格后正文样式标记"));
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("renders a numbered markdown document into the Word style template by sections", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-md-template-"));
    const outputPath = join(dir, "MD内容模板方案.docx");

    try {
      const result = await writeSchemeDocxFromTemplate(
        join(process.cwd(), "docs", "密码应用方案.docx"),
        outputPath,
        {
          prompt: "系统名称：智慧医疗大数据共享平台\n建设单位：示例卫健委",
          memory: "",
          generatedMarkdown: [
            "# 1. 背景",
            "## 1.1. 系统建设规划",
            "这是由模板 JSON 指引写入的系统建设规划正文样式标记。",
            "## 1.2. 法律法规要求",
            "| 法规 | 适用要求 |",
            "| --- | --- |",
            "| 密码法 | 商用密码合规建设 |",
            "# 2. 系统概述",
            "## 2.1. 基本情况",
            "模板章节整体渲染正文样式标记。",
            "| 字段 | 内容 |",
            "| --- | --- |",
            "| 系统名称 | 智慧医疗大数据共享平台 |"
          ].join("\n"),
          renderMode: "template_sections"
        }
      );
      const zip = new PizZip(await readFile(outputPath, "binary"));
      const documentXml = zip.file("word/document.xml")?.asText() ?? "";
      const extracted = await mammoth.extractRawText({ path: result.outputPath });

      expect(result.renderMode).toBe("template_sections");
      expect(result.templateAnchorsUsed).toEqual(expect.arrayContaining(["sec_1_1", "sec_1_2", "sec_2_1"]));
      expect(extracted.value).toContain("这是由模板 JSON 指引写入的系统建设规划正文样式标记。");
      expect(extracted.value).toContain("智慧医疗大数据共享平台");
      expect(extracted.value).toContain("系统概述");
      expect(extracted.value).not.toContain("密码是保障网络与信息安全的核心技术和基础支撑");
      expect(documentXml).toMatch(/<w:tbl[\s\S]*密码法[\s\S]*商用密码合规建设[\s\S]*<\/w:tbl>/);
      expect(documentXml).toMatch(/<w:tbl[\s\S]*系统名称[\s\S]*智慧医疗大数据共享平台[\s\S]*<\/w:tbl>/);
      expectTemplateBodyParagraphStyle(findParagraphXmlContaining(documentXml, "系统建设规划正文样式标记"));
      expectTemplateBodyParagraphStyle(findParagraphXmlContaining(documentXml, "模板章节整体渲染正文样式标记"));
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("builds template data from prompt and generated content", () => {
    const data = buildSchemeTemplateData({
      prompt: "系统名称：统一身份认证系统\n建设单位：示例政务服务中心\n单位省份：广东省\n等保级别：三级",
      memory: "",
      generatedMarkdown: "本方案用于统一身份认证系统。"
    });

    expect(data["应用系统"]).toBe("统一身份认证系统");
    expect(data["建设单位"]).toBe("示例政务服务中心");
    expect(data["单位省份"]).toBe("广东省");
    expect(data["等保级别"]).toBe("三级");
  });

  it("lets explicit template fields override inferred values", () => {
    const data = buildSchemeTemplateData({
      prompt: "系统名称：统一身份认证系统\n建设单位：示例政务服务中心",
      memory: "",
      generatedMarkdown: "本方案用于统一身份认证系统。",
      templateFields: [
        { key: "{应用系统}", value: "智慧园区综合管理平台" },
        { key: "cloudPlatform", value: "政务云专有区" }
      ],
      fields: {
        constructionUnit: "示例科技有限公司",
        subsystems: ["统一门户", "权限中心"],
        machineRooms: [{ name: "核心机房", owner: "示例科技有限公司", address: "广州市天河区 1 号" }]
      }
    });

    expect(data["应用系统"]).toBe("智慧园区综合管理平台");
    expect(data["建设单位"]).toBe("示例科技有限公司");
    expect(data["应用子系统1"]).toBe("统一门户");
    expect(data["物理机房1"]).toBe("核心机房");
    expect(data["云平台"]).toBe("政务云专有区");
  });

  it("extracts a structured scheme fact model", () => {
    const facts = extractSchemeFacts({
      prompt: [
        "系统名称：统一身份认证系统",
        "建设单位：示例政务服务中心",
        "单位省份：广东省",
        "等保级别：三级",
        "子系统：认证服务、权限管理",
        "部署模式：B/S 架构",
        "数据库：达梦"
      ].join("\n"),
      memory: "关键数据包括身份鉴别数据、重要业务数据。用户包括业务用户、管理用户。",
      generatedMarkdown: "部署密码服务管理平台、服务器密码机、签名验签服务器。"
    });

    expect(facts.systemName).toBe("统一身份认证系统");
    expect(facts.organizationName).toBe("示例政务服务中心");
    expect(facts.subsystems).toEqual(["认证服务", "权限管理"]);
    expect(facts.keyDataTypes).toContain("身份鉴别数据");
    expect(facts.cryptoProducts).toContain("服务器密码机");
    expect(renderFactSummaryMarkdown(facts)).toContain("应用系统：统一身份认证系统");
  });

  it("writes a docx from the official scheme template", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-"));
    const outputPath = join(dir, "统一身份认证系统密码应用方案.docx");

    try {
      const result = await writeSchemeDocxFromTemplate(
        join(process.cwd(), "docs", "密码应用方案.docx"),
        outputPath,
        {
          prompt: "系统名称：统一身份认证系统\n建设单位：示例政务服务中心\n单位省份：广东省",
          memory: "",
          generatedMarkdown: [
            "## 方案摘要",
            "本方案采用密码服务平台、安全认证网关和服务器密码机。",
            "## 密码应用设计",
            "- 建设统一密码服务能力",
            "| 对象 | 措施 |",
            "| --- | --- |",
            "| 重要业务数据 | 加密存储 |"
          ].join("\n")
        }
      );
      const extracted = await mammoth.extractRawText({ path: result.outputPath });

      expect(result.fileName).toBe("统一身份认证系统密码应用方案.docx");
      expect(result.filledFields).toContain("应用系统");
      expect(result.appendedMarkdown).toBe(true);
      expect(result.facts.systemName).toBe("统一身份认证系统");
      expect(extracted.value).toContain("统一身份认证系统");
      expect(extracted.value).toContain("示例政务服务中心");
      expect(extracted.value).toContain("Agent 生成方案正文");
      expect(extracted.value).toContain("密码应用设计");
      expect(extracted.value).toContain("重要业务数据");
      expect(extracted.value).toContain("结构化事实摘要");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("embeds generated diagram images into the docx package", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-diagram-"));
    const outputPath = join(dir, "带图密码应用方案.docx");
    const diagramPath = join(dir, "architecture.png");

    try {
      await writeFile(
        diagramPath,
        Buffer.from(
          "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
          "base64"
        )
      );

      const result = await writeSchemeDocxFromTemplate(
        join(process.cwd(), "docs", "密码应用方案.docx"),
        outputPath,
        {
          prompt: "系统名称：统一身份认证系统\n建设单位：示例政务服务中心",
          memory: "",
          generatedMarkdown: "## 技术架构\n方案包含密码应用技术架构图。",
          diagrams: [
            {
              label: "密码应用技术架构图",
              kind: "architecture",
              path: diagramPath
            }
          ]
        }
      );
      const zip = new PizZip(await readFile(outputPath, "binary"));
      const documentXml = zip.file("word/document.xml")?.asText() ?? "";
      const relsXml = zip.file("word/_rels/document.xml.rels")?.asText() ?? "";
      const contentTypesXml = zip.file("[Content_Types].xml")?.asText() ?? "";
      const extracted = await mammoth.extractRawText({ path: result.outputPath });

      expect(result.embeddedDiagrams).toEqual(["密码应用技术架构图"]);
      expect(zip.file("word/media/diagram-1.png")).toBeTruthy();
      expect(zip.file(/^word\/media\/.*[\u4e00-\u9fa5].*$/)).toHaveLength(0);
      expect(documentXml).toContain("<w:drawing>");
      expect(documentXml).toContain("密码应用技术架构图");
      expect(documentXml).not.toContain("方案图示");
      expect(relsXml).toContain("relationships/image");
      expect(relsXml).toContain("Target=\"media/");
      expect(contentTypesXml).toContain('Extension="png"');
      expect(extracted.value).not.toContain("方案图示");
      expect(documentXml).toContain("密码应用技术框架");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("embeds generated diagram images by exact template figure id", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-figure-id-"));
    const outputPath = join(dir, "精确图位嵌入方案.docx");
    const diagramPath = join(dir, "storage-flow.png");

    try {
      await writeFile(
        diagramPath,
        Buffer.from(
          "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
          "base64"
        )
      );

      const result = await writeSchemeDocxFromTemplate(
        join(process.cwd(), "docs", "密码应用方案.docx"),
        outputPath,
        {
          prompt: "系统名称：统一身份认证系统\n建设单位：示例政务服务中心",
          memory: "",
          generatedMarkdown: "",
          diagrams: [
            {
              label: "自定义标签也应按 figure_id 精确嵌入",
              kind: "flow",
              path: diagramPath,
              figureId: "fig_12_5_4_9_4"
            }
          ]
        }
      );
      const zip = new PizZip(await readFile(outputPath, "binary"));
      const documentXml = zip.file("word/document.xml")?.asText() ?? "";
      const targetFigureXml = findSdtXmlContaining(documentXml, "ps:figure:fig_12_5_4_9_4:image");

      expect(result.embeddedDiagrams).toEqual(["自定义标签也应按 figure_id 精确嵌入"]);
      expect(targetFigureXml).toContain("<w:drawing>");
      expect(documentXml).toContain("图 510");
      expect(documentXml).toContain("重要数据存储保护流程图");
      expect(targetFigureXml).not.toContain("【图片占位】");
      expect(documentXml).not.toContain("方案图示");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("does not append unmatched diagrams when the template has figure anchors", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-unmatched-diagram-"));
    const outputPath = join(dir, "未匹配图位不追加.docx");
    const diagramPath = join(dir, "unmatched.png");

    try {
      await writeFile(
        diagramPath,
        Buffer.from(
          "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
          "base64"
        )
      );

      const result = await writeSchemeDocxFromTemplate(
        join(process.cwd(), "docs", "密码应用方案.docx"),
        outputPath,
        {
          prompt: "系统名称：统一身份认证系统\n建设单位：示例政务服务中心",
          memory: "",
          generatedMarkdown: "",
          diagrams: [
            {
              label: "模板中不存在的随手配图",
              kind: "architecture",
              path: diagramPath
            }
          ]
        }
      );
      const documentXml = await readDocumentXml(outputPath);

      expect(result.embeddedDiagrams).toEqual([]);
      expect(documentXml).not.toContain("方案图示");
      expect(documentXml).not.toContain("模板中不存在的随手配图");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("converts svg diagrams to png media before embedding into Word", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-svg-diagram-"));
    const outputPath = join(dir, "SVG转PNG图示方案.docx");
    const diagramPath = join(dir, "architecture.svg");

    try {
      await writeFile(
        diagramPath,
        [
          '<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360" viewBox="0 0 640 360">',
          '<rect width="640" height="360" fill="#ffffff"/>',
          '<rect x="60" y="90" width="180" height="80" fill="#dbeafe" stroke="#1d4ed8"/>',
          '<rect x="400" y="90" width="180" height="80" fill="#dcfce7" stroke="#15803d"/>',
          '<line x1="240" y1="130" x2="400" y2="130" stroke="#334155" stroke-width="4"/>',
          '<text x="150" y="138" text-anchor="middle" font-size="22">应用系统</text>',
          '<text x="490" y="138" text-anchor="middle" font-size="22">密码服务</text>',
          "</svg>"
        ].join(""),
        "utf-8"
      );

      const result = await writeSchemeDocxFromTemplate(
        join(process.cwd(), "docs", "密码应用方案.docx"),
        outputPath,
        {
          prompt: "系统名称：统一身份认证系统\n建设单位：示例政务服务中心",
          memory: "",
          generatedMarkdown: "## 技术架构\n方案包含密码应用技术架构图。",
          diagrams: [
            {
              label: "密码应用技术架构图",
              kind: "architecture",
              path: diagramPath
            }
          ]
        }
      );
      const zip = new PizZip(await readFile(outputPath, "binary"));
      const documentXml = zip.file("word/document.xml")?.asText() ?? "";
      const relsXml = zip.file("word/_rels/document.xml.rels")?.asText() ?? "";
      const contentTypesXml = zip.file("[Content_Types].xml")?.asText() ?? "";

      expect(result.embeddedDiagrams).toEqual(["密码应用技术架构图"]);
      expect(zip.file("word/media/diagram-1.png")).toBeTruthy();
      expect(zip.file(/^word\/media\/.*[\u4e00-\u9fa5].*$/)).toHaveLength(0);
      expect(zip.file(/^word\/media\/.*\.svg$/)).toHaveLength(0);
      expect(documentXml).toContain("<w:drawing>");
      expect(documentXml).toContain('r:embed="');
      expect(relsXml).toContain('Target="media/diagram-1.png"');
      expect(relsXml).not.toContain(".svg");
      expect(contentTypesXml).toContain('Extension="png"');
      expect(contentTypesXml).not.toContain('Extension="svg"');
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("replaces official template placeholders without leaving template markers", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-docx-fields-"));
    const outputPath = join(dir, "智慧园区综合管理平台密码应用方案.docx");

    try {
      const result = await writeSchemeDocxFromTemplate(
        join(process.cwd(), "docs", "密码应用方案.docx"),
        outputPath,
        {
          prompt: "系统名称：统一身份认证系统\n建设单位：示例政务服务中心",
          memory: "",
          generatedMarkdown: "## 方案摘要\n本方案用于智慧园区综合管理平台。",
          templateFields: [
            { key: "{应用系统}", value: "智慧园区综合管理平台" },
            { key: "建设单位", value: "示例科技有限公司" },
            { key: "单位省份", value: "浙江省" },
            { key: "单位地址", value: "杭州市西湖区 88 号" },
            { key: "单位邮编", value: "310000" },
            { key: "物理机房1", value: "核心机房" },
            { key: "物理机房1地址", value: "杭州市西湖区数据中心" },
            { key: "物理机房1管理单位", value: "示例科技有限公司" },
            { key: "应用子系统1", value: "统一门户" },
            { key: "应用子系统2", value: "权限中心" },
            { key: "云平台", value: "政务云专有区" }
          ]
        }
      );
      const zip = new PizZip(await readFile(outputPath, "binary"));
      const documentXml = zip.file("word/document.xml")?.asText() ?? "";
      const extracted = await mammoth.extractRawText({ path: result.outputPath });

      expect(result.templateReplacementCount).toBeGreaterThan(10);
      expect(extracted.value).toContain("智慧园区综合管理平台");
      expect(extracted.value).toContain("示例科技有限公司");
      expect(extracted.value).toContain("核心机房");
      expect(documentXml).not.toContain("[[PS:field:应用系统]]");
      expect(documentXml).not.toContain("[[PS:field:物理机房1地址]]");
      expect(documentXml).not.toContain("[[PS:field:云平台]]");
      expect(documentXml).not.toContain("${应用系统}");
      expect(documentXml).not.toContain("{应用系统}");
      expect(documentXml).not.toContain("{物理机房1地址}");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("replaces docx-templates field markers", async () => {
    const zip = new PizZip();
    zip.file(
      "[Content_Types].xml",
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    );
    zip.file(
      "word/document.xml",
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>',
        "<w:r><w:t>{应用系统}</w:t></w:r>",
        "<w:r><w:t>涉及</w:t></w:r>",
        "<w:r><w:t>{九条}</w:t></w:r>",
        "<w:r><w:t>通信信道</w:t></w:r>",
        "</w:p></w:body></w:document>"
      ].join("")
    );

    const count = await replaceTemplatePlaceholders(zip, {
      应用系统: "统一身份认证系统",
      九条: "六条"
    });
    const xml = zip.file("word/document.xml")?.asText() ?? "";

    expect(count).toBe(2);
    expect(xml).toContain("统一身份认证系统");
    expect(xml).toContain("六条");
    expect(xml).not.toContain("{应用系统}");
    expect(xml).not.toContain("{九条}");
  });

  it("keeps unspecified docx-templates fields intact during partial rendering", async () => {
    const zip = new PizZip();
    zip.file(
      "[Content_Types].xml",
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    );
    zip.file(
      "word/document.xml",
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>',
        "<w:r><w:t>{应用系统}</w:t></w:r>",
        "<w:r><w:t>由</w:t></w:r>",
        "<w:r><w:t>{建设单位}</w:t></w:r>",
        "<w:r><w:t>建设</w:t></w:r>",
        "</w:p></w:body></w:document>"
      ].join("")
    );

    const count = await replaceTemplatePlaceholders(zip, {
      应用系统: "统一身份认证系统"
    });
    const xml = zip.file("word/document.xml")?.asText() ?? "";

    expect(count).toBe(1);
    expect(xml).toContain("统一身份认证系统");
    expect(xml).toContain("{建设单位}");
  });

  it("replaces placeholders split across Word text nodes", async () => {
    const zip = new PizZip();
    zip.file(
      "[Content_Types].xml",
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    );
    zip.file(
      "word/document.xml",
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>',
        "<w:r><w:t>{应用</w:t></w:r>",
        "<w:r><w:t>系统}</w:t></w:r>",
        "<w:r><w:t>由</w:t></w:r>",
        "<w:r><w:t>{建设</w:t></w:r>",
        "<w:r><w:t>单位}</w:t></w:r>",
        "</w:p></w:body></w:document>"
      ].join("")
    );

    const count = await replaceTemplatePlaceholders(zip, {
      应用系统: "统一身份认证系统",
      建设单位: "示例政务服务中心"
    });
    const xml = zip.file("word/document.xml")?.asText() ?? "";

    expect(count).toBe(2);
    expect(xml).toContain("统一身份认证系统");
    expect(xml).toContain("示例政务服务中心");
    expect(xml).not.toContain("{应用");
    expect(xml).not.toContain("{建设");
  });
});
