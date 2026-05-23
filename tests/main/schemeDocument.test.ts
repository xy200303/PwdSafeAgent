import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import mammoth from "mammoth";
import PizZip from "pizzip";
import { describe, expect, it } from "vitest";
import {
  buildSchemeTemplateData,
  extractSchemeFacts,
  replaceTemplatePlaceholders,
  renderFactSummaryMarkdown,
  writeSchemeDocxFromTemplate
} from "../../src/main/schemeDocument";

describe("schemeDocument", () => {
  it("builds template data from prompt and generated content", () => {
    const data = buildSchemeTemplateData({
      prompt: "系统名称：统一身份认证系统\n建设单位：示例政务服务中心\n单位省份：广东省\n等保级别：三级",
      memory: "",
      generatedMarkdown: "本方案用于统一身份认证系统。"
    });

    expect(data["应用系统"]).toBe("统一身份认证系统");
    expect(data["$应用系统"]).toBe("统一身份认证系统");
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
        { key: "${应用系统}", value: "智慧园区综合管理平台" },
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
      expect(zip.file(/^word\/media\/diagram-1-.*\.png$/)).toHaveLength(1);
      expect(documentXml).toContain("<w:drawing>");
      expect(documentXml).toContain("方案图示");
      expect(documentXml).toContain("密码应用技术架构图");
      expect(relsXml).toContain("relationships/image");
      expect(relsXml).toContain("Target=\"media/");
      expect(contentTypesXml).toContain('Extension="png"');
      expect(extracted.value).toContain("方案图示");
      expect(extracted.value).toContain("密码应用技术架构图");
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
            { key: "应用系统", value: "智慧园区综合管理平台" },
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
      expect(documentXml).not.toContain("${应用系统}");
      expect(documentXml).not.toContain("{物理机房1地址}");
      expect(documentXml).not.toContain("${云平台}");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("replaces placeholders split across Word text nodes", () => {
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
        "<w:r><w:t>${应用</w:t></w:r>",
        "<w:r><w:t>系统}</w:t></w:r>",
        "<w:r><w:t>由</w:t></w:r>",
        "<w:r><w:t>{建设</w:t></w:r>",
        "<w:r><w:t>单位}</w:t></w:r>",
        "</w:p></w:body></w:document>"
      ].join("")
    );

    const count = replaceTemplatePlaceholders(zip, {
      应用系统: "统一身份认证系统",
      建设单位: "示例政务服务中心"
    });
    const xml = zip.file("word/document.xml")?.asText() ?? "";

    expect(count).toBe(2);
    expect(xml).toContain("统一身份认证系统");
    expect(xml).toContain("示例政务服务中心");
    expect(xml).not.toContain("${应用");
    expect(xml).not.toContain("{建设");
  });
});
