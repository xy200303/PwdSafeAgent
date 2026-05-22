import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import mammoth from "mammoth";
import PizZip from "pizzip";
import { describe, expect, it } from "vitest";
import {
  buildSchemeTemplateData,
  extractSchemeFacts,
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
});
