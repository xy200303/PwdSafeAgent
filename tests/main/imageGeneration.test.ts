import { mkdtemp, readFile, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { buildDiagramPrompt, generateDiagramImage, shouldGenerateDiagramArtifacts } from "../../src/main/imageGeneration";

describe("imageGeneration", () => {
  it("detects scheme prompts that should generate diagram artifacts", () => {
    expect(shouldGenerateDiagramArtifacts("请生成密码应用方案", "需要技术架构设计")).toBe(true);
    expect(shouldGenerateDiagramArtifacts("只总结一下附件", "没有图表要求")).toBe(false);
    expect(shouldGenerateDiagramArtifacts("你好", "我会按密码应用方案模板推进，并可生成技术架构图。")).toBe(false);
  });

  it("builds a professional diagram prompt from scheme context", () => {
    const prompt = buildDiagramPrompt({
      kind: "architecture",
      sessionTitle: "统一身份认证系统",
      prompt: "系统名称：统一身份认证系统\n建设单位：示例政务服务中心",
      memory: "模板章节：密码应用技术架构",
      generatedMarkdown: "方案采用服务器密码机和签名验签服务器。",
      outputDir: "unused",
      artifactStamp: "test"
    });

    expect(prompt).toContain("密码应用技术架构图");
    expect(prompt).toContain("统一身份认证系统");
    expect(prompt).toContain("服务器密码机");
    expect(prompt).toContain("中文标签清晰可读");
  });

  it("writes a local svg placeholder when OpenAI image API key is not configured", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-image-"));

    try {
      const result = await generateDiagramImage(
        {
          apiKey: "",
          baseUrl: "https://api.openai.com/v1",
          imageModel: "gpt-image-2",
          imageSize: "1536x1024",
          imageQuality: "high",
          requestTimeoutMs: 120000
        },
        {
          kind: "flow",
          sessionTitle: "统一身份认证系统",
          prompt: "系统名称：统一身份认证系统\n建设单位：示例政务服务中心",
          memory: "",
          generatedMarkdown: "生成业务流程图。",
          outputDir: dir,
          artifactStamp: "test"
        }
      );
      const content = await readFile(result.outputPath, "utf-8");
      const fileStat = await stat(result.outputPath);

      expect(result.mode).toBe("local-svg");
      expect(result.fileName).toContain(".svg");
      expect(fileStat.size).toBeGreaterThan(100);
      expect(content).toContain("典型业务密码应用流程图");
      expect(content).toContain("统一身份认证系统");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
