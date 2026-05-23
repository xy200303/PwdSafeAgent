import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import mammoth from "mammoth";
import type { ChatCompletionMessageToolCall } from "openai/resources/chat/completions";
import { describe, expect, it } from "vitest";
import {
  buildAgentChatTools,
  executeAgentToolCall,
  parseDuckDuckGoHtml,
  parseToolArguments,
  searchWeb,
  type AgentToolExecutionContext
} from "../../src/main/agentToolRegistry";
import type { AppSettings } from "../../src/shared/types";

describe("agentToolRegistry", () => {
  it("builds tool definitions and keeps exec_bash opt-in", () => {
    const safeTools = buildAgentChatTools({ includeExecBash: false });
    const fullTools = buildAgentChatTools({ includeExecBash: true });

    expect(safeTools.map((tool) => tool.function.name)).toContain("remember_project");
    expect(safeTools.map((tool) => tool.function.name)).toContain("web_search");
    expect(safeTools.map((tool) => tool.function.name)).toContain("read_word");
    expect(safeTools.map((tool) => tool.function.name)).toContain("read_pdf");
    expect(safeTools.map((tool) => tool.function.name)).toContain("write_word");
    expect(safeTools.map((tool) => tool.function.name)).not.toContain("exec_bash");
    expect(fullTools.map((tool) => tool.function.name)).toContain("exec_bash");
    expect(fullTools.every((tool) => tool.type === "function" && tool.function.strict === true)).toBe(true);
  });

  it("can hide final artifact tools while collecting project information", () => {
    const tools = buildAgentChatTools({ includeExecBash: false, includeArtifactTools: false });
    const names = tools.map((tool) => tool.function.name);

    expect(names).toContain("remember_project");
    expect(names).toContain("read_word");
    expect(names).toContain("write_file");
    expect(names).not.toContain("write_word");
    expect(names).not.toContain("write_pdf");
    expect(names).not.toContain("image_generate");
    expect(names).not.toContain("send_file");
  });

  it("parses tool arguments defensively", () => {
    expect(parseToolArguments('{"query":"GB/T 39786"}')).toEqual({ query: "GB/T 39786" });
    expect(parseToolArguments("not-json")).toEqual({});
    expect(parseToolArguments("[1,2,3]")).toEqual({});
  });

  it("parses DuckDuckGo html search results", () => {
    const html = `
      <html><body>
        <div class="result">
          <a class="result__a" href="/l/?kh=-1&amp;uddg=https%3A%2F%2Fexample.com%2Fstandard">GB/T 39786 标准</a>
          <div class="result__snippet">密码应用基本要求。</div>
        </div>
      </body></html>
    `;

    expect(parseDuckDuckGoHtml(html)).toEqual([
      {
        title: "GB/T 39786 标准",
        url: "https://example.com/standard",
        snippet: "密码应用基本要求。"
      }
    ]);
  });

  it("records project memory facts without producing artifacts", async () => {
    const result = await executeAgentToolCall(
      createToolCall("remember_project", {
        summary: "已确认统一身份认证系统基础信息",
        facts: [
          { key: "应用系统", value: "统一身份认证系统" },
          { key: "建设单位", value: "示例政务服务中心" }
        ],
        gaps: ["单位地址", "等保级别"],
        ready_for_generation: false
      }),
      createContext(process.cwd())
    );

    expect(result.toolName).toBe("remember_project");
    expect(result.summary).toContain("继续收集");
    expect(result.content).toContain("应用系统：统一身份认证系统");
    expect(result.artifactPath).toBeUndefined();
  });

  it("formats web search results from an injected fetch implementation", async () => {
    const html = `
      <div class="result">
        <a class="result__a" href="https://example.com/a">结果 A</a>
        <a class="result__snippet" href="https://example.com/a">摘要 A</a>
      </div>
    `;
    const results = await searchWeb("密码应用", {
      fetchImpl: async () => new Response(html, { status: 200, statusText: "OK" })
    });

    expect(results[0]?.title).toBe("结果 A");
    expect(results[0]?.url).toBe("https://example.com/a");
  });

  it("executes write_file into the output directory", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-"));

    try {
      const result = await executeAgentToolCall(createToolCall("write_file", { name: "清单.md", content: "# 检查清单" }), {
        ...createContext(dir),
        outputDir: dir
      });

      expect(result.toolName).toBe("write_file");
      expect(result.artifactPath).toBeTruthy();
      await expect(readFile(result.artifactPath!, "utf-8")).resolves.toContain("检查清单");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("sends an existing output file when the model only provides the file name", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-send-"));
    const outputDir = join(dir, "output");
    const outputPath = join(outputDir, "方案.md");

    try {
      await mkdir(outputDir, { recursive: true });
      await writeFile(outputPath, "# 方案", "utf-8");
      const result = await executeAgentToolCall(createToolCall("send_file", { path: "方案.md" }), {
        ...createContext(dir),
        outputDir
      });

      expect(result.toolName).toBe("send_file");
      expect(result.artifactPath).toBe(outputPath);
      expect(result.summary).toContain("方案.md");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("rejects send_file when the target output file does not exist", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-send-missing-"));

    try {
      await expect(executeAgentToolCall(createToolCall("send_file", { path: "missing.md" }), createContext(dir))).rejects.toThrow(
        "文件不存在"
      );
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("refuses read_file outside allowed directories", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-"));

    try {
      await expect(
        executeAgentToolCall(createToolCall("read_file", { path: join(tmpdir(), "outside.txt") }), createContext(dir))
      ).rejects.toThrow("只能访问允许目录");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("allows model-selected reads for user-picked files explicitly whitelisted by the session", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-allowed-"));
    const externalDir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-external-"));
    const externalPath = join(externalDir, "需求说明.txt");

    try {
      await writeFile(externalPath, "系统名称：统一身份认证系统", "utf-8");
      const result = await executeAgentToolCall(createToolCall("read_file", { path: externalPath }), {
        ...createContext(dir),
        allowedReadFiles: [externalPath]
      });

      expect(result.toolName).toBe("read_file");
      expect(result.content).toContain("统一身份认证系统");
    } finally {
      await rm(dir, { recursive: true, force: true });
      await rm(externalDir, { recursive: true, force: true });
    }
  });

  it("validates explicit read_word and read_pdf extensions", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-read-"));
    const textPath = join(dir, "note.txt");

    try {
      await writeFile(textPath, "hello", "utf-8");
      const result = await executeAgentToolCall(createToolCall("read_word", { path: textPath }), createContext(dir));

      expect(result.toolName).toBe("read_word");
      expect(result.summary).toContain(".docx");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("skips image generation when the user did not explicitly request diagrams", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-image-"));

    try {
      const result = await executeAgentToolCall(
        createToolCall("image_generate", {
          kind: "architecture",
          prompt: "生成密码应用方案配图"
        }),
        createContext(dir)
      );

      expect(result.toolName).toBe("image_generate");
      expect(result.summary).toContain("已跳过");
      expect(result.artifactPath).toBeUndefined();
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("skips image generation when the image tool is disabled in settings", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-image-disabled-"));

    try {
      const settings = createSettings();
      settings.openai.autoImageGeneration = false;
      const result = await executeAgentToolCall(
        createToolCall("image_generate", {
          kind: "architecture",
          prompt: "生成密码应用技术架构图"
        }),
        {
          ...createContext(dir),
          settings,
          userPrompt: "请生成密码应用方案和技术架构图"
        }
      );

      expect(result.toolName).toBe("image_generate");
      expect(result.summary).toContain("关闭");
      expect(result.artifactPath).toBeUndefined();
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("executes write_word with the official scheme template", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-word-"));

    try {
      const result = await executeAgentToolCall(
        createToolCall("write_word", {
          name: "统一身份认证系统密码应用方案.docx",
          prompt: "系统名称：统一身份认证系统\n建设单位：示例政务服务中心\n单位省份：广东省",
          content: "## 方案摘要\n本方案通过密码服务管理平台和服务器密码机提供密码应用能力。"
        }),
        {
          ...createContext(process.cwd()),
          outputDir: dir
        }
      );
      const extracted = await mammoth.extractRawText({ path: result.artifactPath! });

      expect(result.toolName).toBe("write_word");
      expect(result.artifactPath).toBeTruthy();
      expect(result.summary).toContain("已生成");
      expect(extracted.value).toContain("统一身份认证系统");
      expect(extracted.value).toContain("方案摘要");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});

function createToolCall(name: string, args: Record<string, unknown>): ChatCompletionMessageToolCall {
  return {
    id: "call_1",
    type: "function",
    function: {
      name,
      arguments: JSON.stringify(args)
    }
  };
}

function createContext(dir: string): AgentToolExecutionContext {
  return {
    rootDir: dir,
    docsDir: join(dir, "docs"),
    outputDir: join(dir, "output"),
    sessionTitle: "统一身份认证系统",
    memory: "",
    settings: createSettings(),
    allowedReadDirs: [dir],
    execBashEnabled: false
  };
}

function createSettings(): AppSettings {
  return {
    runtime: {
      envFilePath: "",
      configSource: "process",
      bundledPython: {
        available: false
      }
    },
    openai: {
      baseUrl: "https://api.openai.com/v1",
      imageBaseUrl: "",
      chatModel: "gpt-5.5",
      imageModel: "gpt-image-2",
      imageSize: "1536x1024",
      imageQuality: "high",
      autoImageGeneration: true,
      requestTimeoutMs: 120000,
      maxOutputTokens: 16000,
      apiKeyConfigured: false,
      imageApiKeyConfigured: false
    },
    document: {
      autoPdfExport: false,
      libreOfficePath: ""
    },
    agent: {
      runtime: "openai-chat",
      execBashEnabled: false
    }
  };
}
