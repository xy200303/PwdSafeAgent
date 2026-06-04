import { existsSync } from "node:fs";
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
import {
  DRAFT_SECTION_PARALLELISM_DEFAULT,
  IMAGE_GENERATION_PARALLELISM_DEFAULT,
  IMAGE_GENERATION_REQUEST_TIMEOUT_DEFAULT_MS,
  clampDraftSectionParallelism,
  clampImageGenerationParallelism,
  type AppSettings,
  type SchemeProgressItem
} from "../../src/shared/types";

const BUILT_IN_TEMPLATE_DOCX_PATH = join(process.cwd(), "resources", "docs", "templates", "密码应用方案.docx");
const BUILT_IN_TEMPLATE_JSON_PATH = join(process.cwd(), "resources", "docs", "templates", "密码应用方案.template.json");

describe("agentToolRegistry", () => {
  it("defaults drafting to twenty parallel tasks and image generation to ten", () => {
    expect(DRAFT_SECTION_PARALLELISM_DEFAULT).toBe(20);
    expect(clampDraftSectionParallelism(undefined)).toBe(20);
    expect(clampDraftSectionParallelism(99)).toBe(20);
    expect(IMAGE_GENERATION_REQUEST_TIMEOUT_DEFAULT_MS).toBe(300000);
    expect(IMAGE_GENERATION_PARALLELISM_DEFAULT).toBe(10);
    expect(clampImageGenerationParallelism(undefined)).toBe(10);
    expect(clampImageGenerationParallelism(99)).toBe(10);
  });

  it("builds tool definitions and keeps exec_bash opt-in", () => {
    const safeTools = buildAgentChatTools({ includeExecBash: false });
    const fullTools = buildAgentChatTools({ includeExecBash: true });
    const writeWordTool = safeTools.find((tool) => tool.function.name === "write_word");

    expect(safeTools.map((tool) => tool.function.name)).toContain("remember_project");
    expect(safeTools.map((tool) => tool.function.name)).toContain("web_search");
    expect(safeTools.map((tool) => tool.function.name)).toContain("read_word");
    expect(safeTools.map((tool) => tool.function.name)).toContain("read_pdf");
    expect(safeTools.map((tool) => tool.function.name)).toContain("plan_scheme_batches");
    expect(safeTools.map((tool) => tool.function.name)).toContain("plan_scheme_assets");
    expect(safeTools.map((tool) => tool.function.name)).toContain("draft_scheme_sections");
    expect(safeTools.map((tool) => tool.function.name)).toContain("create_word");
    expect(safeTools.map((tool) => tool.function.name)).toContain("write_word");
    expect(safeTools.map((tool) => tool.function.name)).not.toContain("exec_bash");
    expect(fullTools.map((tool) => tool.function.name)).toContain("exec_bash");
    expect(fullTools.every((tool) => tool.type === "function" && tool.function.strict === true)).toBe(true);
    expect(writeWordTool?.function.description).toContain("局部正文块 textBlocks");
    expect(JSON.stringify(writeWordTool)).toContain("sec_2_2_2_text_1");
    expect(JSON.stringify(writeWordTool)).toContain("template_tables");
  });

  it("can hide final artifact tools while collecting project information", () => {
    const tools = buildAgentChatTools({ includeExecBash: false, includeArtifactTools: false });
    const names = tools.map((tool) => tool.function.name);

    expect(names).toContain("remember_project");
    expect(names).toContain("read_word");
    expect(names).toContain("write_file");
    expect(names).not.toContain("plan_scheme_batches");
    expect(names).not.toContain("plan_scheme_assets");
    expect(names).not.toContain("draft_scheme_sections");
    expect(names).not.toContain("create_word");
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

  it("summarizes the built-in Word template JSON as section planning tasks", async () => {
    const result = await executeAgentToolCall(
      createToolCall("read_file", { path: "docs/templates/密码应用方案.template.json" }),
      createContext(process.cwd())
    );

    expect(result.toolName).toBe("read_file");
    expect(result.summary).toContain("规划任务清单");
    expect(result.content).toContain("draft_scheme_sections 每批尽量传 20 个 section");
    expect(result.content).toContain("content_controls[].tag 可直接传 fieldBlocks/textBlocks 的 id");
    expect(result.content).toContain("sec_2_2_2 | 2.2.2 网络环境");
    expect(result.content).toContain("task: 正文：描述网络整体结构");
    expect(result.content).toContain("网络通道/通信信道按“访问者通过网络访问系统”的形式定义");
    expect(result.content).toContain("按“访问者通过网络访问系统”的形式定义网络通道/通信信道");
    expect(result.content).toContain("局部正文块：sec_2_2_2_text_1（tag：ps:section:sec_2_2_2:text:1）");
    expect(result.content).toContain("段落：说明本节范围和已确认对象");
    expect(result.content).toContain("表格：表 22 物理环境情况");
    expect(result.content).toContain("图示：网络框架图");
    expect(result.content).not.toContain("图示：fig_1_2_2_2_1");
    expect(result.content).toContain("固定模板块");
    expect(result.content).toContain("field_block_front_9");
    expect(result.content).not.toContain("\"schemaVersion\"");
  });

  it("does not support the legacy top-level template json path", async () => {
    await expect(
      executeAgentToolCall(
        createToolCall("read_file", { path: "docs/密码应用方案.template.json" }),
        createContext(process.cwd())
      )
    ).rejects.toThrow();
  });

  it("plans stable scheme section batches from the template json", async () => {
    const result = await executeAgentToolCall(
      createToolCall("plan_scheme_batches", {
        start_section: "sec_2_2_2",
        batch_size: 5,
        completed_sections: ["sec_2_2_2_1"]
      }),
      createContext(process.cwd())
    );

    expect(result.toolName).toBe("plan_scheme_batches");
    expect(result.summary).toContain("首批 5 个章节");
    expect(result.content).toContain("first_draft_call");
    expect(result.content).toContain("\"max_parallel\": 5");
    expect(result.content).toContain("\"section\": \"sec_2_2_2\"");
    expect(result.content).toContain("\"paragraph_tasks\"");
    expect(result.content).not.toContain("\"section\": \"sec_2_2_2_1\"");
    expect(result.content).toContain("BATCH 1 (5)");
  });

  it("uses natural figure names instead of template ids in section paragraph tasks", async () => {
    const result = await executeAgentToolCall(
      createToolCall("plan_scheme_batches", {
        start_section: "sec_5_4_9_4",
        batch_size: 1
      }),
      createContext(process.cwd())
    );

    expect(result.toolName).toBe("plan_scheme_batches");
    expect(result.content).toContain("重要数据存储保护流程图");
    expect(result.content).toContain("重要数据存储读取流程图");
    expect(result.content).not.toContain("为后续 fig_12_5_4_9_4");
    expect(result.content).not.toContain("为后续 fig_13_5_4_9_4");
  });

  it("skips already drafted or completed sections when planning batches", async () => {
    const progress: SchemeProgressItem = {
      id: "scheme_progress_test",
      kind: "scheme_progress",
      title: "方案章节生成",
      status: "running",
      total: 3,
      drafted: 1,
      completed: 1,
      failed: 0,
      createdAt: "2026-05-27T00:00:00.000Z",
      sections: [
        { id: "sec_2_2_2", number: "2.2.2", title: "网络环境", headingLevel: 3, status: "completed" },
        { id: "sec_2_2_2_1", number: "2.2.2.1", title: "网络框架", headingLevel: 4, status: "drafted" }
      ]
    };

    const result = await executeAgentToolCall(
      createToolCall("plan_scheme_batches", {
        start_section: "sec_2_2_2",
        batch_size: 3
      }),
      {
        ...createContext(process.cwd()),
        schemeProgress: progress
      }
    );

    expect(result.toolName).toBe("plan_scheme_batches");
    expect(result.content).toContain("已跳过章节数：2");
    expect(result.content).not.toContain("\"section\": \"sec_2_2_2\"");
    expect(result.content).not.toContain("\"section\": \"sec_2_2_2_1\"");
    expect(result.content).toContain("\"section\": \"sec_2_2_2_2\"");
  });

  it("does not silently fall back when planning from an unknown section", async () => {
    const result = await executeAgentToolCall(
      createToolCall("plan_scheme_batches", {
        start_section: "sec_7_2",
        batch_size: 3
      }),
      createContext(process.cwd())
    );

    expect(result.toolName).toBe("plan_scheme_batches");
    expect(result.summary).toContain("起始章节不在模板中");
    expect(result.content).toContain("unknown start_section sec_7_2");
    expect(result.content).not.toContain("\"section\": \"sec_1\"");
  });

  it("does not ignore unknown completed sections when planning batches", async () => {
    const result = await executeAgentToolCall(
      createToolCall("plan_scheme_batches", {
        completed_sections: ["sec_404"]
      }),
      createContext(process.cwd())
    );

    expect(result.toolName).toBe("plan_scheme_batches");
    expect(result.summary).toContain("跳过章节不在模板中");
    expect(result.content).toContain("unknown completed_sections sec_404");
    expect(result.content).not.toContain("first_draft_call");
  });

  it("plans table cells and figure tasks for selected scheme sections", async () => {
    const result = await executeAgentToolCall(
      createToolCall("plan_scheme_assets", {
        section_ids: ["sec_5_4_9_4"],
        max_items: 10
      }),
      createContext(process.cwd())
    );

    expect(result.toolName).toBe("plan_scheme_assets");
    expect(result.summary).toContain("表格 3 个、图示 2 个");
    expect(result.content).toContain("template_tables_plan");
    expect(result.content).toContain("template_cells_plan");
    expect(result.content).toContain("本批单元格");
    expect(result.content).toContain("\"table_id\": \"table_30_5_4_9_4\"");
    expect(result.content).toContain("\"caption\": \"表 5-15 数据存储的保护对象\"");
    expect(result.content).toContain("\"table_id\": \"table_31_5_4_9_4\"");
    expect(result.content).toContain("\"recommended_write\": \"template_tables\"");
    expect(result.content).toContain("\"figure_id\": \"fig_12_5_4_9_4\"");
    expect(result.content).toContain("\"label\": \"重要数据存储保护流程图\"");
    expect(result.content).toContain("write_word_diagrams_plan");
    expect(result.content).toContain("模板题注仅供匹配参考");
    expect(result.content).toContain("图内不要单独放图号、题注或标题");
  });

  it("paginates full template table cell planning to avoid truncating later tables", async () => {
    const result = await executeAgentToolCall(
      createToolCall("plan_scheme_assets", {
        include_figures: false,
        max_cells: 25,
        max_figures: 0
      }),
      createContext(process.cwd())
    );

    expect(result.toolName).toBe("plan_scheme_assets");
    expect(result.content).toContain("template_tables_plan");
    expect(result.content).toContain("可填单元格：272 个；本批单元格：25 个");
    expect(result.content).toContain("next_plan_scheme_assets_call");
    expect(result.content).toContain("\"cell_offset\": 25");
    expect(result.content).toContain("\"include_tables\": true");
    expect(result.content).toContain("\"include_figures\": false");
    expect(result.content).not.toContain("内容过长，已截断");
  });

  it("rejects asset planning for unknown sections", async () => {
    const result = await executeAgentToolCall(
      createToolCall("plan_scheme_assets", {
        section_ids: ["sec_7_2"]
      }),
      createContext(process.cwd())
    );

    expect(result.toolName).toBe("plan_scheme_assets");
    expect(result.summary).toContain("章节不在模板中");
    expect(result.content).toContain("unknown section_ids sec_7_2");
  });

  it("drafts scheme sections without writing Word artifacts", async () => {
    const result = await executeAgentToolCall(
      createToolCall("draft_scheme_sections", {
        sections: [
          { section: "2.1" },
          { section: "2.2.2", title: "网络环境" }
        ],
        project_context: "系统名称：统一身份认证系统；建设单位：示例政务服务中心",
        max_parallel: 2
      }),
      {
        ...createContext(process.cwd()),
        memory: "等保级别：三级"
      }
    );

    expect(result.toolName).toBe("draft_scheme_sections");
    expect(result.summary).toContain("已并行起草 2");
    expect(result.artifactPath).toBeUndefined();
    expect(result.content).toContain("按 paragraph_tasks 分段后的正文草稿");
    expect(result.content).toContain("## 2.1 基本情况");
    expect(result.content).toContain("## 2.2.2 网络环境");
    expect(result.content).toContain("本节结论将为");
    expect(result.schemeProgressUpdates).toEqual([
      expect.objectContaining({ section: "2.1", status: "drafted" }),
      expect.objectContaining({ section: "2.2.2", status: "drafted" })
    ]);
  });

  it("keeps raw template figure ids out of fallback draft content", async () => {
    const result = await executeAgentToolCall(
      createToolCall("draft_scheme_sections", {
        sections: [
          {
            section: "sec_5_4_9_4",
            paragraph_tasks: [
              "说明重要数据存储保护措施",
              "为后续 fig_12_5_4_9_4、fig_13_5_4_9_4 图示生成提供场景说明，不在正文中生成图片"
            ]
          }
        ],
        max_parallel: 1
      }),
      createContext(process.cwd())
    );

    expect(result.toolName).toBe("draft_scheme_sections");
    expect(result.content).toContain("重要数据存储保护流程图");
    expect(result.content).toContain("重要数据存储读取流程图");
    expect(result.content).not.toContain("fig_12_5_4_9_4");
    expect(result.content).not.toContain("fig_13_5_4_9_4");
    expect(result.content).not.toContain("本节关联模板图示");
  });

  it("uses the configured draft section parallelism when no override is provided", async () => {
    const settings = createSettings();
    settings.agent.draftSectionParallelism = 3;

    const result = await executeAgentToolCall(
      createToolCall("draft_scheme_sections", {
        sections: [{ section: "2.1" }, { section: "2.2.2" }],
        project_context: "系统名称：统一身份认证系统"
      }),
      {
        ...createContext(process.cwd()),
        settings
      }
    );

    expect(result.toolName).toBe("draft_scheme_sections");
    expect(result.content).toContain("并行度：3");
  });

  it("drafts up to twenty sections in one batch", async () => {
    const template = JSON.parse(
      await readFile(BUILT_IN_TEMPLATE_JSON_PATH, "utf-8")
    ) as { sections: Array<{ id: string }> };
    const sections = template.sections.slice(0, 21).map((section) => ({ section: section.id }));

    expect(sections.length).toBeGreaterThan(20);

    const result = await executeAgentToolCall(
      createToolCall("draft_scheme_sections", {
        sections,
        project_context: "系统名称：统一身份认证系统"
      }),
      createContext(process.cwd())
    );

    expect(result.toolName).toBe("draft_scheme_sections");
    expect(result.summary).toContain("已并行起草 20");
    expect(result.content).toContain("并行度：20");
    expect(result.content).toContain("本批超过 20 个章节");
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

  it("sends timestamp-prefixed generated files when the model provides the original name", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-send-prefixed-"));
    const outputDir = join(dir, "output");
    const outputPath = join(outputDir, "mpl38ooc-test_file.txt");

    try {
      await mkdir(outputDir, { recursive: true });
      await writeFile(outputPath, "ok", "utf-8");
      const result = await executeAgentToolCall(createToolCall("send_file", { path: "data/output/test_file.txt" }), {
        ...createContext(dir),
        outputDir
      });

      expect(result.toolName).toBe("send_file");
      expect(result.artifactPath).toBe(outputPath);
      expect(result.summary).toContain("mpl38ooc-test_file.txt");
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

  it("generates images whenever the image tool is called and enabled", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-image-"));
    const previousImageKey = process.env.OPENAI_IMAGE_API_KEY;
    const previousApiKey = process.env.OPENAI_API_KEY;

    try {
      delete process.env.OPENAI_IMAGE_API_KEY;
      delete process.env.OPENAI_API_KEY;

      const result = await executeAgentToolCall(
        createToolCall("image_generate", {
          kind: "architecture",
          prompt: "生成密码应用方案配图"
        }),
        createContext(dir)
      );

      expect(result.toolName).toBe("image_generate");
      expect(result.summary).toContain("已生成");
      expect(result.artifactPath).toContain(".svg");
    } finally {
      if (previousImageKey === undefined) {
        delete process.env.OPENAI_IMAGE_API_KEY;
      } else {
        process.env.OPENAI_IMAGE_API_KEY = previousImageKey;
      }
      if (previousApiKey === undefined) {
        delete process.env.OPENAI_API_KEY;
      } else {
        process.env.OPENAI_API_KEY = previousApiKey;
      }
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("can run multiple image generation tool calls concurrently with unique output files", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-image-parallel-"));
    const previousImageKey = process.env.OPENAI_IMAGE_API_KEY;
    const previousApiKey = process.env.OPENAI_API_KEY;

    try {
      delete process.env.OPENAI_IMAGE_API_KEY;
      delete process.env.OPENAI_API_KEY;

      const [first, second] = await Promise.all([
        executeAgentToolCall(
          createToolCall("image_generate", {
            kind: "architecture",
            label: "密码应用技术架构图",
            prompt: "生成密码应用方案配图"
          }),
          createContext(dir)
        ),
        executeAgentToolCall(
          createToolCall("image_generate", {
            kind: "architecture",
            label: "密码应用技术架构图",
            prompt: "生成密码应用方案配图"
          }),
          createContext(dir)
        )
      ]);

      expect(first.toolName).toBe("image_generate");
      expect(second.toolName).toBe("image_generate");
      expect(first.artifactPath).toContain(".svg");
      expect(second.artifactPath).toContain(".svg");
      expect(first.artifactPath).not.toBe(second.artifactPath);
    } finally {
      if (previousImageKey === undefined) {
        delete process.env.OPENAI_IMAGE_API_KEY;
      } else {
        process.env.OPENAI_IMAGE_API_KEY = previousImageKey;
      }
      if (previousApiKey === undefined) {
        delete process.env.OPENAI_API_KEY;
      } else {
        process.env.OPENAI_API_KEY = previousApiKey;
      }
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

  it("creates a Word file directly from the built-in template", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-create-word-"));

    try {
      const result = await executeAgentToolCall(
        createToolCall("create_word", {
          name: "模板副本.docx"
        }),
        {
          ...createContext(process.cwd()),
          outputDir: dir
        }
      );

      expect(result.toolName).toBe("create_word");
      expect(result.summary).toContain("已基于模板创建");
      expect(result.artifactPath).toBeTruthy();
      expect(Buffer.from(await readFile(result.artifactPath!)).equals(Buffer.from(await readFile(BUILT_IN_TEMPLATE_DOCX_PATH)))).toBe(
        true
      );
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("updates a single Word section without requiring the whole scheme", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-section-word-"));

    try {
      const created = await executeAgentToolCall(
        createToolCall("create_word", {
          name: "增量方案.docx"
        }),
        {
          ...createContext(process.cwd()),
          outputDir: dir
        }
      );
      const result = await executeAgentToolCall(
        createToolCall("write_word", {
          path: created.artifactPath,
          section: "1.1",
          content: "本节为统一身份认证系统的系统建设规划增量内容。"
        }),
        {
          ...createContext(process.cwd()),
          outputDir: dir
        }
      );
      const extracted = await mammoth.extractRawText({ path: result.artifactPath! });

      expect(result.toolName).toBe("write_word");
      expect(result.summary).toContain("已更新");
      expect(result.artifactPath).toBe(created.artifactPath);
      expect(extracted.value).toContain("系统建设规划");
      expect(extracted.value).toContain("统一身份认证系统的系统建设规划增量内容");
      expect(extracted.value).toContain("法律法规要求");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("updates template fields without requiring generated section content", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-field-word-"));

    try {
      const created = await executeAgentToolCall(
        createToolCall("create_word", {
          name: "字段直改方案.docx"
        }),
        {
          ...createContext(process.cwd()),
          outputDir: dir
        }
      );
      const result = await executeAgentToolCall(
        createToolCall("write_word", {
          path: created.artifactPath,
          fields: {
            应用系统: "字段直改统一身份认证系统",
            建设单位: "字段直改示例政务服务中心"
          }
        }),
        {
          ...createContext(process.cwd()),
          outputDir: dir
        }
      );
      const extracted = await mammoth.extractRawText({ path: result.artifactPath! });

      expect(result.toolName).toBe("write_word");
      expect(result.summary).toContain("模板字段/控件");
      expect(result.artifactPath).toBe(created.artifactPath);
      expect(extracted.value).toContain("字段直改统一身份认证系统");
      expect(extracted.value).toContain("字段直改示例政务服务中心");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("executes write_word with the official scheme template", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-word-"));
    const diagramPaths = [
      join(dir, "network-architecture.png"),
      join(dir, "network-topology.png"),
      join(dir, "crypto-architecture.png"),
      join(dir, "business-flow.png")
    ];

    try {
      for (const diagramPath of diagramPaths) {
        await writeFile(
          diagramPath,
          Buffer.from(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
            "base64"
          )
        );
      }

      const result = await executeAgentToolCall(
        createToolCall("write_word", {
          name: "统一身份认证系统密码应用方案.docx",
          prompt: "系统名称：统一身份认证系统\n建设单位：示例政务服务中心\n单位省份：广东省",
          content: createCompleteSchemeContent(),
          template_fields: createCompleteTemplateFields(),
          diagrams: [
            { label: "网络架构图", kind: "architecture", path: diagramPaths[0] },
            { label: "网络拓扑图", kind: "architecture", path: diagramPaths[1] },
            { label: "密码应用技术架构图", kind: "architecture", path: diagramPaths[2] },
            { label: "典型业务密码应用流程图", kind: "flow", path: diagramPaths[3] }
          ],
          render_mode: "full_document"
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
      expect(extracted.value).toContain("密码应用设计");
      expect(extracted.value).toContain("网络拓扑图");
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
  const bundledDocsDir = join(dir, "resources", "docs");
  return {
    rootDir: dir,
    docsDir: existsSync(bundledDocsDir) ? bundledDocsDir : join(dir, "docs"),
    outputDir: join(dir, "output"),
    sessionTitle: "统一身份认证系统",
    memory: "",
    settings: createSettings(),
    allowedReadDirs: [dir],
    execBashEnabled: false
  };
}

function createCompleteSchemeContent(): string {
  return [
    "# 统一身份认证系统密码应用方案",
    "## 1. 背景",
    "### 1.1. 系统建设规划",
    "统一身份认证系统面向政务服务统一认证、权限管理和日志审计场景建设。",
    "### 1.2. 法律法规要求",
    "方案依据《密码法》、GB/T 39786-2021 等要求进行规划。",
    "## 2. 系统概述",
    "### 2.1. 基本情况",
    "建设单位为示例政务服务中心，系统安全保护等级为三级。",
    "### 2.2. 计算平台现状",
    "系统采用 B/S 架构，部署在核心机房，包含应用服务器、数据库和日志审计组件。",
    "### 2.2.2. 网络环境",
    "网络划分为安全接入区、交换区、服务器区和安全运维区。",
    "图：网络架构图",
    "图：网络拓扑图",
    "### 2.3. 业务应用现状",
    "系统包含认证服务和权限管理两个应用子系统，处理身份鉴别数据、重要业务数据和日志数据。",
    "## 3. 密码应用需求分析",
    "### 3.1. 物理和环境安全",
    "需保护门禁记录和视频监控记录完整性。",
    "### 3.2. 网络和通信安全",
    "需实现通信实体身份鉴别、重要数据传输机密性和完整性保护。",
    "### 3.3. 设备和计算安全",
    "需保护远程管理通道、系统资源访问控制信息和日志记录完整性。",
    "### 3.4. 应用和数据安全",
    "需实现应用身份鉴别、访问控制完整性、重要数据传输与存储保护。",
    "## 4. 安全目标及设计原则",
    "### 4.1. 安全目标",
    "建立覆盖物理、网络、设备、应用和管理层面的密码应用保障体系。",
    "### 4.2. 设计原则和依据",
    "遵循合规性、适用性、体系化和可运维原则。",
    "## 5. 密码应用设计",
    "### 5.1. 密码应用技术框架",
    "图：密码应用技术架构图",
    "系统通过密码服务管理平台、服务器密码机、签名验签服务器和数字证书认证系统提供统一密码能力。",
    "### 5.2. 计算平台密码应用方案",
    "在网络和通信、设备和计算、应用和数据等层面部署密码防护措施。",
    "### 5.3.7. 密钥管理方式",
    "密钥管理覆盖密钥生成、分发、存储、使用、更新、归档、撤销、备份、恢复和销毁全过程。",
    "### 5.4. 业务应用的密码应用方案",
    "图：典型业务密码应用流程图",
    "业务流程包括用户登录、证书校验、签名验签、数据加密存储和日志审计。",
    "## 6. 安全管理方案",
    "### 6.1. 管理制度",
    "建立密码应用安全管理制度和密钥管理规则。",
    "### 6.2. 人员管理",
    "明确密钥管理员、密码操作员和密码审计员职责。",
    "### 6.3. 建设运行",
    "按照方案实施建设并在投运前完成密码应用安全性评估。",
    "### 6.4. 应急处置",
    "建立密码应用安全事件应急处置流程。",
    "## 7. 安全与合规性分析",
    "本方案对照 GB/T 39786-2021 对物理、网络、设备、应用和管理要求进行符合性说明。",
    "## 8. 实施保障方案",
    "### 8.1. 实施内容",
    "实施内容包括设备采购部署、系统集成改造、联调测试和试运行。",
    "### 8.2. 实施计划",
    "项目按启动、调研、设计、实施、测试、试运行和验收阶段推进。",
    "### 8.3. 保障措施",
    "通过组织、人员、经费和质量保障确保项目落地。",
    "### 8.4. 经费概算",
    "经费覆盖密码产品、集成实施、测试评估和运维保障。"
  ].join("\n\n");
}

function createCompleteTemplateFields(): Array<{ key: string; value: string }> {
  return [
    { key: "应用系统", value: "统一身份认证系统" },
    { key: "建设单位", value: "示例政务服务中心" },
    { key: "单位省份", value: "广东省" },
    { key: "单位地址", value: "广州市天河区示例路 1 号" },
    { key: "单位邮编", value: "510000" },
    { key: "等保级别", value: "三级" },
    { key: "应用子系统1", value: "认证服务" },
    { key: "应用子系统2", value: "权限管理" },
    { key: "物理机房1", value: "核心机房" },
    { key: "物理机房1管理单位", value: "示例政务服务中心" },
    { key: "物理机房1地址", value: "广州市天河区数据中心" },
    { key: "物理机房2", value: "不涉及灾备机房" },
    { key: "物理机房2管理单位", value: "不涉及" },
    { key: "物理机房2地址", value: "不涉及" },
    { key: "云平台", value: "不涉及云平台" },
    { key: "密码系统产品", value: "密码服务管理平台、服务器密码机、签名验签服务器" },
    { key: "密码安全产品", value: "密码服务管理平台、服务器密码机、签名验签服务器" }
  ];
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
      thinkingEnabled: true,
      reasoningEffort: "",
      requestTimeoutMs: 120000,
      imageRequestTimeoutMs: 300000,
      maxOutputTokens: 16000,
      apiKeyConfigured: false,
      imageApiKeyConfigured: false
    },
    document: {
      autoPdfExport: false,
      libreOfficePath: ""
    },
    agent: {
      execBashEnabled: false,
      draftSectionParallelism: 20,
      imageGenerationParallelism: 10
    }
  };
}
