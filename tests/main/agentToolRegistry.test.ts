import { existsSync } from "node:fs";
import { copyFile, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import mammoth from "mammoth";
import PizZip from "pizzip";
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
import { registerDocumentTemplateFromDocx } from "../../src/main/documentTemplateRegistration";
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
    const buildDocumentConfigTool = safeTools.find((tool) => tool.function.name === "build_document_config");
    const writeDocumentWordTool = safeTools.find((tool) => tool.function.name === "write_document_word");

    expect(safeTools.map((tool) => tool.function.name)).toContain("remember_project");
    expect(safeTools.map((tool) => tool.function.name)).toContain("web_search");
    expect(safeTools.map((tool) => tool.function.name)).toContain("read_word");
    expect(safeTools.map((tool) => tool.function.name)).toContain("read_pdf");
    expect(safeTools.map((tool) => tool.function.name)).toContain("read_image");
    expect(safeTools.map((tool) => tool.function.name)).not.toContain("register_document_template");
    expect(safeTools.map((tool) => tool.function.name)).toContain("update_document_profile");
    expect(safeTools.map((tool) => tool.function.name)).toContain("build_document_config");
    expect(safeTools.map((tool) => tool.function.name)).toContain("list_document_sections");
    expect(safeTools.map((tool) => tool.function.name)).toContain("get_document_section");
    expect(safeTools.map((tool) => tool.function.name)).toContain("draft_document_sections");
    expect(safeTools.map((tool) => tool.function.name)).toContain("update_document_section_draft");
    expect(safeTools.map((tool) => tool.function.name)).toContain("audit_document_sections");
    expect(safeTools.map((tool) => tool.function.name)).toContain("revise_document_sections_evidence");
    expect(safeTools.map((tool) => tool.function.name)).toContain("polish_document_sections");
    expect(safeTools.map((tool) => tool.function.name)).toContain("assemble_document_sections");
    expect(safeTools.map((tool) => tool.function.name)).not.toContain("draft_document_modules");
    expect(safeTools.map((tool) => tool.function.name)).not.toContain("assemble_document_content");
    expect(safeTools.map((tool) => tool.function.name)).toContain("audit_document_evidence");
    expect(safeTools.map((tool) => tool.function.name)).toContain("revise_document_evidence");
    expect(safeTools.map((tool) => tool.function.name)).toContain("plan_document_assets");
    expect(safeTools.map((tool) => tool.function.name)).toContain("write_document_word");
    expect(safeTools.map((tool) => tool.function.name)).not.toContain("plan_scheme_batches");
    expect(safeTools.map((tool) => tool.function.name)).not.toContain("plan_scheme_assets");
    expect(safeTools.map((tool) => tool.function.name)).not.toContain("draft_scheme_sections");
    expect(safeTools.map((tool) => tool.function.name)).not.toContain("assemble_scheme_markdown");
    expect(safeTools.map((tool) => tool.function.name)).not.toContain("create_word");
    expect(safeTools.map((tool) => tool.function.name)).not.toContain("write_word");
    expect(safeTools.map((tool) => tool.function.name)).not.toContain("exec_bash");
    expect(fullTools.map((tool) => tool.function.name)).toContain("exec_bash");
    expect(fullTools.every((tool) => tool.type === "function" && tool.function.strict === true)).toBe(true);
    expect(writeDocumentWordTool?.function.description).toContain("section-first");
    expect(JSON.stringify(buildDocumentConfigTool)).toContain("generation_plan");
    expect(JSON.stringify(buildDocumentConfigTool)).toContain("section_group_plans");
    expect(safeTools.find((tool) => tool.function.name === "list_document_sections")?.function.description).toContain("section-first");
    expect(JSON.stringify(writeDocumentWordTool)).toContain("template_path");
    expect(JSON.stringify(writeDocumentWordTool)).toContain("template_json_path");
    expect(JSON.stringify(writeDocumentWordTool)).toContain("template_tables");
    expect(JSON.stringify(writeDocumentWordTool)).toContain("template_cells");
    expect(JSON.stringify(writeDocumentWordTool)).toContain("diagrams");
  });

  it("can hide final artifact tools while collecting project information", () => {
    const tools = buildAgentChatTools({ includeExecBash: false, includeArtifactTools: false });
    const names = tools.map((tool) => tool.function.name);

    expect(names).toContain("remember_project");
    expect(names).toContain("read_word");
    expect(names).toContain("read_image");
    expect(names).toContain("write_file");
    expect(names).not.toContain("register_document_template");
    expect(names).not.toContain("update_document_profile");
    expect(names).not.toContain("build_document_config");
    expect(names).not.toContain("list_document_sections");
    expect(names).not.toContain("get_document_section");
    expect(names).not.toContain("draft_document_sections");
    expect(names).not.toContain("update_document_section_draft");
    expect(names).not.toContain("audit_document_sections");
    expect(names).not.toContain("revise_document_sections_evidence");
    expect(names).not.toContain("polish_document_sections");
    expect(names).not.toContain("assemble_document_sections");
    expect(names).not.toContain("draft_document_modules");
    expect(names).not.toContain("assemble_document_content");
    expect(names).not.toContain("audit_document_evidence");
    expect(names).not.toContain("revise_document_evidence");
    expect(names).not.toContain("plan_document_assets");
    expect(names).not.toContain("write_document_word");
    expect(names).not.toContain("plan_scheme_batches");
    expect(names).not.toContain("plan_scheme_assets");
    expect(names).not.toContain("draft_scheme_sections");
    expect(names).not.toContain("assemble_scheme_markdown");
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
          { key: "应用系统", value: "统一身份认证系统", source: "用户原话" },
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
    expect(result.content).toContain("来源：用户原话");
    expect(result.content).toContain("待补充信息”必须继续保留为待补充/需确认");
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
      expect(basename(result.artifactPath!)).toMatch(/^清单-[a-z0-9]+-[a-z0-9]{6}\.md$/);
      await expect(readFile(result.artifactPath!, "utf-8")).resolves.toContain("检查清单");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("reads the built-in Word template structure JSON", async () => {
    const result = await executeAgentToolCall(
      createToolCall("read_file", { path: "docs/templates/密码应用方案.template.json" }),
      createContext(process.cwd())
    );

    expect(result.toolName).toBe("read_file");
    expect(result.summary).not.toContain("规划任务清单");
    expect(result.content).toContain("\"schemaVersion\"");
    expect(result.content).toContain("\"templateId\"");
    expect(result.content).not.toContain("draft_scheme_sections 每批尽量传 20 个 section");
  });

  it("builds a project document config artifact before drafting content", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-doc-config-"));
    const outputDir = join(dir, "output");

    try {
      const result = await executeAgentToolCall(
        createToolCall("build_document_config", {
          title: "统一身份认证系统密码应用方案",
          project_context: "系统名称：统一身份认证系统；建设单位：示例政务服务中心",
          generation_plan: "先按模板拆分章节组，再围绕已确认事实起草；缺少证据的内容保留待补充。",
          planning_assumptions: ["当前只确认系统名称和建设单位"],
          planning_risks: ["不得把等保级别写成已确认事实"],
          section_group_plans: [
            {
              section_group_id: "section_group_1",
              objective: "交代系统建设背景、边界和资料缺口。",
              outline: ["说明系统建设规划", "列出本阶段待补充信息"],
              key_points: ["应用系统名称", "建设单位"],
              evidence_needs: ["系统名称来源", "建设单位来源"],
              open_questions: ["等保级别"]
            }
          ]
        }),
        {
          ...createContext(process.cwd()),
          outputDir,
          userPrompt: [
            "办法2",
            "",
            "## 工具 read_file",
            "这是一整段不应进入 document-config.sourceSummary 的标准原文和 profile 原文。"
          ].join("\n"),
          memory: [
            "项目档案更新",
            "已确认事实：",
            "- 应用系统：统一身份认证系统（来源：用户原话）",
            "待补充信息：",
            "- 等保级别"
          ].join("\n")
        }
      );

      expect(result.toolName).toBe("build_document_config");
      expect(result.summary).toContain("document-config.json");
      expect(result.artifactPath).toBeTruthy();
      expect(result.content).toContain("sectionGroups:");
      expect(result.content).toContain("planning:");
      expect(result.content).toContain("交代系统建设背景");
      expect(result.content).toContain("应用系统：统一身份认证系统");

      const config = JSON.parse(await readFile(result.artifactPath!, "utf-8")) as {
        profile: string;
        sourceSummary: string;
        planning?: { summary: string; assumptions: string[]; risks: string[] };
        sectionGroups: Array<{
          id: string;
          writingRules: string[];
          requiredFacts: string[];
          plan?: {
            objective: string;
            outline: string[];
            keyPoints: string[];
            evidenceNeeds: string[];
            openQuestions: string[];
          };
        }>;
        facts: Array<{ key: string; value: string; source?: string }>;
        gaps: string[];
        globalRules: string[];
      };
      expect(config.profile).toBe("generic_document");
      expect(config.sourceSummary).toContain("办法2");
      expect(config.sourceSummary).toContain("系统名称：统一身份认证系统");
      expect(config.sourceSummary).not.toContain("## 工具 read_file");
      expect(config.sourceSummary).not.toContain("不应进入 document-config.sourceSummary");
      expect(config.globalRules.join("\n")).toContain("拟采用/建议采用/待确认");
      expect(config.planning?.summary).toContain("先按模板拆分章节组");
      expect(config.planning?.assumptions).toContain("当前只确认系统名称和建设单位");
      expect(config.planning?.risks).toContain("不得把等保级别写成已确认事实");
      expect(config.sectionGroups.length).toBeGreaterThan(0);
      expect(config.sectionGroups[0].plan?.objective).toBe("交代系统建设背景、边界和资料缺口。");
      expect(config.sectionGroups[0].plan?.outline).toContain("说明系统建设规划");
      expect(config.sectionGroups[0].plan?.openQuestions).toContain("等保级别");
      expect(config.sectionGroups[0].writingRules.join("\n")).toContain("不能把模板提示");
      expect(config.facts).toEqual([expect.objectContaining({ key: "应用系统", value: "统一身份认证系统", source: "用户原话" })]);
      expect(config.gaps).toContain("等保级别");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("builds a document config from a reusable document profile", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-doc-profile-"));
    const outputDir = join(dir, "output");

    try {
      const result = await executeAgentToolCall(
        createToolCall("build_document_config", {
          profile: "generic_document",
          title: "统一身份认证系统建设文档",
          project_context: "系统名称：统一身份认证系统；建设单位：示例政务服务中心；文档用途：建设方案"
        }),
        {
          ...createContext(process.cwd()),
          outputDir,
          memory: [
            "项目档案更新",
            "已确认事实：",
            "- 应用系统：统一身份认证系统（来源：用户原话）",
            "- 文档用途：建设方案（来源：用户原话）",
            "待补充信息：",
            "- 实施计划"
          ].join("\n")
        }
      );

      expect(result.toolName).toBe("build_document_config");
      expect(result.summary).toContain("8 个章节组");
      expect(result.content).toContain("profile_source: docs/document-profiles/generic_document.json");
      expect(result.content).toContain("section_group_1");
      expect(result.content).toContain("section_group_8");
      expect(result.content).toContain("密码应用设计");

      const config = JSON.parse(await readFile(result.artifactPath!, "utf-8")) as {
        profile: string;
        sectionGroups: Array<{ id: string; sectionNumbers: string[]; sectionTitles: string[]; requiredFacts: string[] }>;
        globalRules: string[];
        wordTemplate?: { templatePath?: string; templateJsonPath?: string; renderMode?: string; outputNameSuffix?: string };
      };
      expect(config.profile).toBe("generic_document");
      expect(config.sectionGroups.map((group) => group.id)).toEqual([
        "section_group_1",
        "section_group_2",
        "section_group_3",
        "section_group_4",
        "section_group_5",
        "section_group_6",
        "section_group_7",
        "section_group_8"
      ]);
      expect(config.sectionGroups[0].sectionNumbers).toContain("1.1");
      expect(config.sectionGroups[0].sectionTitles).toContain("系统建设规划");
      expect(config.sectionGroups[0].requiredFacts).toContain("等保级别或合规要求");
      expect(config.globalRules.join("\n")).toContain("当前 Word 模板解析出的章节结构");
      expect(config.wordTemplate?.templatePath).toBe("docs/templates/密码应用方案.docx");
      expect(config.wordTemplate?.templateJsonPath).toBe("docs/templates/密码应用方案.template.json");
      expect(config.wordTemplate?.renderMode).toBe("template_sections");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("rejects the built-in template json when used as profile_path", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-template-json-profile-"));
    const outputDir = join(dir, "output");

    try {
      await expect(
        executeAgentToolCall(
          createToolCall("build_document_config", {
            profile_path: "docs/templates/密码应用方案.template.json",
            title: "默认模板生成文档",
            project_context: "系统名称：统一身份认证系统"
          }),
          {
            ...createContext(process.cwd()),
            outputDir
          }
        )
      ).rejects.toThrow("无法读取有效 profile");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("builds a config from a UI-registered user Word template profile", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-register-template-"));
    const outputDir = join(dir, "output");

    try {
      const registered = await registerDocumentTemplateFromDocx({
        sourceDocxPath: BUILT_IN_TEMPLATE_DOCX_PATH,
        outputDir,
        name: "用户上传模板",
        profile: "user_template"
      });
      expect(registered.templatePathForTool).toBe("document-templates/user_template/template.docx");
      expect(registered.templateJsonPathForTool).toBe("document-templates/user_template/template.json");
      expect(registered.profilePathForTool).toBe("document-templates/user_template/profile.json");

      const templateDir = join(outputDir, "document-templates", "user_template");
      expect(existsSync(join(templateDir, "template.docx"))).toBe(true);
      expect(existsSync(join(templateDir, "template.json"))).toBe(true);
      expect(existsSync(join(templateDir, "profile.json"))).toBe(true);

      const templateJson = JSON.parse(await readFile(join(templateDir, "template.json"), "utf-8")) as {
        sections: Array<{ number: string; title: string; anchors?: { body?: { tag: string } } }>;
        source: { docx: string };
      };
      expect(templateJson.source.docx).toBe("document-templates/user_template/template.docx");
      expect(templateJson.sections.length).toBeGreaterThan(0);
      expect(templateJson.sections.map((section) => section.title)).toContain("背景");
      expect(templateJson.sections.some((section) => section.anchors?.body?.tag.startsWith("ps:section:"))).toBe(true);
      const generatedProfile = JSON.parse(await readFile(join(templateDir, "profile.json"), "utf-8")) as {
        globalRules?: string[];
        sectionRules?: Array<{ match?: string; rules?: string[] }>;
        wordTemplate?: unknown;
      };
      expect(generatedProfile.globalRules?.join("\n")).toContain("拟采用/建议采用/待确认");
      expect(generatedProfile.sectionRules?.map((rule) => rule.match)).toContain("背景");
      expect(generatedProfile.wordTemplate).toBeUndefined();

      const configResult = await executeAgentToolCall(
        createToolCall("build_document_config", {
          profile_path: "document-templates/user_template/profile.json",
          title: "用户模板生成文档",
          project_context: "系统名称：统一身份认证系统"
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );

      expect(configResult.toolName).toBe("build_document_config");
      expect(configResult.content).toContain("profile_source: document-templates/user_template/profile.json");

      const config = JSON.parse(await readFile(configResult.artifactPath!, "utf-8")) as {
        profile: string;
        sectionGroups: Array<{ id: string; sectionTitles: string[] }>;
        wordTemplate?: { templatePath?: string; templateJsonPath?: string; renderMode?: string };
      };
      expect(config.profile).toBe("user_template");
      expect(config.sectionGroups.length).toBeGreaterThan(0);
      expect(config.sectionGroups[0].sectionTitles).toContain("背景");
      expect(config.wordTemplate?.templatePath).toBe("document-templates/user_template/template.docx");
      expect(config.wordTemplate?.templateJsonPath).toBe("document-templates/user_template/template.json");
      expect(config.wordTemplate?.renderMode).toBe("template_sections");

      const assetPlan = await executeAgentToolCall(
        createToolCall("plan_document_assets", {
          section_ids: ["sec_5_4_9_4"],
          max_items: 10
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );
      expect(assetPlan.toolName).toBe("plan_document_assets");
      expect(assetPlan.content).toContain("模板来源：document-templates/user_template/template.json");
      expect(assetPlan.content).toContain("template_tables_plan");
      expect(assetPlan.content).toContain("write_document_word_diagrams_plan");

      const markdownPath = join(outputDir, "用户模板终稿.md");
      await writeFile(
        markdownPath,
        [
          "# 用户模板生成文档",
          "",
          "## 1.1 系统建设规划",
          "",
          "这是用户模板章节替换正文，来自自动注册模板的隐藏章节锚点。"
        ].join("\n"),
        "utf-8"
      );

      const wordResult = await executeAgentToolCall(
        createToolCall("write_document_word", {
          markdown_path: "用户模板终稿.md",
          name: "用户模板生成文档.docx"
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );
      const extracted = await mammoth.extractRawText({ path: wordResult.artifactPath! });

      expect(wordResult.toolName).toBe("write_document_word");
      expect(wordResult.content).toContain("写入模式：template_sections");
      expect(wordResult.content).toContain("sec_1_1");
      expect(extracted.value).toContain("这是用户模板章节替换正文");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("rejects a custom template json when used as profile_path", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-custom-template-json-profile-"));
    const outputDir = join(dir, "output");

    try {
      await registerDocumentTemplateFromDocx({
        sourceDocxPath: BUILT_IN_TEMPLATE_DOCX_PATH,
        outputDir,
        name: "用户上传模板",
        profile: "user_template"
      });

      await expect(
        executeAgentToolCall(
          createToolCall("build_document_config", {
            profile_path: "document-templates/user_template/template.json",
            title: "用户模板生成文档",
            project_context: "系统名称：统一身份认证系统"
          }),
          {
            ...createContext(process.cwd()),
            outputDir
          }
        )
      ).rejects.toThrow("无法读取有效 profile");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("reads and updates global template profiles from a session output context", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-global-template-"));
    const sessionOutputDir = join(dir, "session-output");
    const globalOutputDir = join(dir, "global-output");

    try {
      const baseContext = createContext(process.cwd());
      const context = {
        ...baseContext,
        outputDir: sessionOutputDir,
        globalOutputDir,
        allowedReadDirs: [baseContext.docsDir, sessionOutputDir, join(globalOutputDir, "document-templates")]
      };
      await registerDocumentTemplateFromDocx({
        sourceDocxPath: BUILT_IN_TEMPLATE_DOCX_PATH,
        outputDir: globalOutputDir,
        name: "全局模板",
        profile: "global_template"
      });
      expect(existsSync(join(globalOutputDir, "document-templates", "global_template", "profile.json"))).toBe(true);
      expect(existsSync(join(sessionOutputDir, "document-templates", "global_template", "profile.json"))).toBe(false);

      const profilePath = join(globalOutputDir, "document-templates", "global_template", "profile.json");
      const readProfileResult = await executeAgentToolCall(
        createToolCall("read_file", {
          path: "document-templates/global_template/profile.json"
        }),
        context
      );
      expect(readProfileResult.toolName).toBe("read_file");
      expect(readProfileResult.content).toContain('"globalRules"');

      const profile = JSON.parse(await readFile(profilePath, "utf-8")) as {
        globalRules: string[];
        sectionRules: Array<{ rules?: string[] }>;
      };
      profile.globalRules.push("输出必须包含项目假设清单。");
      profile.sectionRules[0].rules = [...(profile.sectionRules[0].rules ?? []), "每个章节组末尾列出 profile 微调后的检查点。"];

      const updateResult = await executeAgentToolCall(
        createToolCall("update_document_profile", {
          profile_path: "document-templates/global_template/profile.json",
          content: JSON.stringify(profile),
          change_summary: "增加项目假设清单和章节组检查点要求"
        }),
        context
      );

      expect(updateResult.toolName).toBe("update_document_profile");
      expect(updateResult.content).toContain("document-templates/global_template/profile.json");
      expect(await readFile(profilePath, "utf-8")).toContain("项目假设清单");

      const configResult = await executeAgentToolCall(
        createToolCall("build_document_config", {
          profile_path: "document-templates/global_template/profile.json",
          title: "全局模板生成文档",
          project_context: "系统名称：统一身份认证系统"
        }),
        context
      );

      expect(configResult.toolName).toBe("build_document_config");
      expect(configResult.artifactPath).toContain(sessionOutputDir);
      expect(configResult.content).toContain("profile_source: document-templates/global_template/profile.json");
      const config = JSON.parse(await readFile(configResult.artifactPath!, "utf-8")) as {
        globalRules: string[];
        wordTemplate?: { templatePath?: string };
      };
      expect(config.globalRules).toContain("输出必须包含项目假设清单。");
      expect(config.wordTemplate?.templatePath).toBe("document-templates/global_template/template.docx");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("indexes, updates, and assembles document section drafts", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-doc-sections-"));
    const outputDir = join(dir, "output");

    try {
      await executeAgentToolCall(
        createToolCall("build_document_config", {
          profile: "generic_document",
          title: "统一身份认证系统建设文档",
          project_context: "系统名称：统一身份认证系统；建设单位：示例政务服务中心"
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );

      const listResult = await executeAgentToolCall(
        createToolCall("list_document_sections", {
          query: "sec_5_4_9_4",
          max_sections: 5
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );

      expect(listResult.toolName).toBe("list_document_sections");
      expect(listResult.content).toContain("document-sections/manifest.json");
      expect(listResult.content).toContain("sec_5_4_9_4");
      expect(existsSync(join(outputDir, "document-sections", "manifest.json"))).toBe(true);

      const draftResult = await executeAgentToolCall(
        createToolCall("draft_document_sections", {
          section_ids: ["sec_5_4_9_4"],
          project_context: "系统名称：统一身份认证系统；建设单位：示例政务服务中心",
          max_parallel: 1
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );

      expect(draftResult.toolName).toBe("draft_document_sections");
      expect(draftResult.summary).toContain("已并行起草 1");
      expect(draftResult.content).toContain("document-sections/");
      expect(draftResult.schemeProgressUpdates?.[0]).toEqual(
        expect.objectContaining({
          section: "sec_5_4_9_4",
          status: "drafted"
        })
      );

      const updateResult = await executeAgentToolCall(
        createToolCall("update_document_section_draft", {
          section: "sec_5_4_9_4",
          content: [
            "# 5.4.9.4 数据存储",
            "",
            "本节说明统一身份认证系统重要数据存储场景的密码应用设计，未确认的产品型号保留待补充/需确认。",
            "",
            "系统已部署服务器密码机并已配置签名验签能力。"
          ].join("\n"),
          status: "edited",
          change_summary: "补充章节正文并保留产品型号缺口",
          tables_json: JSON.stringify([{ table_id: "table_30_5_4_9_4", rows: [["保护对象", "保护措施"]] }])
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );

      expect(updateResult.toolName).toBe("update_document_section_draft");
      expect(updateResult.content).toContain("状态：edited");
      expect(updateResult.content).toContain("tables.json");
      expect(updateResult.schemeProgressUpdates?.[0]).toEqual(
        expect.objectContaining({
          section: "sec_5_4_9_4",
          status: "drafted"
        })
      );

      const manifest = JSON.parse(await readFile(join(outputDir, "document-sections", "manifest.json"), "utf-8")) as {
        sections: Array<{ id: string; status: string; draftPath: string; tablesPath: string }>;
      };
      const section = manifest.sections.find((item) => item.id === "sec_5_4_9_4");
      expect(section?.status).toBe("edited");
      expect(section?.draftPath).toContain("document-sections/");
      expect(await readFile(join(outputDir, section!.draftPath), "utf-8")).toContain("统一身份认证系统重要数据存储场景");
      expect(await readFile(join(outputDir, section!.tablesPath), "utf-8")).toContain("table_30_5_4_9_4");

      const auditResult = await executeAgentToolCall(
        createToolCall("audit_document_sections", {
          section_ids: ["sec_5_4_9_4"],
          max_findings: 10
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );
      expect(auditResult.toolName).toBe("audit_document_sections");
      expect(auditResult.summary).toContain("证据风险");
      expect(auditResult.content).toContain("next_revise_document_sections_evidence_call");
      expect(auditResult.schemeProgressUpdates?.[0]).toEqual(
        expect.objectContaining({
          section: "sec_5_4_9_4",
          status: "drafted"
        })
      );

      const revisedResult = await executeAgentToolCall(
        createToolCall("revise_document_sections_evidence", {
          section_ids: ["sec_5_4_9_4"],
          max_rewrites: 10
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );
      expect(revisedResult.toolName).toBe("revise_document_sections_evidence");
      expect(revisedResult.summary).toContain("已修订");
      const revisedDraft = await readFile(join(outputDir, section!.draftPath), "utf-8");
      expect(revisedDraft).toContain("待补充/需确认：系统是否部署服务器密码机并是否配置签名验签能力。");
      expect(revisedDraft).not.toContain("系统已部署服务器密码机并已配置签名验签能力。");

      const polishResult = await executeAgentToolCall(
        createToolCall("polish_document_sections", {
          section_ids: ["sec_5_4_9_4"],
          style_rules: ["保留待补充/需确认标记", "减少空泛表述"],
          max_parallel: 1
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );
      expect(polishResult.toolName).toBe("polish_document_sections");
      expect(polishResult.summary).toContain("已润色 1");
      expect(polishResult.schemeProgressUpdates?.[0]).toEqual(
        expect.objectContaining({
          section: "sec_5_4_9_4",
          status: "drafted"
        })
      );
      const polishedDraft = await readFile(join(outputDir, section!.draftPath), "utf-8");
      expect(polishedDraft).toContain("待补充/需确认：系统是否部署服务器密码机并是否配置签名验签能力。");

      const detailResult = await executeAgentToolCall(
        createToolCall("get_document_section", {
          section: "sec_5_4_9_4"
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );
      expect(detailResult.toolName).toBe("get_document_section");
      expect(detailResult.content).toContain("section_context");
      expect(detailResult.content).toContain("draft.md");
      expect(detailResult.content).toContain("tables.json");

      const assembleResult = await executeAgentToolCall(
        createToolCall("assemble_document_sections", {
          section_ids: ["sec_5_4_9_4"],
          name: "章节终稿.md",
          include_draft: false
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );

      expect(assembleResult.toolName).toBe("assemble_document_sections");
      expect(assembleResult.content).toContain("next_word_call");
      expect(assembleResult.artifactPath).toBeTruthy();
      const finalMarkdown = await readFile(assembleResult.artifactPath!, "utf-8");
      expect(finalMarkdown).toContain("## 5.4.9.4 数据存储");
      expect(finalMarkdown).toContain("统一身份认证系统重要数据存储场景");
      expect(finalMarkdown).toContain("待补充/需确认：系统是否部署服务器密码机并是否配置签名验签能力。");
      expect(finalMarkdown).not.toContain("系统已部署服务器密码机并已配置签名验签能力。");
      expect(finalMarkdown).not.toContain("document-section-meta");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("audits final Markdown for unsupported definitive claims", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-audit-doc-"));
    const outputDir = join(dir, "output");
    const markdownPath = join(outputDir, "审查文档.md");

    try {
      await mkdir(outputDir, { recursive: true });
      await writeFile(
        join(outputDir, "document-config.json"),
        JSON.stringify(
          {
            version: 1,
            profile: "generic_document",
            title: "统一身份认证系统建设文档",
            generatedAt: "2026-06-08T00:00:00.000Z",
            sourceSummary: "",
            globalRules: [],
            facts: [{ key: "应用系统", value: "统一身份认证系统", source: "用户原话" }],
            gaps: ["等保级别", "密码产品型号"],
            sectionGroups: []
          },
          null,
          2
        ),
        "utf-8"
      );
      await writeFile(
        markdownPath,
        [
          "# 统一身份认证系统建设文档",
          "",
          "统一身份认证系统面向统一认证场景建设。",
          "",
          "系统已部署服务器密码机并已配置签名验签能力。",
          "",
          "等保级别为三级。"
        ].join("\n"),
        "utf-8"
      );

      const result = await executeAgentToolCall(
        createToolCall("audit_document_evidence", {
          markdown_path: "审查文档.md",
          max_findings: 10
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );

      expect(result.toolName).toBe("audit_document_evidence");
      expect(result.summary).toContain("疑似证据风险");
      expect(result.content).toContain("不要直接进入最终 Word 交付");
      expect(result.artifactPath).toBeTruthy();

      const report = await readFile(result.artifactPath!, "utf-8");
      expect(report).toContain("文档证据审查报告");
      expect(report).toContain("系统已部署服务器密码机");
      expect(report).toContain("待补充信息被写成确定表述：等保级别");
      expect(report).toContain("应用系统：统一身份认证系统");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("revises unsupported definitive claims into confirmation gaps", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-revise-doc-"));
    const outputDir = join(dir, "output");
    const markdownPath = join(outputDir, "待修订文档.md");

    try {
      await mkdir(outputDir, { recursive: true });
      await writeFile(
        join(outputDir, "document-config.json"),
        JSON.stringify(
          {
            version: 1,
            profile: "generic_document",
            title: "统一身份认证系统建设文档",
            generatedAt: "2026-06-08T00:00:00.000Z",
            sourceSummary: "",
            globalRules: [],
            facts: [{ key: "应用系统", value: "统一身份认证系统", source: "用户原话" }],
            gaps: ["等保级别"],
            sectionGroups: []
          },
          null,
          2
        ),
        "utf-8"
      );
      await writeFile(
        markdownPath,
        [
          "# 统一身份认证系统建设文档",
          "",
          "系统已部署服务器密码机并已配置签名验签能力。",
          "",
          "- 等保级别为三级。"
        ].join("\n"),
        "utf-8"
      );

      const result = await executeAgentToolCall(
        createToolCall("revise_document_evidence", {
          markdown_path: "待修订文档.md"
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );

      expect(result.toolName).toBe("revise_document_evidence");
      expect(result.summary).toContain("已修订 2 行");
      expect(result.content).toContain("next_audit_call");
      expect(result.content).toContain("next_word_call");
      expect(result.artifactPath).toBeTruthy();

      const revisedMarkdown = await readFile(result.artifactPath!, "utf-8");
      expect(revisedMarkdown).toContain("待补充/需确认：系统是否部署服务器密码机并是否配置签名验签能力。");
      expect(revisedMarkdown).toContain("- 待补充/需确认：等保级别为三级。");
      expect(revisedMarkdown).not.toContain("系统已部署服务器密码机并已配置签名验签能力。");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("writes assembled document Markdown into a Word file", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-doc-word-"));
    const outputDir = join(dir, "output");
    const markdownPath = join(outputDir, "建设文档终稿.md");

    try {
      await executeAgentToolCall(
        createToolCall("build_document_config", {
          profile: "generic_document",
          title: "统一身份认证系统建设文档",
          project_context: "系统名称：统一身份认证系统；建设单位：示例政务服务中心"
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );
      await mkdir(outputDir, { recursive: true });
      await writeFile(
        markdownPath,
        [
          "# 统一身份认证系统建设文档",
          "",
          "## 1 背景",
          "",
          "背景正文，说明统一身份认证系统的建设对象和合规背景。",
          "",
          "## 5 密码应用设计",
          "",
          "密码应用设计正文，说明建设思路和实施要点。"
        ].join("\n"),
        "utf-8"
      );

      const result = await executeAgentToolCall(
        createToolCall("write_document_word", {
          markdown_path: "建设文档终稿.md",
          name: "统一身份认证系统建设文档.docx"
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );
      const extracted = await mammoth.extractRawText({ path: result.artifactPath! });

      expect(result.toolName).toBe("write_document_word");
      expect(result.summary).toContain("已生成 Word");
      expect(result.content).toContain("写入模式：template_sections");
      expect(result.content).toContain("Word 模板：docs/templates/密码应用方案.docx");
      expect(result.artifactPath).toBeTruthy();
      expect(extracted.value).toContain("背景正文");
      expect(extracted.value).toContain("密码应用设计正文");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("auto-embeds generated diagram files when writing assembled document Word", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-doc-word-diagram-"));
    const outputDir = join(dir, "output");
    const markdownPath = join(outputDir, "网络环境终稿.md");
    const diagramPath = join(outputDir, "统一身份认证系统-网络架构图-mq5abcde-a1b2c3.png");

    try {
      await executeAgentToolCall(
        createToolCall("build_document_config", {
          profile: "generic_document",
          title: "统一身份认证系统建设文档",
          project_context: "系统名称：统一身份认证系统；建设单位：示例政务服务中心"
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );
      await mkdir(outputDir, { recursive: true });
      await writeFile(
        diagramPath,
        Buffer.from(
          "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
          "base64"
        )
      );
      await writeFile(
        markdownPath,
        [
          "# 统一身份认证系统建设文档",
          "",
          "## 2 系统概述",
          "",
          "### 2.2 计算平台现状",
          "",
          "#### 2.2.2 网络环境",
          "",
          "网络环境正文，说明业务用户通过互联网访问统一身份认证系统的通信信道。"
        ].join("\n"),
        "utf-8"
      );

      const result = await executeAgentToolCall(
        createToolCall("write_document_word", {
          markdown_path: "网络环境终稿.md",
          name: "统一身份认证系统建设文档.docx"
        }),
        {
          ...createContext(process.cwd()),
          outputDir
        }
      );
      const zip = new PizZip(await readFile(result.artifactPath!, "binary"));
      const documentXml = zip.file("word/document.xml")?.asText() ?? "";
      const markerIndex = documentXml.indexOf("ps:figure:fig_1_2_2_2_1:image");
      const sdtStart = documentXml.lastIndexOf("<w:sdt", markerIndex);
      const sdtEnd = documentXml.indexOf("</w:sdt>", markerIndex) + "</w:sdt>".length;
      const figureXml = documentXml.slice(sdtStart, sdtEnd);

      expect(result.toolName).toBe("write_document_word");
      expect(result.content).toContain("嵌入图示：网络架构图");
      expect(markerIndex).toBeGreaterThan(0);
      expect(figureXml).toContain("<w:drawing>");
      expect(figureXml).not.toContain("【图片占位】");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("does not support the legacy top-level template json path", async () => {
    await expect(
      executeAgentToolCall(
        createToolCall("read_file", { path: "docs/密码应用方案.template.json" }),
        createContext(process.cwd())
      )
    ).rejects.toThrow();
  });

  it("plans document table cells and figure tasks from the active Word template", async () => {
    const result = await executeAgentToolCall(
      createToolCall("plan_document_assets", {
        section_ids: ["sec_5_4_9_4"],
        max_items: 10
      }),
      createContext(process.cwd())
    );

    expect(result.toolName).toBe("plan_document_assets");
    expect(result.summary).toContain("表格 3 个、图示 2 个");
    expect(result.content).toContain("plan_document_assets completed");
    expect(result.content).toContain("template_tables_plan");
    expect(result.content).toContain("template_cells_plan");
    expect(result.content).toContain("\"table_id\": \"table_30_5_4_9_4\"");
    expect(result.content).toContain("\"figure_id\": \"fig_12_5_4_9_4\"");
    expect(result.content).toContain("next_plan_document_assets_call");
    expect(result.content).toContain("write_document_word_diagrams_plan");
    expect(result.content).toContain("write_document_word.template_cells");
  });

  it("paginates full template table cell planning to avoid truncating later tables", async () => {
    const result = await executeAgentToolCall(
      createToolCall("plan_document_assets", {
        include_figures: false,
        max_cells: 25,
        max_figures: 0
      }),
      createContext(process.cwd())
    );

    expect(result.toolName).toBe("plan_document_assets");
    expect(result.content).toContain("template_tables_plan");
    expect(result.content).toMatch(/可填单元格：\d+ 个；本批单元格：25 个/);
    expect(result.content).toContain("next_plan_document_assets_call");
    expect(result.content).toContain("\"cell_offset\": 25");
    expect(result.content).toContain("\"include_tables\": true");
    expect(result.content).toContain("\"include_figures\": false");
    expect(result.content).not.toContain("内容过长，已截断");
  });

  it("rejects asset planning for unknown sections", async () => {
    const result = await executeAgentToolCall(
      createToolCall("plan_document_assets", {
        section_ids: ["sec_7_2"]
      }),
      createContext(process.cwd())
    );

    expect(result.toolName).toBe("plan_document_assets");
    expect(result.summary).toContain("章节不在模板中");
    expect(result.content).toContain("unknown section_ids sec_7_2");
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

  it("sends generated files with a trailing stamp when the model provides the original name", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-send-stamped-"));
    const outputDir = join(dir, "output");
    const outputPath = join(outputDir, "test_file-mpl38ooc-a1b2c3.txt");

    try {
      await mkdir(outputDir, { recursive: true });
      await writeFile(outputPath, "ok", "utf-8");
      const result = await executeAgentToolCall(createToolCall("send_file", { path: "data/output/test_file.txt" }), {
        ...createContext(dir),
        outputDir
      });

      expect(result.toolName).toBe("send_file");
      expect(result.artifactPath).toBe(outputPath);
      expect(result.summary).toContain("test_file-mpl38ooc-a1b2c3.txt");
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

  it("returns an explicit configuration error for image recognition without guessing content", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-tool-vision-"));
    const imagePath = join(dir, "截图.png");

    try {
      await writeFile(
        imagePath,
        Buffer.from(
          "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
          "base64"
        )
      );

      const result = await executeAgentToolCall(
        createToolCall("read_image", { path: imagePath, question: "识别截图里需要修改的问题" }),
        createContext(dir)
      );

      expect(result.toolName).toBe("read_image");
      expect(result.summary).toContain("OPENAI_VISION_API_KEY");
      expect(result.summary).toContain("OPENAI_API_KEY");
      expect(result.content).toContain("截图.png");
      expect(result.content).toContain("未配置");
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
      visionBaseUrl: "",
      chatModel: "gpt-5.5",
      chatImageInputEnabled: false,
      imageModel: "gpt-image-2",
      visionModel: "",
      imageSize: "1536x1024",
      imageQuality: "high",
      autoImageGeneration: true,
      thinkingEnabled: true,
      reasoningEffort: "",
      requestTimeoutMs: 120000,
      imageRequestTimeoutMs: 300000,
      maxOutputTokens: 16000,
      apiKeyConfigured: false,
      imageApiKeyConfigured: false,
      visionApiKeyConfigured: false
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
