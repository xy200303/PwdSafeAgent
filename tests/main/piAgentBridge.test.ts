import { Compile } from "typebox/compile";
import { Value } from "typebox/value";
import { describe, expect, it } from "vitest";
import { buildAgentChatTools } from "../../src/main/agentToolRegistry";
import { buildPwdSafePiToolSchemaSignature, convertJsonSchemaToTypeBoxSchema } from "../../src/main/piAgentBridge";
import type { AppSettings } from "../../src/shared/types";

describe("piAgentBridge tool schemas", () => {
  it("converts every OpenAI tool schema into a TypeBox-compilable schema", () => {
    const tools = buildAgentChatTools({ includeExecBash: true, includeArtifactTools: true });

    for (const tool of tools) {
      if (tool.type !== "function") continue;

      expect(() => Compile(convertJsonSchemaToTypeBoxSchema(tool.function.parameters))).not.toThrow();
    }
  });

  it("validates draft_document_sections arguments through the Pi TypeBox path", () => {
    const tool = buildAgentChatTools({ includeExecBash: false, includeArtifactTools: true }).find(
      (item) => item.type === "function" && item.function.name === "draft_document_sections"
    );
    if (tool?.type !== "function") throw new Error("draft_document_sections tool not found");

    const schema = convertJsonSchemaToTypeBoxSchema(tool.function.parameters);
    const validator = Compile(schema);
    const args = {
      section_ids: ["sec_1", "sec_5_4_9_4"],
      project_context: "系统名称：统一身份认证系统",
      max_parallel: 4
    };

    const convertedArgs = structuredClone(args);
    Value.Convert(schema, convertedArgs);

    expect(validator.Check(convertedArgs)).toBe(true);
    expect([...validator.Errors(convertedArgs)]).toEqual([]);
    expect(validator.Check({ sectionIds: ["不要传 camelCase 字段"] })).toBe(false);
  });

  it("includes the enabled Pi tool set in the schema signature", () => {
    const safeSettings = createSettings();
    const bashSettings = createSettings();
    bashSettings.agent.execBashEnabled = true;
    const visionChatSettings = createSettings();
    visionChatSettings.openai.chatImageInputEnabled = true;

    const safeSignature = buildPwdSafePiToolSchemaSignature(safeSettings);
    const bashSignature = buildPwdSafePiToolSchemaSignature(bashSettings);
    const visionChatSignature = buildPwdSafePiToolSchemaSignature(visionChatSettings);

    expect(safeSignature).not.toContain("draft_scheme_sections");
    expect(safeSignature).not.toContain("register_document_template");
    expect(safeSignature).toContain("build_document_config");
    expect(safeSignature).toContain("list_document_sections");
    expect(safeSignature).toContain("get_document_section");
    expect(safeSignature).toContain("draft_document_sections");
    expect(safeSignature).toContain("update_document_section_draft");
    expect(safeSignature).toContain("audit_document_sections");
    expect(safeSignature).toContain("revise_document_sections_evidence");
    expect(safeSignature).toContain("polish_document_sections");
    expect(safeSignature).toContain("assemble_document_sections");
    expect(safeSignature).not.toContain("draft_document_modules");
    expect(safeSignature).not.toContain("assemble_document_content");
    expect(safeSignature).toContain("audit_document_evidence");
    expect(safeSignature).toContain("revise_document_evidence");
    expect(safeSignature).toContain("write_document_word");
    expect(safeSignature).not.toContain("assemble_scheme_markdown");
    expect(safeSignature).not.toContain("plan_scheme_batches");
    expect(safeSignature).not.toContain("plan_scheme_assets");
    expect(safeSignature).not.toContain("create_word");
    expect(safeSignature).not.toContain("write_word");
    expect(safeSignature).not.toContain("exec_bash");
    expect(bashSignature).toContain("exec_bash");
    expect(bashSignature).not.toBe(safeSignature);
    expect(visionChatSignature).toBe(safeSignature);
  });
});

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
