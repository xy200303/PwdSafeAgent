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

  it("validates draft_scheme_sections arguments through the Pi TypeBox path", () => {
    const tool = buildAgentChatTools({ includeExecBash: false, includeArtifactTools: true }).find(
      (item) => item.type === "function" && item.function.name === "draft_scheme_sections"
    );
    if (tool?.type !== "function") throw new Error("draft_scheme_sections tool not found");

    const schema = convertJsonSchemaToTypeBoxSchema(tool.function.parameters);
    const validator = Compile(schema);
    const args = {
      sections: [
        { section: "sec_1", title: "背景", writing_hint: "提示", paragraph_tasks: ["说明保护对象和建设目标"] },
        { section: "sec_1_1", paragraph_tasks: ["承接风险结论说明本节需求来源"] },
        { section: "sec_1_2", paragraph_tasks: ["说明管理、运维或审计配套要求"] },
        { section: "sec_1_2_1", paragraph_tasks: ["归纳需通过密码技术或管理措施控制的重点"] }
      ],
      max_parallel: 4
    };

    const convertedArgs = structuredClone(args);
    Value.Convert(schema, convertedArgs);

    expect(validator.Check(convertedArgs)).toBe(true);
    expect([...validator.Errors(convertedArgs)]).toEqual([]);
    expect(validator.Check({ sections: [{ section: "sec_1", writingHint: "不要传 camelCase 字段" }] })).toBe(false);
  });

  it("includes the enabled Pi tool set in the schema signature", () => {
    const safeSettings = createSettings();
    const bashSettings = createSettings();
    bashSettings.agent.execBashEnabled = true;

    const safeSignature = buildPwdSafePiToolSchemaSignature(safeSettings);
    const bashSignature = buildPwdSafePiToolSchemaSignature(bashSettings);

    expect(safeSignature).toContain("draft_scheme_sections");
    expect(safeSignature).not.toContain("exec_bash");
    expect(bashSignature).toContain("exec_bash");
    expect(bashSignature).not.toBe(safeSignature);
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
