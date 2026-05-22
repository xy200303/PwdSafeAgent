import { describe, expect, it } from "vitest";
import {
  hasDiagramArtifactIntent,
  hasSchemeArtifactIntent,
  shouldCreateDraftArtifact
} from "../../src/main/artifactIntent";

describe("artifactIntent", () => {
  it("does not treat greetings as document delivery requests", () => {
    expect(hasSchemeArtifactIntent("你好")).toBe(false);
    expect(hasDiagramArtifactIntent("你好")).toBe(false);
    expect(
      shouldCreateDraftArtifact(
        "你好",
        "我会按密码应用方案模板推进，后续可生成技术架构、流程、密钥管理和实施计划等内容。".repeat(3)
      )
    ).toBe(false);
  });

  it("detects explicit scheme delivery requests", () => {
    expect(hasSchemeArtifactIntent("请生成密码应用方案")).toBe(true);
    expect(hasSchemeArtifactIntent("按照模板编制一份商用密码应用方案文档")).toBe(true);
    expect(
      shouldCreateDraftArtifact(
        "请生成密码应用方案",
        "本方案包含系统概况、密码应用需求、技术架构、密钥管理、签名验签、加密存储和日志审计。".repeat(3)
      )
    ).toBe(true);
  });

  it("detects explicit diagram requests", () => {
    expect(hasDiagramArtifactIntent("请画一张密码应用技术架构图")).toBe(true);
    expect(hasDiagramArtifactIntent("输出业务流程图")).toBe(true);
  });
});
