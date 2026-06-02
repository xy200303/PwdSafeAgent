import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { buildSchemeStandardReferenceContext } from "../../src/main/schemeReference";

const BUILT_IN_DOCS_DIR = join(process.cwd(), "resources", "docs");

describe("schemeReference", () => {
  it("selects third-level application and data requirements for relevant sections", () => {
    const context = buildSchemeStandardReferenceContext(BUILT_IN_DOCS_DIR, {
      sectionNumber: "5.4.9.4",
      sectionTitle: "重要数据存储保护",
      paragraphTasks: ["说明身份鉴别、重要数据存储机密性与完整性控制措施"],
      projectContext: "等保级别：三级"
    });

    expect(context).toContain("docs/references/密码应用国家标准.md");
    expect(context).toContain("8.4 应用和数据安全");
    expect(context).toContain("应采用密码技术对登录用户进行身份鉴别");
    expect(context).toContain("重要数据在存储过程中的机密性");
  });

  it("includes key management guidance when the section focuses on key lifecycle", () => {
    const context = buildSchemeStandardReferenceContext(BUILT_IN_DOCS_DIR, {
      sectionNumber: "5.3.7",
      sectionTitle: "密钥管理方式",
      paragraphTasks: ["说明密钥产生、分发、存储、使用、更新和销毁要求"],
      projectContext: "等保级别：三级"
    });

    expect(context).toContain("密钥管理 key management");
    expect(context).toContain("密钥的产生、分发、存储、使用、更新");
    expect(context).toContain("8.7 建设运行");
  });

  it("falls back to framework and general requirements for principle sections", () => {
    const context = buildSchemeStandardReferenceContext(BUILT_IN_DOCS_DIR, {
      sectionNumber: "4",
      sectionTitle: "安全目标及设计原则",
      paragraphTasks: ["说明密码应用技术框架、通用要求和方案设计依据"]
    });

    expect(context).toContain("4.1 信息系统密码应用技术框架");
    expect(context).toContain("5 通用要求");
    expect(context).toContain("信息系统中使用的密码算法应符合法律、法规的规定");
  });
});
