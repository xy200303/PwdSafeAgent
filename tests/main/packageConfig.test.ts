import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";

interface PackageBuildConfig {
  build?: {
    extraResources?: Array<{
      from?: string;
      to?: string;
      filter?: string[];
    }>;
  };
}

describe("package config", () => {
  it("packages the built-in Word scheme template as an extra resource", async () => {
    const packageJson = JSON.parse(await readFile(join(process.cwd(), "package.json"), "utf-8")) as PackageBuildConfig;
    const docsResource = packageJson.build?.extraResources?.find((resource) => resource.from === "docs" && resource.to === "docs");

    expect(docsResource).toBeTruthy();
    expect(docsResource?.filter ?? []).toContain("**/*");
  });
});
