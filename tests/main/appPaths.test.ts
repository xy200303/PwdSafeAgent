import { mkdir, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { resolveAppPaths } from "../../src/main/appPaths";

describe("appPaths", () => {
  it("keeps development data beside the project root", () => {
    const paths = resolveAppPaths({
      projectRootDir: "C:\\Project\\PwdSafeAgent",
      userDataDir: "C:\\Users\\me\\AppData\\Roaming\\PwdSafeAgent",
      packaged: false
    });

    expect(paths.rootDir).toBe("C:\\Project\\PwdSafeAgent");
    expect(paths.dataDir).toBe("C:\\Project\\PwdSafeAgent\\data");
    expect(paths.envLocalPath).toBe("C:\\Project\\PwdSafeAgent\\.env.local");
    expect(paths.envPath).toBe("C:\\Project\\PwdSafeAgent\\.env");
  });

  it("moves packaged writable data into Electron userData", () => {
    const paths = resolveAppPaths({
      projectRootDir: "C:\\Program Files\\PwdSafeAgent",
      userDataDir: "C:\\Users\\me\\AppData\\Roaming\\PwdSafeAgent",
      packaged: true
    });

    expect(paths.rootDir).toBe("C:\\Users\\me\\AppData\\Roaming\\PwdSafeAgent");
    expect(paths.dataDir).toBe("C:\\Users\\me\\AppData\\Roaming\\PwdSafeAgent\\data");
    expect(paths.outputDir).toBe("C:\\Users\\me\\AppData\\Roaming\\PwdSafeAgent\\data\\output");
    expect(paths.envLocalPath).toBe("C:\\Users\\me\\AppData\\Roaming\\PwdSafeAgent\\.env.local");
    expect(paths.envPath).toBe("C:\\Program Files\\PwdSafeAgent\\.env");
  });

  it("resolves packaged docs from resources while keeping writable data in userData", async () => {
    const projectRootDir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-paths-project-"));
    const resourcesDir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-paths-resources-"));
    const userDataDir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-paths-user-"));
    const docsDir = join(resourcesDir, "docs");

    try {
      await mkdir(docsDir, { recursive: true });

      const paths = resolveAppPaths({
        projectRootDir,
        resourcesDir,
        userDataDir,
        packaged: true
      });

      expect(paths.docsDir).toBe(docsDir);
      expect(paths.dataDir).toBe(join(userDataDir, "data"));
    } finally {
      await rm(projectRootDir, { recursive: true, force: true });
      await rm(resourcesDir, { recursive: true, force: true });
      await rm(userDataDir, { recursive: true, force: true });
    }
  });
});
