import { join } from "node:path";
import { resolveBundledResourceDir } from "./bundledRuntime";

export interface AppPathOptions {
  projectRootDir: string;
  resourcesDir?: string;
  userDataDir: string;
  packaged: boolean;
}

export interface AppPaths {
  rootDir: string;
  projectRootDir: string;
  docsDir: string;
  dataDir: string;
  inputDir: string;
  outputDir: string;
  statePath: string;
  envLocalPath: string;
  envPath: string;
}

export function resolveAppPaths(options: AppPathOptions): AppPaths {
  const rootDir = options.packaged ? options.userDataDir : options.projectRootDir;
  const dataDir = join(rootDir, "data");

  return {
    rootDir,
    projectRootDir: options.projectRootDir,
    docsDir: resolveBundledResourceDir(options.projectRootDir, "docs", options.resourcesDir),
    dataDir,
    inputDir: join(dataDir, "input"),
    outputDir: join(dataDir, "output"),
    statePath: join(dataDir, "state.json"),
    envLocalPath: join(rootDir, ".env.local"),
    envPath: join(options.projectRootDir, ".env")
  };
}
