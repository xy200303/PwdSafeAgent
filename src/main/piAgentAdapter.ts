import type { AppSettings } from "../shared/types";
import type { AgentRuntime, AgentRuntimeHost } from "./agentRuntime";

export interface PiAgentPluginContext {
  settings: AppSettings;
  host: AgentRuntimeHost;
  fallbackRuntime: AgentRuntime;
}

export type PiAgentRuntimeFactory = (context: PiAgentPluginContext) => AgentRuntime | Promise<AgentRuntime>;

export type PiAgentPluginModule =
  | PiAgentRuntimeFactory
  | AgentRuntime
  | {
      default?: PiAgentRuntimeFactory | AgentRuntime;
      createRuntime?: PiAgentRuntimeFactory;
      createAgentRuntime?: PiAgentRuntimeFactory;
      createPiAgentRuntime?: PiAgentRuntimeFactory;
      runtime?: AgentRuntime;
    };

export interface PiAgentAdapterLoadResult {
  status: "loaded" | "not-configured" | "failed";
  runtime?: AgentRuntime;
  summary: string;
}

export async function loadPiAgentRuntime(
  settings: AppSettings,
  host: AgentRuntimeHost,
  fallbackRuntime: AgentRuntime
): Promise<PiAgentAdapterLoadResult> {
  const packageName = settings.agent.piAgentPackage.trim();
  if (!packageName) {
    return {
      status: "not-configured",
      summary: "未配置 PI_AGENT_PACKAGE，继续使用 OpenAI Chat Runtime 回退"
    };
  }

  try {
    const pluginModule = (await import(packageName)) as PiAgentPluginModule;
    const candidate = resolvePiAgentExport(pluginModule, settings.agent.piAgentExport);
    const runtime = await resolvePiAgentRuntime(candidate, {
      settings,
      host,
      fallbackRuntime
    });

    if (!isAgentRuntime(runtime)) {
      return {
        status: "failed",
        summary: `Pi Agent 插件 ${packageName} 未返回有效 runtime`
      };
    }

    return {
      status: "loaded",
      runtime,
      summary: `已加载 Pi Agent 插件 ${packageName}`
    };
  } catch (error) {
    return {
      status: "failed",
      summary: `Pi Agent 插件加载失败：${error instanceof Error ? error.message : String(error)}`
    };
  }
}

export function resolvePiAgentExport(
  pluginModule: PiAgentPluginModule,
  exportName: string
): PiAgentRuntimeFactory | AgentRuntime | undefined {
  if (typeof pluginModule === "function" || isAgentRuntime(pluginModule)) return pluginModule;
  const normalizedExportName = exportName.trim();
  if (normalizedExportName) {
    return (pluginModule as Record<string, unknown>)[normalizedExportName] as PiAgentRuntimeFactory | AgentRuntime | undefined;
  }
  return (
    pluginModule.default ??
    pluginModule.createRuntime ??
    pluginModule.createAgentRuntime ??
    pluginModule.createPiAgentRuntime ??
    pluginModule.runtime
  );
}

async function resolvePiAgentRuntime(
  candidate: PiAgentRuntimeFactory | AgentRuntime | undefined,
  context: PiAgentPluginContext
): Promise<AgentRuntime | undefined> {
  if (!candidate) return undefined;
  if (isAgentRuntime(candidate)) return candidate;
  if (typeof candidate === "function") return candidate(context);
  return undefined;
}

function isAgentRuntime(value: unknown): value is AgentRuntime {
  return Boolean(
    value &&
      typeof value === "object" &&
      typeof (value as AgentRuntime).kind === "string" &&
      typeof (value as AgentRuntime).runTurn === "function"
  );
}
