import { describe, expect, it } from "vitest";
import { resolvePiAgentExport } from "../../src/main/piAgentAdapter";
import type { AgentRuntime } from "../../src/main/agentRuntime";

describe("piAgentAdapter", () => {
  it("resolves explicit and conventional Pi Agent exports", () => {
    const runtime = createRuntime();
    const factory = () => runtime;
    const module = {
      default: undefined,
      createRuntime: factory,
      namedRuntime: runtime
    };

    expect(resolvePiAgentExport(module, "namedRuntime")).toBe(runtime);
    expect(resolvePiAgentExport(module, "")).toBe(factory);
  });

  it("accepts runtime objects as plugin modules", () => {
    const runtime = createRuntime();

    expect(resolvePiAgentExport(runtime, "")).toBe(runtime);
  });
});

function createRuntime(): AgentRuntime {
  return {
    kind: "pi-agent",
    runTurn: async () => ({
      id: "msg_1",
      kind: "message",
      role: "assistant",
      content: "ok",
      isFinished: true,
      createdAt: "2026-05-23T00:00:00.000Z"
    })
  };
}
