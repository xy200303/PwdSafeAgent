import { describe, expect, it } from "vitest";
import {
  createAgentRuntime,
  normalizeAgentRuntimeKind,
  setOpenAIClientFactoryForTesting,
  type AgentRuntimeHost,
  type MessageStreamItem
} from "../../src/main/agentRuntime";
import type { AppSettings, ChatSession, StreamItem } from "../../src/shared/types";

describe("agentRuntime", () => {
  it("normalizes runtime names with openai-chat as the safe default", () => {
    expect(normalizeAgentRuntimeKind(undefined)).toBe("openai-chat");
    expect(normalizeAgentRuntimeKind("")).toBe("openai-chat");
    expect(normalizeAgentRuntimeKind("pi-agent")).toBe("pi-agent");
    expect(normalizeAgentRuntimeKind("unknown")).toBe("openai-chat");
  });

  it("routes pi-agent through the adapter layer and falls back to mock OpenAI runtime without an API key", async () => {
    const originalApiKey = process.env.OPENAI_API_KEY;
    process.env.OPENAI_API_KEY = "";
    const items: StreamItem[] = [];
    const host = createHost(items);

    try {
      const runtime = createAgentRuntime("pi-agent", host);
      const assistant = await runtime.runTurn({
        sessionId: "session_1",
        userPrompt: "你好",
        controller: new AbortController(),
        onAssistantCreated: () => {}
      });

      expect(assistant.content).toContain("mock response");
      expect(items.some((item) => item.kind === "tool" && item.toolName === "pi-agent.runtime")).toBe(true);
      expect(items.some((item) => item.kind === "tool" && item.summary?.includes("未配置 PI_AGENT_PACKAGE"))).toBe(true);
      expect(items.some((item) => item.kind === "message" && item.role === "assistant")).toBe(true);
    } finally {
      if (originalApiKey === undefined) {
        delete process.env.OPENAI_API_KEY;
      } else {
        process.env.OPENAI_API_KEY = originalApiKey;
      }
    }
  });

  it("records redacted tool input and output previews", async () => {
    const originalApiKey = process.env.OPENAI_API_KEY;
    process.env.OPENAI_API_KEY = "sk-test-runtime";
    setOpenAIClientFactoryForTesting(() => createMockOpenAIClient([
      {
        choices: [
          {
            message: {
              role: "assistant",
              content: null,
              tool_calls: [
                {
                  id: "call_1",
                  type: "function",
                  function: {
                    name: "time",
                    arguments: "{\"OPENAI_API_KEY\":\"sk-secret123456789\",\"note\":\"hello\"}"
                  }
                }
              ]
            }
          }
        ]
      },
      {
        choices: [
          {
            message: {
              role: "assistant",
              content: "工具已完成",
              tool_calls: []
            }
          }
        ]
      },
      {
        async *[Symbol.asyncIterator]() {
          yield { choices: [{ delta: { content: "完成" } }] };
        }
      }
    ]));
    const items: StreamItem[] = [];

    try {
      const runtime = createAgentRuntime("openai-chat", createHost(items));
      await runtime.runTurn({
        sessionId: "session_1",
        userPrompt: "现在几点",
        controller: new AbortController(),
        onAssistantCreated: () => {}
      });

      const toolItem = items.find((item) => item.kind === "tool" && item.toolName === "time");
      expect(toolItem?.kind).toBe("tool");
      if (toolItem?.kind !== "tool") return;
      expect(toolItem.inputPreview).toContain("[REDACTED]");
      expect(toolItem.inputPreview).not.toContain("sk-secret123456789");
      expect(toolItem.outputPreview).toMatch(/\d{4}年/);
    } finally {
      setOpenAIClientFactoryForTesting();
      if (originalApiKey === undefined) {
        delete process.env.OPENAI_API_KEY;
      } else {
        process.env.OPENAI_API_KEY = originalApiKey;
      }
    }
  });
});

function createMockOpenAIClient(responses: unknown[]) {
  let index = 0;
  return {
    chat: {
      completions: {
        create: async () => responses[index++]
      }
    }
  };
}

function createHost(items: StreamItem[]): AgentRuntimeHost {
  const session: ChatSession = {
    id: "session_1",
    title: "测试会话",
    status: "running",
    createdAt: "2026-05-23T00:00:00.000Z",
    updatedAt: "2026-05-23T00:00:00.000Z",
    items
  };

  return {
    rootDir: process.cwd(),
    docsDir: process.cwd(),
    inputDir: process.cwd(),
    outputDir: process.cwd(),
    loadSettings: createSettings,
    getSession: () => session,
    buildMessages: () => [],
    createAssistantMessage: () => {
      const message: MessageStreamItem = {
        id: "msg_1",
        kind: "message",
        role: "assistant",
        content: "",
        isFinished: false,
        createdAt: "2026-05-23T00:00:01.000Z"
      };
      items.push(message);
      return message;
    },
    updateItem: (_sessionId, item) => {
      const index = items.findIndex((existing) => existing.id === item.id);
      if (index >= 0) items[index] = item;
    },
    startToolCall: (_sessionId, toolName, summary) => {
      const tool: StreamItem = {
        id: `tool_${items.length}`,
        kind: "tool",
        toolCallId: `toolcall_${items.length}`,
        toolName,
        status: "running",
        summary,
        createdAt: "2026-05-23T00:00:01.000Z"
      };
      items.push(tool);
      return tool;
    },
    finishToolCall: (_sessionId, item, status, summary) => {
      if (item.kind !== "tool") return;
      item.status = status;
      item.summary = summary;
    },
    addStage: (sessionId, title, detail) => {
      items.push({
        id: `stage_${items.length}`,
        kind: "stage",
        title: `${sessionId}:${title}`,
        detail,
        createdAt: "2026-05-23T00:00:01.000Z"
      });
    },
    appendSessionMemory: () => {},
    formatSessionMemory: () => "",
    getSessionReadableFiles: () => [],
    getBundledPythonRuntime: () => undefined,
    createArtifact: () => {},
    streamMockResponse: async (_sessionId, assistantItem) => {
      assistantItem.content = "mock response";
    }
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
      chatModel: "gpt-5.5",
      imageModel: "gpt-image-2",
      imageSize: "1536x1024",
      imageQuality: "high",
      autoImageGeneration: true,
      requestTimeoutMs: 120000,
      maxOutputTokens: 16000,
      apiKeyConfigured: false,
      imageApiKeyConfigured: false
    },
    document: {
      autoPdfExport: false,
      libreOfficePath: ""
    },
    agent: {
      runtime: "pi-agent",
      execBashEnabled: false,
      piAgentPackage: "",
      piAgentExport: ""
    }
  };
}
