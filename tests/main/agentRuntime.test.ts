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
      expect(items.some((item) => item.kind === "tool" && item.toolName === "openai.chat.tools")).toBe(false);
    } finally {
      setOpenAIClientFactoryForTesting();
      if (originalApiKey === undefined) {
        delete process.env.OPENAI_API_KEY;
      } else {
        process.env.OPENAI_API_KEY = originalApiKey;
      }
    }
  });

  it("normalizes invalid tool arguments before appending tool calls to chat history", async () => {
    const originalApiKey = process.env.OPENAI_API_KEY;
    process.env.OPENAI_API_KEY = "sk-test-runtime";
    const requestMessages: unknown[] = [];
    setOpenAIClientFactoryForTesting(() => createMockOpenAIClient(
      [
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
                      name: "read_word",
                      arguments: "docs/密码应用方案.docx"
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
      ],
      requestMessages
    ));

    try {
      const runtime = createAgentRuntime("openai-chat", createHost([]));
      await runtime.runTurn({
        sessionId: "session_1",
        userPrompt: "读取模板",
        controller: new AbortController(),
        onAssistantCreated: () => {}
      });

      const secondRequest = requestMessages[1] as { messages?: Array<Record<string, unknown>> };
      const assistantToolMessage = secondRequest.messages?.find((message) => message.role === "assistant" && message.tool_calls);
      const toolCalls = assistantToolMessage?.tool_calls as Array<{ function: { arguments: string } }> | undefined;
      expect(toolCalls?.[0]?.function.arguments).toBe(JSON.stringify({ path: "docs/密码应用方案.docx" }));
    } finally {
      setOpenAIClientFactoryForTesting();
      if (originalApiKey === undefined) {
        delete process.env.OPENAI_API_KEY;
      } else {
        process.env.OPENAI_API_KEY = originalApiKey;
      }
    }
  });

  it("hides final artifact tools while the user is still providing project information", async () => {
    const originalApiKey = process.env.OPENAI_API_KEY;
    process.env.OPENAI_API_KEY = "sk-test-runtime";
    const requestMessages: unknown[] = [];
    setOpenAIClientFactoryForTesting(() => createMockOpenAIClient(
      [
        {
          choices: [
            {
              message: {
                role: "assistant",
                content: "已记录项目基础信息，继续补充建设单位和等保级别即可。",
                tool_calls: []
              }
            }
          ]
        },
        {
          async *[Symbol.asyncIterator]() {
            yield { choices: [{ delta: { content: "已记录" } }] };
          }
        }
      ],
      requestMessages
    ));

    try {
      const runtime = createAgentRuntime("openai-chat", createHost([]));
      await runtime.runTurn({
        sessionId: "session_1",
        userPrompt: "系统名称是统一身份认证系统",
        controller: new AbortController(),
        onAssistantCreated: () => {}
      });

      const planningRequest = requestMessages[0] as { tools?: Array<{ function: { name: string } }> };
      const toolNames = planningRequest.tools?.map((tool) => tool.function.name) ?? [];
      expect(toolNames).toContain("remember_project");
      expect(toolNames).toContain("read_word");
      expect(toolNames).not.toContain("write_word");
      expect(toolNames).not.toContain("image_generate");
      expect(toolNames).not.toContain("send_file");
    } finally {
      setOpenAIClientFactoryForTesting();
      if (originalApiKey === undefined) {
        delete process.env.OPENAI_API_KEY;
      } else {
        process.env.OPENAI_API_KEY = originalApiKey;
      }
    }
  });

  it("exposes final artifact tools when the user explicitly asks to generate a scheme", async () => {
    const originalApiKey = process.env.OPENAI_API_KEY;
    process.env.OPENAI_API_KEY = "sk-test-runtime";
    const requestMessages: unknown[] = [];
    setOpenAIClientFactoryForTesting(() => createMockOpenAIClient(
      [
        {
          choices: [
            {
              message: {
                role: "assistant",
                content: "准备生成方案。",
                tool_calls: []
              }
            }
          ]
        },
        {
          async *[Symbol.asyncIterator]() {
            yield { choices: [{ delta: { content: "准备生成" } }] };
          }
        }
      ],
      requestMessages
    ));

    try {
      const runtime = createAgentRuntime("openai-chat", createHost([]));
      await runtime.runTurn({
        sessionId: "session_1",
        userPrompt: "请生成密码应用方案 Word",
        controller: new AbortController(),
        onAssistantCreated: () => {}
      });

      const planningRequest = requestMessages[0] as { tools?: Array<{ function: { name: string } }> };
      const toolNames = planningRequest.tools?.map((tool) => tool.function.name) ?? [];
      expect(toolNames).toContain("write_word");
      expect(toolNames).toContain("image_generate");
      expect(toolNames).toContain("send_file");
    } finally {
      setOpenAIClientFactoryForTesting();
      if (originalApiKey === undefined) {
        delete process.env.OPENAI_API_KEY;
      } else {
        process.env.OPENAI_API_KEY = originalApiKey;
      }
    }
  });

  it("falls back to plain message history when final streaming rejects tool call history", async () => {
    const originalApiKey = process.env.OPENAI_API_KEY;
    process.env.OPENAI_API_KEY = "sk-test-runtime";
    const requestMessages: unknown[] = [];
    const items: StreamItem[] = [];
    setOpenAIClientFactoryForTesting(() => createMockOpenAIClient(
      [
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
                      arguments: "{}"
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
        new Error("400 InternalError.Algo.InvalidParameter: The \"function.arguments\" parameter must be in JSON format."),
        {
          async *[Symbol.asyncIterator]() {
            yield { choices: [{ delta: { content: "降级完成" } }] };
          }
        }
      ],
      requestMessages
    ));

    try {
      const runtime = createAgentRuntime("openai-chat", createHost(items));
      const assistant = await runtime.runTurn({
        sessionId: "session_1",
        userPrompt: "现在几点",
        controller: new AbortController(),
        onAssistantCreated: () => {}
      });

      expect(assistant.content).toBe("降级完成");
      expect(items.some((item) => item.kind === "stage" && item.title.includes("最终回复已兼容降级"))).toBe(true);

      const fallbackRequest = requestMessages[3] as { messages?: Array<Record<string, unknown>> };
      expect(fallbackRequest.messages?.some((message) => message.role === "tool")).toBe(false);
      expect(fallbackRequest.messages?.some((message) => "tool_calls" in message)).toBe(false);
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

function createMockOpenAIClient(responses: unknown[], requestBodies: unknown[] = []) {
  let index = 0;
  return {
    chat: {
      completions: {
        create: async (body: unknown) => {
          requestBodies.push(body);
          const response = responses[index++];
          if (response instanceof Error) throw response;
          return response;
        }
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
