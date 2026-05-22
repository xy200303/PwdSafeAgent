import { exec } from "node:child_process";
import { promisify } from "node:util";
import { basename, extname, isAbsolute, join, resolve } from "node:path";
import type { ChatCompletionMessageToolCall, ChatCompletionTool } from "openai/resources/chat/completions";
import type { AppSettings } from "../shared/types";
import {
  compactText,
  getCurrentTimeText,
  readDocumentText,
  sanitizeFileName,
  writeUtf8File,
  type BuiltinToolName
} from "./agentTools";
import { exportDocxToPdf } from "./documentExport";
import { generateDiagramImage, type DiagramKind } from "./imageGeneration";
import { writeSchemeDocxFromTemplate } from "./schemeDocument";

const execAsync = promisify(exec);

export interface AgentToolExecutionContext {
  rootDir: string;
  outputDir: string;
  sessionTitle: string;
  memory: string;
  settings: AppSettings;
  signal?: AbortSignal;
  allowedReadDirs: string[];
  execBashEnabled: boolean;
}

export interface AgentToolExecutionResult {
  toolName: BuiltinToolName;
  summary: string;
  content: string;
  artifactPath?: string;
}

export interface WebSearchResult {
  title: string;
  url: string;
  snippet: string;
}

export function buildAgentChatTools(options: { includeExecBash: boolean }): ChatCompletionTool[] {
  const tools: ChatCompletionTool[] = [
    {
      type: "function",
      function: {
        name: "time",
        description: "获取当前北京时间，用于方案编制日期、时间敏感资料判断。",
        parameters: {
          type: "object",
          properties: {},
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "web_search",
        description: "检索互联网公开信息。只在需要查询最新政策、标准、产品背景或用户明确要求联网时使用。",
        parameters: {
          type: "object",
          properties: {
            query: {
              type: "string",
              description: "搜索关键词，建议包含中文标准号、政策名称、产品名称或机构名称。"
            },
            max_results: {
              type: "integer",
              description: "返回结果数量，默认 5，最大 8。"
            }
          },
          required: ["query"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "read_file",
        description: "读取 docs、data/input、data/output 范围内的文本文件内容；Word/PDF 优先使用 read_word/read_pdf。",
        parameters: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "文件路径。可以是绝对路径，也可以是相对项目根目录的路径。"
            }
          },
          required: ["path"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "read_word",
        description: "读取 docs、data/input、data/output 范围内的 .docx Word 文件，提取正文文本用于方案生成。",
        parameters: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "Word 文件路径。可以是绝对路径，也可以是相对项目根目录的路径。"
            }
          },
          required: ["path"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "read_pdf",
        description: "读取 docs、data/input、data/output 范围内的 .pdf 文件，提取正文文本用于方案生成。",
        parameters: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "PDF 文件路径。可以是绝对路径，也可以是相对项目根目录的路径。"
            }
          },
          required: ["path"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "write_file",
        description: "把中间分析、清单或方案片段写入 data/output，并发送为前端文件卡片。",
        parameters: {
          type: "object",
          properties: {
            name: {
              type: "string",
              description: "输出文件名，建议使用 .md、.txt 或 .json 后缀。"
            },
            content: {
              type: "string",
              description: "需要写入文件的 UTF-8 文本内容。"
            }
          },
          required: ["name", "content"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "write_word",
        description: "按照 docs/密码应用方案.docx 模板生成专业密码应用方案 Word 文档，并发送为前端文件卡片。",
        parameters: {
          type: "object",
          properties: {
            name: {
              type: "string",
              description: "输出 Word 文件名，建议以 .docx 结尾。"
            },
            prompt: {
              type: "string",
              description: "用户需求或系统事实描述。"
            },
            content: {
              type: "string",
              description: "已经生成的方案正文 Markdown。"
            }
          },
          required: ["content"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "image_generate",
        description: "生成密码应用技术架构图或业务流程图，并发送为前端文件卡片。",
        parameters: {
          type: "object",
          properties: {
            kind: {
              type: "string",
              enum: ["architecture", "flow"],
              description: "architecture 表示技术架构图，flow 表示业务流程图。"
            },
            prompt: {
              type: "string",
              description: "图片生成说明，应包含系统名称、设备、流程和中文标签要求。"
            }
          },
          required: ["kind", "prompt"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "send_file",
        description: "把 data/output 中已经存在的文件发送给前端显示为文件卡片。",
        parameters: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "data/output 中的文件路径。"
            }
          },
          required: ["path"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "write_pdf",
        description: "将 data/output 中已经存在的 Word 方案文档导出为 PDF，并发送为前端文件卡片。需要本机安装 LibreOffice/soffice。",
        parameters: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "data/output 中的 .docx 文件路径。"
            }
          },
          required: ["path"],
          additionalProperties: false
        }
      }
    }
  ];

  if (options.includeExecBash) {
    tools.push({
      type: "function",
      function: {
        name: "exec_bash",
        description: "执行诊断命令。仅用于非破坏性的本地检查，禁止删除、移动、覆盖系统文件。",
        parameters: {
          type: "object",
          properties: {
            command: {
              type: "string",
              description: "要执行的命令。"
            }
          },
          required: ["command"],
          additionalProperties: false
        }
      }
    });
  }

  return tools;
}

export async function executeAgentToolCall(
  toolCall: ChatCompletionMessageToolCall,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  if (toolCall.type !== "function") {
    return {
      toolName: "read_file",
      summary: `暂不支持 custom tool：${toolCall.type}`,
      content: `Unsupported tool call type: ${toolCall.type}`
    };
  }

  const toolName = toolCall.function.name as BuiltinToolName;
  const args = parseToolArguments(toolCall.function.arguments);

  switch (toolName) {
    case "time": {
      const time = getCurrentTimeText();
      return { toolName, summary: time, content: time };
    }
    case "web_search":
      return executeWebSearch(args);
    case "read_file":
      return executeReadFile(args, context, "read_file");
    case "read_word":
      return executeReadFile(args, context, "read_word");
    case "read_pdf":
      return executeReadFile(args, context, "read_pdf");
    case "write_file":
      return executeWriteFile(args, context);
    case "write_word":
      return executeWriteWord(args, context);
    case "image_generate":
      return executeImageGenerate(args, context);
    case "send_file":
      return executeSendFile(args, context);
    case "write_pdf":
      return executeWritePdf(args, context);
    case "exec_bash":
      return executeBash(args, context);
    default:
      return {
        toolName: "read_file",
        summary: `未知工具：${String(toolName)}`,
        content: `Unknown tool: ${String(toolName)}`
      };
  }
}

export function parseToolArguments(raw: string): Record<string, unknown> {
  if (!raw.trim()) return {};
  try {
    const parsed = JSON.parse(raw) as unknown;
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? (parsed as Record<string, unknown>) : {};
  } catch {
    return {};
  }
}

export async function searchWeb(
  query: string,
  options: { maxResults?: number; fetchImpl?: typeof fetch } = {}
): Promise<WebSearchResult[]> {
  const fetchImpl = options.fetchImpl ?? fetch;
  const maxResults = Math.min(Math.max(options.maxResults ?? 5, 1), 8);
  const url = `https://duckduckgo.com/html/?q=${encodeURIComponent(query)}`;
  const response = await fetchImpl(url, {
    headers: {
      "user-agent": "PwdSafeAgent/0.1 (+https://local.agent)"
    }
  });

  if (!response.ok) {
    throw new Error(`web_search 请求失败：${response.status} ${response.statusText}`);
  }

  const html = await response.text();
  return parseDuckDuckGoHtml(html).slice(0, maxResults);
}

export function parseDuckDuckGoHtml(html: string): WebSearchResult[] {
  const results: WebSearchResult[] = [];
  const anchorMatches = Array.from(
    html.matchAll(/<a[^>]+class="[^"]*result__a[^"]*"[^>]+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi)
  );
  const blocks = anchorMatches.length
    ? anchorMatches.map((match, index) => {
        const next = anchorMatches[index + 1];
        const end = next?.index ?? html.length;
        return html.slice(match.index ?? 0, end);
      })
    : html.match(/<div[^>]+class="[^"]*result[^"]*"[\s\S]*?(?=<div[^>]+class="[^"]*result[^"]*"|<\/body>|$)/gi) ?? [];

  for (const block of blocks) {
    const linkMatch =
      block.match(/<a[^>]+class="[^"]*result__a[^"]*"[^>]+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/i) ??
      block.match(/<a[^>]+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/i);
    if (!linkMatch) continue;

    const snippetMatch = block.match(/<a[^>]+class="[^"]*result__snippet[^"]*"[^>]*>([\s\S]*?)<\/a>/i) ??
      block.match(/<div[^>]+class="[^"]*result__snippet[^"]*"[^>]*>([\s\S]*?)<\/div>/i);
    const title = cleanHtmlText(linkMatch[2]);
    const url = normalizeDuckDuckGoUrl(decodeHtml(linkMatch[1]));
    const snippet = cleanHtmlText(snippetMatch?.[1] ?? "");

    if (title && url && !results.some((item) => item.url === url)) {
      results.push({ title, url, snippet });
    }
  }

  return results;
}

function executeWebSearch(args: Record<string, unknown>): Promise<AgentToolExecutionResult> {
  const query = readStringArg(args, "query");
  const maxResults = readNumberArg(args, "max_results", 5);
  if (!query) {
    return Promise.resolve({
      toolName: "web_search",
      summary: "缺少搜索关键词",
      content: "web_search failed: missing query"
    });
  }

  return searchWeb(query, { maxResults }).then((results) => ({
    toolName: "web_search",
    summary: results.length ? `已检索“${query}”，返回 ${results.length} 条结果` : `未检索到“${query}”的公开结果`,
    content: formatWebSearchResults(query, results)
  }));
}

async function executeReadFile(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext,
  requestedToolName: "read_file" | "read_word" | "read_pdf"
): Promise<AgentToolExecutionResult> {
  const inputPath = readStringArg(args, "path");
  if (!inputPath) {
    return { toolName: requestedToolName, summary: "缺少文件路径", content: `${requestedToolName} failed: missing path` };
  }

  const filePath = resolveToolPath(inputPath, context.rootDir);
  assertPathInside(filePath, context.allowedReadDirs, "read_file");
  const extension = extname(filePath).toLowerCase();
  if (requestedToolName === "read_word" && extension !== ".docx") {
    return {
      toolName: "read_word",
      summary: "read_word 只能读取 .docx 文件",
      content: `read_word failed: unsupported extension ${extension || "(none)"}`
    };
  }
  if (requestedToolName === "read_pdf" && extension !== ".pdf") {
    return {
      toolName: "read_pdf",
      summary: "read_pdf 只能读取 .pdf 文件",
      content: `read_pdf failed: unsupported extension ${extension || "(none)"}`
    };
  }

  const result = await readDocumentText(filePath, 18000);
  return {
    toolName: result.toolName,
    summary: result.summary,
    content: result.content || result.summary
  };
}

async function executeWriteWord(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const content = readStringArg(args, "content");
  if (!content) {
    return { toolName: "write_word", summary: "缺少方案正文", content: "write_word failed: missing content" };
  }

  const prompt = readStringArg(args, "prompt");
  const rawName = readStringArg(args, "name") || `${context.sessionTitle}-密码应用方案.docx`;
  const safeName = sanitizeFileName(rawName.endsWith(".docx") ? rawName : `${rawName}.docx`);
  const outputPath = join(context.outputDir, `${Date.now().toString(36)}-${safeName}`);
  const templatePath = join(context.rootDir, "docs", "密码应用方案.docx");
  const result = await writeSchemeDocxFromTemplate(templatePath, outputPath, {
    prompt: prompt || context.sessionTitle,
    memory: context.memory,
    generatedMarkdown: content
  });

  return {
    toolName: "write_word",
    summary: `已生成 ${result.fileName}，填充 ${result.filledFields.length} 项，待补充 ${result.missingFields.length} 项`,
    content: [
      `write_word completed: ${result.outputPath}`,
      `填充字段：${result.filledFields.join("、") || "无"}`,
      `待补充字段：${result.missingFields.join("、") || "无"}`
    ].join("\n"),
    artifactPath: result.outputPath
  };
}

async function executeWriteFile(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const name = sanitizeFileName(readStringArg(args, "name") || "agent-output.md");
  const content = readStringArg(args, "content");
  if (!content) {
    return { toolName: "write_file", summary: "缺少写入内容", content: "write_file failed: missing content" };
  }

  const outputPath = join(context.outputDir, `${Date.now().toString(36)}-${name}`);
  await writeUtf8File(outputPath, content);
  return {
    toolName: "write_file",
    summary: `已写入 ${basename(outputPath)}`,
    content: compactText(content, 12000),
    artifactPath: outputPath
  };
}

async function executeImageGenerate(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const kind = normalizeDiagramKind(readStringArg(args, "kind"));
  const prompt = readStringArg(args, "prompt") || "生成密码应用方案配图";
  const result = await generateDiagramImage(
    {
      apiKey: process.env.OPENAI_API_KEY,
      baseUrl: context.settings.openai.baseUrl,
      imageModel: context.settings.openai.imageModel,
      imageSize: context.settings.openai.imageSize,
      imageQuality: context.settings.openai.imageQuality,
      requestTimeoutMs: context.settings.openai.requestTimeoutMs
    },
    {
      kind,
      sessionTitle: context.sessionTitle,
      prompt,
      memory: context.memory,
      generatedMarkdown: "",
      outputDir: context.outputDir,
      artifactStamp: Date.now().toString(36)
    },
    context.signal
  );

  return {
    toolName: "image_generate",
    summary: `已生成 ${result.fileName}`,
    content: `image_generate completed: ${result.fileName}`,
    artifactPath: result.outputPath
  };
}

function executeSendFile(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): AgentToolExecutionResult {
  const inputPath = readStringArg(args, "path");
  if (!inputPath) {
    return { toolName: "send_file", summary: "缺少文件路径", content: "send_file failed: missing path" };
  }

  const filePath = resolveToolPath(inputPath, context.rootDir);
  assertPathInside(filePath, [context.outputDir], "send_file");
  return {
    toolName: "send_file",
    summary: `已发送 ${basename(filePath)}`,
    content: `send_file completed: ${filePath}`,
    artifactPath: filePath
  };
}

async function executeWritePdf(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const inputPath = readStringArg(args, "path");
  if (!inputPath) {
    return { toolName: "write_pdf", summary: "缺少 Word 文件路径", content: "write_pdf failed: missing path" };
  }

  const filePath = resolveToolPath(inputPath, context.rootDir);
  assertPathInside(filePath, [context.outputDir], "write_pdf");
  const result = await exportDocxToPdf(filePath, context.outputDir, {
    libreOfficePath: context.settings.document.libreOfficePath,
    timeoutMs: context.settings.openai.requestTimeoutMs
  });

  return {
    toolName: "write_pdf",
    summary: result.summary,
    content: result.summary,
    artifactPath: result.status === "success" ? result.outputPath : undefined
  };
}

async function executeBash(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const command = readStringArg(args, "command");
  if (!context.execBashEnabled) {
    return {
      toolName: "exec_bash",
      summary: "exec_bash 未启用",
      content: "exec_bash is disabled. Set AGENT_EXEC_BASH_ENABLED=true only in trusted local development."
    };
  }
  if (!command) {
    return { toolName: "exec_bash", summary: "缺少命令", content: "exec_bash failed: missing command" };
  }
  if (looksDestructiveCommand(command)) {
    return {
      toolName: "exec_bash",
      summary: "拒绝执行高风险命令",
      content: `Rejected destructive command: ${command}`
    };
  }

  const { stdout, stderr } = await execAsync(command, {
    cwd: context.rootDir,
    timeout: 15000,
    maxBuffer: 1024 * 256,
    windowsHide: true
  });
  const output = compactText([stdout, stderr].filter(Boolean).join("\n"), 12000);
  return {
    toolName: "exec_bash",
    summary: output ? "命令执行完成" : "命令执行完成，无输出",
    content: output || "(empty output)"
  };
}

function formatWebSearchResults(query: string, results: WebSearchResult[]): string {
  if (!results.length) return `No web search results for: ${query}`;
  return [
    `Web search results for: ${query}`,
    ...results.map((result, index) =>
      [`${index + 1}. ${result.title}`, `URL: ${result.url}`, result.snippet ? `摘要: ${result.snippet}` : ""]
        .filter(Boolean)
        .join("\n")
    )
  ].join("\n\n");
}

function cleanHtmlText(value: string): string {
  return decodeHtml(value.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim());
}

function decodeHtml(value: string): string {
  return value
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&#39;/g, "'")
    .replace(/&#x27;/g, "'")
    .replace(/&#x2F;/g, "/");
}

function normalizeDuckDuckGoUrl(rawUrl: string): string {
  if (rawUrl.startsWith("//")) return `https:${rawUrl}`;
  if (rawUrl.startsWith("/l/?")) {
    const match = rawUrl.match(/[?&]uddg=([^&]+)/);
    if (match?.[1]) return decodeURIComponent(match[1]);
  }
  return rawUrl;
}

function resolveToolPath(inputPath: string, rootDir: string): string {
  return isAbsolute(inputPath) ? resolve(inputPath) : resolve(rootDir, inputPath);
}

function assertPathInside(filePath: string, allowedDirs: string[], toolName: string): void {
  const normalizedPath = resolve(filePath).toLowerCase();
  const normalizedDirs = allowedDirs.map((dir) => resolve(dir).toLowerCase());
  const isAllowed = normalizedDirs.some((dir) => normalizedPath === dir || normalizedPath.startsWith(`${dir}\\`) || normalizedPath.startsWith(`${dir}/`));
  if (!isAllowed) {
    throw new Error(`${toolName} 只能访问允许目录内的文件`);
  }
}

function readStringArg(args: Record<string, unknown>, key: string): string {
  const value = args[key];
  return typeof value === "string" ? value.trim() : "";
}

function readNumberArg(args: Record<string, unknown>, key: string, fallback: number): number {
  const value = args[key];
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string" && value.trim() && Number.isFinite(Number(value))) return Number(value);
  return fallback;
}

function normalizeDiagramKind(value: string): DiagramKind {
  return value === "flow" ? "flow" : "architecture";
}

function looksDestructiveCommand(command: string): boolean {
  return /\b(rm\s+-rf|del\s+\/|rmdir\s+\/s|remove-item\b|format\b|shutdown\b|restart-computer\b|git\s+reset\s+--hard)\b/i.test(command);
}
