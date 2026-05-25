import { basename, dirname, extname } from "node:path";
import { mkdir, readFile, stat, writeFile } from "node:fs/promises";
import mammoth from "mammoth";
import { PDFParse } from "pdf-parse";

export type BuiltinToolName =
  | "time"
  | "remember_project"
  | "read_file"
  | "read_word"
  | "read_pdf"
  | "write_file"
  | "create_word"
  | "write_word"
  | "write_pdf"
  | "image_generate"
  | "web_search"
  | "exec_bash"
  | "send_file";

export interface ReadToolResult {
  toolName: BuiltinToolName;
  sourceName: string;
  content: string;
  summary: string;
  charCount: number;
}

const TEXT_EXTENSIONS = new Set([".md", ".txt", ".json", ".csv", ".log", ".yaml", ".yml"]);

export function getReadToolName(filePath: string): BuiltinToolName {
  const ext = extname(filePath).toLowerCase();
  if (ext === ".docx") return "read_word";
  if (ext === ".pdf") return "read_pdf";
  return "read_file";
}

export function compactText(text: string, maxChars: number): string {
  const normalized = text.replace(/\r\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  if (normalized.length <= maxChars) return normalized;
  return `${normalized.slice(0, maxChars)}\n\n[内容过长，已截断 ${normalized.length - maxChars} 字符]`;
}

export function sanitizeFileName(name: string): string {
  return name.replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_").replace(/\s+/g, " ").trim() || "artifact";
}

export function getCurrentTimeText(): string {
  return new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    dateStyle: "full",
    timeStyle: "medium"
  }).format(new Date());
}

export async function readDocumentText(filePath: string, maxChars: number): Promise<ReadToolResult> {
  const ext = extname(filePath).toLowerCase();
  const sourceName = basename(filePath);
  const fileStat = await stat(filePath);
  let content = "";

  if (ext === ".docx") {
    const result = await mammoth.extractRawText({ path: filePath });
    content = result.value;
  } else if (ext === ".pdf") {
    const data = await readFile(filePath);
    const parser = new PDFParse({ data });
    try {
      const result = await parser.getText();
      content = result.text;
    } finally {
      await parser.destroy();
    }
  } else if (TEXT_EXTENSIONS.has(ext)) {
    content = await readFile(filePath, "utf-8");
  } else {
    return {
      toolName: getReadToolName(filePath),
      sourceName,
      content: "",
      summary: `暂不支持解析 ${ext || "未知"} 文件，已保留附件引用`,
      charCount: 0
    };
  }

  const compacted = compactText(content, maxChars);
  return {
    toolName: getReadToolName(filePath),
    sourceName,
    content: compacted,
    summary: `已读取 ${sourceName}，大小 ${formatBytes(fileStat.size)}，抽取 ${compacted.length} 字符`,
    charCount: compacted.length
  };
}

export async function writeUtf8File(filePath: string, content: string): Promise<void> {
  await mkdir(dirname(filePath), { recursive: true });
  await writeFile(filePath, content, "utf-8");
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}
