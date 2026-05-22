import { readFile, stat } from "node:fs/promises";
import { extname } from "node:path";
import { compactText, readDocumentText } from "./agentTools";
import type { ArtifactPreview, ArtifactSummary } from "../shared/types";

const TEXT_EXTENSIONS = new Set([".md", ".txt", ".json", ".log", ".csv", ".yaml", ".yml"]);
const IMAGE_MIME_TYPES: Record<string, string> = {
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".webp": "image/webp",
  ".svg": "image/svg+xml"
};

export async function buildArtifactPreview(
  artifact: ArtifactSummary,
  options: { maxTextChars?: number; maxDataBytes?: number } = {}
): Promise<ArtifactPreview> {
  const maxTextChars = options.maxTextChars ?? 60000;
  const maxDataBytes = options.maxDataBytes ?? 12 * 1024 * 1024;
  const extension = extname(artifact.path).toLowerCase();
  const fileStat = await stat(artifact.path);

  if (IMAGE_MIME_TYPES[extension]) {
    return buildDataPreview(artifact, IMAGE_MIME_TYPES[extension], "image", fileStat.size, maxDataBytes);
  }

  if (extension === ".pdf") {
    if (fileStat.size <= maxDataBytes) {
      return buildDataPreview(artifact, "application/pdf", "pdf", fileStat.size, maxDataBytes);
    }
    return buildDocumentTextPreview(artifact, maxTextChars, "PDF 文件较大，已改为抽取文本预览。");
  }

  if (extension === ".docx") {
    return buildDocumentTextPreview(artifact, maxTextChars);
  }

  if (TEXT_EXTENSIONS.has(extension)) {
    const text = await readFile(artifact.path, "utf-8");
    return {
      artifactId: artifact.id,
      name: artifact.name,
      kind: artifact.kind,
      mode: "text",
      text: compactText(text, maxTextChars)
    };
  }

  return {
    artifactId: artifact.id,
    name: artifact.name,
    kind: artifact.kind,
    mode: "unsupported",
    summary: "该文件类型暂不支持内置预览，请使用“打开”。"
  };
}

async function buildDataPreview(
  artifact: ArtifactSummary,
  mimeType: string,
  mode: "image" | "pdf",
  size: number,
  maxDataBytes: number
): Promise<ArtifactPreview> {
  if (size > maxDataBytes) {
    return {
      artifactId: artifact.id,
      name: artifact.name,
      kind: artifact.kind,
      mode: "unsupported",
      summary: "文件较大，暂不内置预览，请使用“打开”。"
    };
  }

  const data = await readFile(artifact.path);
  return {
    artifactId: artifact.id,
    name: artifact.name,
    kind: artifact.kind,
    mode,
    mimeType,
    dataUrl: `data:${mimeType};base64,${data.toString("base64")}`
  };
}

async function buildDocumentTextPreview(
  artifact: ArtifactSummary,
  maxTextChars: number,
  summary?: string
): Promise<ArtifactPreview> {
  const result = await readDocumentText(artifact.path, maxTextChars);
  return {
    artifactId: artifact.id,
    name: artifact.name,
    kind: artifact.kind,
    mode: "text",
    text: result.content || result.summary,
    summary
  };
}
