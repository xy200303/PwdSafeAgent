import { mkdir, writeFile } from "node:fs/promises";
import { basename, dirname, join } from "node:path";
import OpenAI from "openai";
import type { ImageGenerateParamsNonStreaming, ImagesResponse } from "openai/resources/images";
import { compactText, sanitizeFileName } from "./agentTools";
import { hasDiagramArtifactIntent } from "./artifactIntent";

export type DiagramKind = "architecture" | "flow";

export interface ImageGenerationConfig {
  apiKey?: string;
  baseUrl: string;
  imageModel: string;
  imageSize: string;
  imageQuality: string;
  requestTimeoutMs: number;
}

export interface DiagramGenerationInput {
  kind: DiagramKind;
  label?: string;
  sessionTitle: string;
  prompt: string;
  memory: string;
  generatedMarkdown: string;
  outputDir: string;
  artifactStamp: string;
}

export interface GeneratedDiagramResult {
  outputPath: string;
  fileName: string;
  prompt: string;
  mode: "openai" | "local-svg";
  model: string;
}

const DIAGRAM_CONTENT_KEYWORDS = /架构|流程|拓扑|密码|密钥|加密|签名|验签|证书|身份鉴别|日志审计|合规|等保/;

export function shouldGenerateDiagramArtifacts(prompt: string, content: string): boolean {
  if (!hasDiagramArtifactIntent(prompt)) return false;
  return DIAGRAM_CONTENT_KEYWORDS.test(`${prompt}\n${content}`);
}

export function buildDiagramPrompt(input: DiagramGenerationInput): string {
  const title = input.label || (input.kind === "architecture" ? "密码应用技术架构图" : "典型业务密码应用流程图");
  const context = compactText(`${input.prompt}\n\n${input.generatedMarkdown}\n\n${input.memory}`, 9000);
  const focus =
    input.kind === "architecture"
      ? "展示用户、应用系统、统一身份认证/业务系统、密码服务管理平台、服务器密码机、签名验签服务器、数字证书认证系统、数据库、日志审计与运维入口之间的关系。"
      : "展示业务发起、身份鉴别、电子签名/验签、数据加密存储、密钥调用、日志审计和异常处置的端到端流程。";

  return [
    `生成一张用于正式密码应用方案文档的${title}。`,
    "视觉风格：专业政务/企业安全方案插图，蓝灰色系，扁平矢量信息图，白底，线条清晰，中文标签清晰可读，不要照片质感。",
    `内容重点：${focus}`,
    "输出要求：使用分区、箭头、编号和简洁中文标签；不要出现乱码、英文占位符、水印、品牌 Logo 或虚构产品厂商。",
    "如果资料不足，应优先依据已确认信息绘制通用结构，不要在最终图中出现“待补充”“需确认”“XXX”等占位标识。",
    "",
    "参考上下文：",
    context
  ].join("\n");
}

export async function generateDiagramImage(
  config: ImageGenerationConfig,
  input: DiagramGenerationInput,
  signal?: AbortSignal
): Promise<GeneratedDiagramResult> {
  const prompt = buildDiagramPrompt(input);
  const diagramName = input.label || (input.kind === "architecture" ? "密码应用技术架构图" : "典型业务密码应用流程图");
  const safeBaseName = sanitizeFileName(`${input.sessionTitle}-${diagramName}-${input.artifactStamp}`);

  if (!config.apiKey) {
    const outputPath = join(input.outputDir, `${safeBaseName}.svg`);
    await writeLocalSvgDiagram(outputPath, input.kind, diagramName, input);
    return {
      outputPath,
      fileName: basename(outputPath),
      prompt,
      mode: "local-svg",
      model: "local-svg"
    };
  }

  const client = new OpenAI({
    apiKey: config.apiKey,
    baseURL: config.baseUrl,
    timeout: config.requestTimeoutMs
  });
  const outputPath = join(input.outputDir, `${safeBaseName}.png`);
  const request = buildImageRequest(config, prompt);
  const response = await client.images.generate(request, { signal });
  await writeImageResponse(outputPath, response);

  return {
    outputPath,
    fileName: basename(outputPath),
    prompt,
    mode: "openai",
    model: config.imageModel
  };
}

function buildImageRequest(config: ImageGenerationConfig, prompt: string): ImageGenerateParamsNonStreaming {
  const request: ImageGenerateParamsNonStreaming = {
    model: config.imageModel,
    prompt,
    n: 1,
    quality: normalizeImageQuality(config.imageQuality),
    size: normalizeImageSize(config.imageSize),
    stream: false
  };

  if (isGptImageModel(config.imageModel)) {
    request.output_format = "png";
  } else {
    request.response_format = "b64_json";
  }

  return request;
}

async function writeImageResponse(outputPath: string, response: ImagesResponse): Promise<void> {
  const first = response.data?.[0];
  if (!first) {
    throw new Error("生图接口未返回图片数据");
  }

  await mkdir(dirname(outputPath), { recursive: true });
  if (first.b64_json) {
    await writeFile(outputPath, Buffer.from(first.b64_json, "base64"));
    return;
  }

  if (first.url) {
    const downloaded = await fetch(first.url);
    if (!downloaded.ok) {
      throw new Error(`下载生图结果失败：${downloaded.status} ${downloaded.statusText}`);
    }
    const arrayBuffer = await downloaded.arrayBuffer();
    await writeFile(outputPath, Buffer.from(arrayBuffer));
    return;
  }

  throw new Error("生图接口未返回 b64_json 或 url");
}

async function writeLocalSvgDiagram(
  outputPath: string,
  kind: DiagramKind,
  title: string,
  input: DiagramGenerationInput
): Promise<void> {
  const svg = createLocalSvgDiagram(kind, title, input);
  await mkdir(dirname(outputPath), { recursive: true });
  await writeFile(outputPath, svg, "utf-8");
}

function createLocalSvgDiagram(kind: DiagramKind, title: string, input: DiagramGenerationInput): string {
  const appName = pickFact(input.prompt, ["系统名称", "项目名称", "应用系统"]) || input.sessionTitle || "应用系统";
  const orgName = pickFact(input.prompt, ["建设单位", "单位名称"]) || "建设单位";
  const subtitle =
    kind === "architecture"
      ? "本地演示图：配置 OPENAI_IMAGE_API_KEY 或 OPENAI_API_KEY 后将使用 gpt-image-2 生成正式图像"
      : "本地演示图：配置 OPENAI_IMAGE_API_KEY 或 OPENAI_API_KEY 后将使用 gpt-image-2 生成正式流程图";
  const nodes =
    kind === "architecture"
      ? [
          ["用户/终端", 90, 245],
          [appName, 360, 245],
          ["密码服务管理平台", 660, 155],
          ["服务器密码机", 660, 340],
          ["签名验签服务器", 930, 155],
          ["数字证书认证系统", 930, 340],
          ["数据库/文件存储", 360, 485],
          ["日志审计", 930, 485]
        ]
      : [
          ["1 业务发起", 90, 245],
          ["2 身份鉴别", 310, 245],
          ["3 调用密码服务", 530, 245],
          ["4 签名/验签", 750, 245],
          ["5 加密存储", 970, 245],
          ["6 日志审计", 530, 455]
        ];

  const nodeMarkup = nodes
    .map(
      ([label, x, y]) => `
        <rect x="${x}" y="${y}" width="170" height="74" rx="14" fill="#ffffff" stroke="#2b5f8f" stroke-width="2"/>
        <text x="${Number(x) + 85}" y="${Number(y) + 43}" text-anchor="middle" font-size="18" fill="#12344d">${escapeXml(
          String(label)
        )}</text>`
    )
    .join("\n");
  const arrowMarkup =
    kind === "architecture"
      ? [
          arrow(260, 282, 360, 282),
          arrow(530, 282, 660, 200),
          arrow(530, 282, 660, 385),
          arrow(830, 192, 930, 192),
          arrow(830, 377, 930, 377),
          arrow(445, 319, 445, 485),
          arrow(745, 414, 930, 522)
        ].join("\n")
      : [
          arrow(260, 282, 310, 282),
          arrow(480, 282, 530, 282),
          arrow(700, 282, 750, 282),
          arrow(920, 282, 970, 282),
          arrow(615, 319, 615, 455)
        ].join("\n");

  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1280" height="720" viewBox="0 0 1280 720">
  <defs>
    <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#f7fbff"/>
      <stop offset="100%" stop-color="#e9f2fb"/>
    </linearGradient>
    <marker id="arrow" markerWidth="12" markerHeight="12" refX="10" refY="6" orient="auto">
      <path d="M2,2 L10,6 L2,10 Z" fill="#2b5f8f"/>
    </marker>
  </defs>
  <rect width="1280" height="720" fill="url(#bg)"/>
  <rect x="48" y="48" width="1184" height="624" rx="28" fill="#ffffff" stroke="#c8d8e6"/>
  <text x="80" y="105" font-size="30" font-weight="700" fill="#0d2b45">${escapeXml(title)}</text>
  <text x="80" y="140" font-size="16" fill="#5e7184">${escapeXml(orgName)} · ${escapeXml(appName)} · ${escapeXml(
    subtitle
  )}</text>
  ${arrowMarkup}
  ${nodeMarkup}
  <text x="80" y="625" font-size="15" fill="#6b7f91">说明：本图为离线占位示意图，可在配置模型后由 image_generate 工具重新生成专业图像。</text>
</svg>`;
}

function arrow(x1: number, y1: number, x2: number, y2: number): string {
  return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#2b5f8f" stroke-width="3" marker-end="url(#arrow)"/>`;
}

function pickFact(source: string, labels: string[]): string {
  for (const label of labels) {
    const match = source.match(new RegExp(`${label}\\s*[：:]\\s*([^\\n，。；;]+)`));
    if (match?.[1]) return match[1].trim().slice(0, 28);
  }
  return "";
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function isGptImageModel(model: string): boolean {
  return /gpt-image|chatgpt-image/i.test(model);
}

function normalizeImageQuality(value: string): ImageGenerateParamsNonStreaming["quality"] {
  if (value === "standard" || value === "hd" || value === "low" || value === "medium" || value === "high" || value === "auto") {
    return value;
  }
  return "high";
}

function normalizeImageSize(value: string): ImageGenerateParamsNonStreaming["size"] {
  return value || "1536x1024";
}
