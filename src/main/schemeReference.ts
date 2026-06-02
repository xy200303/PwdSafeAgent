import { readFileSync } from "node:fs";
import { BUILT_IN_STANDARD_REFERENCE_RELATIVE_PATH, getBuiltInStandardReferencePath } from "./templatePaths";

export interface SchemeStandardReferenceInput {
  sectionNumber?: string;
  sectionTitle: string;
  paragraphTasks?: string[];
  projectContext?: string;
  maxLength?: number;
}

interface SchemeStandardDomain {
  title: string;
  code: string;
}

const standardReferenceLinesCache = new Map<string, string[]>();

const STANDARD_DOMAINS: Array<{ pattern: RegExp; domain: SchemeStandardDomain }> = [
  { pattern: /物理|环境|机房|门禁|监控/, domain: { code: "1", title: "物理和环境安全" } },
  { pattern: /网络|通信|边界|拓扑|链路|传输/, domain: { code: "2", title: "网络和通信安全" } },
  { pattern: /设备|计算|主机|服务器|终端|操作系统|数据库|日志/, domain: { code: "3", title: "设备和计算安全" } },
  {
    pattern: /应用和数据安全|业务应用|应用系统|登录用户|数据|存储|访问控制|身份鉴别|完整性|机密性|不可否认|签名|验签/,
    domain: { code: "4", title: "应用和数据安全" }
  },
  { pattern: /管理制度|制度/, domain: { code: "5", title: "管理制度" } },
  { pattern: /人员管理|岗位|培训|职责/, domain: { code: "6", title: "人员管理" } },
  { pattern: /建设运行|实施|运维|运行|评估|演练/, domain: { code: "7", title: "建设运行" } },
  { pattern: /应急|事件处置/, domain: { code: "8", title: "应急处置" } }
];

const LEVEL_HEADING_PREFIX: Record<1 | 2 | 3 | 4, string> = {
  1: "6",
  2: "7",
  3: "8",
  4: "9"
};

export function buildSchemeStandardReferenceContext(docsDir: string, input: SchemeStandardReferenceInput): string {
  const lines = loadStandardReferenceLines(docsDir);
  if (!lines.length) return `标准参考路径：${BUILT_IN_STANDARD_REFERENCE_RELATIVE_PATH}`;

  const focusText = [
    input.sectionNumber || "",
    input.sectionTitle,
    ...(input.paragraphTasks ?? [])
  ]
    .join(" ")
    .trim();
  const searchText = [focusText, input.projectContext || ""].filter(Boolean).join(" ").trim();
  const level = detectProtectionLevel(searchText);
  const targets = selectReferenceTargets(focusText, searchText, level);
  const excerpts = Array.from(
    new Set(targets.map((target) => extractReferenceExcerpt(lines, target)).filter((item): item is string => Boolean(item)))
  );

  if (!excerpts.length) {
    return `标准参考路径：${BUILT_IN_STANDARD_REFERENCE_RELATIVE_PATH}`;
  }

  const header = `标准参考（摘自 ${BUILT_IN_STANDARD_REFERENCE_RELATIVE_PATH}，写作时据此约束方案要求，不要逐字照抄）：`;
  const body = excerpts.map((excerpt, index) => `${index + 1}. ${excerpt}`).join("\n");
  const context = `${header}\n${body}`;
  return trimToLength(context, input.maxLength ?? 2200);
}

function loadStandardReferenceLines(docsDir: string): string[] {
  const cacheKey = getBuiltInStandardReferencePath(docsDir);
  const cached = standardReferenceLinesCache.get(cacheKey);
  if (cached) return cached;
  try {
    const content = readFileSync(cacheKey, "utf-8");
    const lines = content.split(/\r?\n/);
    standardReferenceLinesCache.set(cacheKey, lines);
    return lines;
  } catch {
    return [];
  }
}

function selectReferenceTargets(focusText: string, searchText: string, level?: 1 | 2 | 3 | 4): string[] {
  const targets = ["5 通用要求"];
  if (/背景|依据|原则|框架|概述|总体|合规|要求|设计/.test(focusText)) {
    targets.unshift("4.1 信息系统密码应用技术框架");
  }

  const matchedDomain =
    STANDARD_DOMAINS.find((item) => item.pattern.test(focusText))?.domain ||
    STANDARD_DOMAINS.find((item) => item.pattern.test(searchText))?.domain;
  if (matchedDomain && level) {
    targets.push(`${LEVEL_HEADING_PREFIX[level]}.${matchedDomain.code} ${matchedDomain.title}`);
  } else if (matchedDomain && /合规|要求|分析/.test(focusText)) {
    targets.push(matchedDomain.title);
  }

  if (/密钥/.test(focusText)) {
    targets.push("密钥管理 key management");
    if (level) {
      targets.push(`${LEVEL_HEADING_PREFIX[level]}.7 建设运行`);
      targets.push(`${LEVEL_HEADING_PREFIX[level]}.5 管理制度`);
    }
  }

  if (/身份鉴别/.test(focusText)) {
    targets.push("身份鉴别 identity authentication");
  }

  return targets;
}

function detectProtectionLevel(text: string): 1 | 2 | 3 | 4 | undefined {
  if (/等保级别\s*[：: ]*\s*(四级|4级|第四级)/.test(text) || /\b四级\b|\b第四级\b/.test(text)) return 4;
  if (/等保级别\s*[：: ]*\s*(三级|3级|第三级)/.test(text) || /\b三级\b|\b第三级\b/.test(text)) return 3;
  if (/等保级别\s*[：: ]*\s*(二级|2级|第二级)/.test(text) || /\b二级\b|\b第二级\b/.test(text)) return 2;
  if (/等保级别\s*[：: ]*\s*(一级|1级|第一级)/.test(text) || /\b一级\b|\b第一级\b/.test(text)) return 1;
  return undefined;
}

function extractReferenceExcerpt(lines: string[], target: string): string | undefined {
  const index = findReferenceTargetLine(lines, target);
  if (index < 0) return undefined;
  const heading = stripHeadingMarker(lines[index]);
  const bodyLines: string[] = [];
  for (let cursor = index + 1; cursor < lines.length; cursor += 1) {
    const line = lines[cursor].trim();
    if (!line) continue;
    if (isHeading(line) && bodyLines.length > 0) break;
    if (line.startsWith("#")) continue;
    if (/^(本级要求包括|本级要求为|第[一二三四五]级到第[一二三四五]级的信息系统应符合以下通用要求|使用密码技术的信息系统应符合以下|技术要求主要由|管理要求由)/.test(line)) {
      continue;
    }
    bodyLines.push(line);
    if (bodyLines.length >= 6) break;
  }

  if (!bodyLines.length) return heading;
  return `${heading}：${bodyLines.join("；")}`;
}

function findReferenceTargetLine(lines: string[], target: string): number {
  const normalizedTarget = normalizeReferenceText(target);
  for (const headingsOnly of [true, false]) {
    for (let index = 0; index < lines.length; index += 1) {
      const rawLine = lines[index];
      if (headingsOnly && !isHeading(rawLine.trim())) continue;
      const normalizedLine = normalizeReferenceText(stripHeadingMarker(rawLine));
      if (!normalizedLine) continue;
      if (normalizedLine === normalizedTarget || normalizedLine.includes(normalizedTarget) || normalizedTarget.includes(normalizedLine)) {
        if (headingsOnly && !hasReferenceBody(lines, index)) continue;
        return index;
      }
    }
  }
  return -1;
}

function hasReferenceBody(lines: string[], startIndex: number): boolean {
  for (let cursor = startIndex + 1; cursor < lines.length; cursor += 1) {
    const line = lines[cursor].trim();
    if (!line) continue;
    if (isHeading(line)) return false;
    return true;
  }
  return false;
}

function isHeading(line: string): boolean {
  return /^#/.test(line);
}

function stripHeadingMarker(line: string): string {
  return line.replace(/^#+\s*/, "").trim();
}

function normalizeReferenceText(value: string): string {
  return value
    .toLowerCase()
    .replace(/[—－-]/g, "")
    .replace(/[\s#.:：()（）[\]【】"'“”‘’]+/g, "")
    .trim();
}

function trimToLength(value: string, maxLength: number): string {
  if (value.length <= maxLength) return value;
  return `${value.slice(0, Math.max(0, maxLength - 1)).trim()}…`;
}
