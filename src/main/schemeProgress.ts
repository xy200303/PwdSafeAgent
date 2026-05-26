import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import type { SchemeProgressItem, SchemeProgressSection, SchemeSectionStatus } from "../shared/types";

interface SchemeTemplateJson {
  sections?: unknown[];
}

export interface CreateSchemeProgressItemInput {
  id: string;
  docsDir: string;
  createdAt: string;
}

export interface SchemeProgressUpdateInput {
  section?: string;
  anchorIds?: string[];
  status: SchemeSectionStatus;
  detail?: string;
  artifactName?: string;
  cascade?: boolean;
}

export function createSchemeProgressItem(input: CreateSchemeProgressItemInput): SchemeProgressItem {
  const sections = loadSchemeTemplateSections(input.docsDir);
  return {
    id: input.id,
    kind: "scheme_progress",
    title: "方案章节生成",
    status: "pending",
    total: sections.length,
    drafted: 0,
    completed: 0,
    failed: 0,
    sections,
    createdAt: input.createdAt
  };
}

export function loadSchemeTemplateSections(docsDir: string): SchemeProgressSection[] {
  const templateJsonPath = join(docsDir, "密码应用方案.template.json");
  if (!existsSync(templateJsonPath)) return [];

  const parsed = JSON.parse(readFileSync(templateJsonPath, "utf-8")) as SchemeTemplateJson;
  const rawSections = Array.isArray(parsed.sections) ? parsed.sections : [];
  return rawSections
    .map((raw, index) => normalizeTemplateSection(raw, index))
    .filter((section): section is SchemeProgressSection => Boolean(section));
}

export function applySchemeProgressUpdate(
  item: SchemeProgressItem,
  update: SchemeProgressUpdateInput,
  updatedAt: string
): SchemeProgressItem {
  const targetIds = resolveSchemeSectionIds(item.sections, update);
  const sections = item.sections.map((section) => {
    if (!targetIds.has(section.id)) return section;
    return {
      ...section,
      status: update.status,
      detail: update.detail ?? section.detail,
      updatedAt
    };
  });

  const completed = sections.filter((section) => section.status === "completed").length;
  const drafted = sections.filter((section) => section.status === "drafted").length;
  const failed = sections.filter((section) => section.status === "failed").length;
  const activeSectionId =
    update.status === "running" || update.status === "drafting"
      ? Array.from(targetIds)[0]
      : clearActiveSection(item.activeSectionId, targetIds);
  const status = resolveOverallStatus(item.status, update.status, {
    total: sections.length,
    drafted,
    completed,
    failed
  });

  return {
    ...item,
    status,
    drafted,
    completed,
    failed,
    activeSectionId,
    artifactName: update.artifactName ?? item.artifactName,
    detail: update.detail ?? item.detail,
    sections,
    updatedAt
  };
}

export function settleSchemeProgressItem(
  item: SchemeProgressItem,
  status: SchemeProgressItem["status"],
  updatedAt: string
): SchemeProgressItem {
  return {
    ...item,
    status,
    activeSectionId: undefined,
    updatedAt
  };
}

export function extractTemplateAnchorIds(text: string): string[] {
  return Array.from(new Set(text.match(/\bsec_\d+(?:_\d+)*\b/g) ?? []));
}

export function resolveSchemeSectionIds(
  sections: SchemeProgressSection[],
  update: Pick<SchemeProgressUpdateInput, "section" | "anchorIds" | "cascade">
): Set<string> {
  const ids = new Set<string>();
  const sectionIds = new Set(sections.map((section) => section.id));
  for (const anchorId of update.anchorIds ?? []) {
    if (sectionIds.has(anchorId)) ids.add(anchorId);
  }

  const matched = update.section ? findSchemeSection(sections, update.section) : undefined;
  if (matched) ids.add(matched.id);

  if (!update.cascade || ids.size === 0) return ids;

  const cascadeIds = new Set(ids);
  for (const id of ids) {
    const root = sections.find((section) => section.id === id);
    if (!root) continue;
    for (const section of sections) {
      if (section.id === root.id) continue;
      if (root.number && section.number.startsWith(`${root.number}.`)) cascadeIds.add(section.id);
      if (section.id.startsWith(`${root.id}_`)) cascadeIds.add(section.id);
    }
  }
  return cascadeIds;
}

function normalizeTemplateSection(raw: unknown, index: number): SchemeProgressSection | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const record = raw as Record<string, unknown>;
  const number = readString(record, "number");
  const title = readString(record, "title");
  const id = readString(record, "id") || (number ? `sec_${number.replace(/\./g, "_")}` : `sec_${index + 1}`);
  if (!number || !title) return undefined;

  return {
    id,
    number,
    title,
    headingLevel: readNumber(record, "headingLevel", number.split(".").length),
    status: "pending",
    writingHint: readString(record, "writingHint") || undefined,
    relatedTables: readStringArray(record, "relatedTables"),
    relatedFigures: readStringArray(record, "relatedFigures")
  };
}

function findSchemeSection(sections: SchemeProgressSection[], target: string): SchemeProgressSection | undefined {
  const normalizedTarget = normalizeSectionLookup(target);
  if (!normalizedTarget) return undefined;

  return sections.find((section) =>
    [
      section.id,
      section.number,
      section.title,
      `${section.number}${section.title}`,
      `${section.number}.${section.title}`,
      `${section.number} ${section.title}`,
      `${section.number}、${section.title}`
    ].some((candidate) => normalizeSectionLookup(candidate) === normalizedTarget)
  );
}

function normalizeSectionLookup(value: string): string {
  return value
    .toLowerCase()
    .replace(/[第章节]/g, "")
    .replace(/[\s.。．、,，:：()（）\[\]【】《》"'“”‘’_-]+/g, "")
    .trim();
}

function resolveOverallStatus(
  current: SchemeProgressItem["status"],
  sectionStatus: SchemeSectionStatus,
  counts: { total: number; drafted: number; completed: number; failed: number }
): SchemeProgressItem["status"] {
  if (sectionStatus === "running" || sectionStatus === "drafting") return "running";
  if (counts.failed > 0) return "failed";
  if (counts.total > 0 && counts.completed === counts.total) return "completed";
  if (sectionStatus === "drafted" || counts.drafted > 0 || counts.completed > 0) return "running";
  if (current === "pending" && counts.failed > 0) return "failed";
  return current;
}

function clearActiveSection(activeSectionId: string | undefined, targetIds: Set<string>): string | undefined {
  if (!activeSectionId) return undefined;
  return targetIds.has(activeSectionId) ? undefined : activeSectionId;
}

function readString(record: Record<string, unknown>, key: string): string {
  const value = record[key];
  return typeof value === "string" ? value.trim() : "";
}

function readNumber(record: Record<string, unknown>, key: string, fallback: number): number {
  const value = record[key];
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function readStringArray(record: Record<string, unknown>, key: string): string[] | undefined {
  const value = record[key];
  if (!Array.isArray(value)) return undefined;
  const items = value.filter((item): item is string => typeof item === "string" && Boolean(item.trim()));
  return items.length ? items : undefined;
}
