import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";

export interface DocumentConfigSectionLike {
  id: string;
  number: string;
  title: string;
  writingHint?: string;
  paragraphTasks?: string[];
  placeholders?: string[];
  relatedTables?: string[];
  relatedFigures?: string[];
}

export interface DocumentConfigFact {
  key: string;
  value: string;
  source?: string;
}

export interface DocumentGenerationPlanning {
  summary: string;
  assumptions: string[];
  risks: string[];
}

export interface DocumentConfigSectionGroupPlan {
  objective: string;
  outline: string[];
  keyPoints: string[];
  evidenceNeeds: string[];
  openQuestions: string[];
}

export interface DocumentConfigSectionGroupPlanInput extends Partial<DocumentConfigSectionGroupPlan> {
  id?: string;
  sectionGroupId?: string;
  match?: string;
  title?: string;
  order?: number;
}

export interface DocumentConfigSectionGroup {
  id: string;
  order: number;
  title: string;
  sectionIds: string[];
  sectionNumbers: string[];
  sectionTitles: string[];
  writingRules: string[];
  evidenceRules: string[];
  requiredFacts: string[];
  plan?: DocumentConfigSectionGroupPlan;
}

export interface DocumentWordTemplateConfig {
  templatePath?: string;
  templateJsonPath?: string;
  renderMode?: "append" | "full_document" | "template_sections";
  outputNameSuffix?: string;
}

export interface DocumentConfig {
  version: 1;
  profile: string;
  title: string;
  generatedAt: string;
  sourceSummary: string;
  globalRules: string[];
  facts: DocumentConfigFact[];
  gaps: string[];
  planning?: DocumentGenerationPlanning;
  sectionGroups: DocumentConfigSectionGroup[];
  wordTemplate?: DocumentWordTemplateConfig;
}

export interface DocumentProfileSectionRule {
  id?: string;
  match?: string;
  order?: number;
  title?: string;
  rules?: string[];
  writingRules?: string[];
  evidenceRules?: string[];
  requiredFacts?: string[];
}

export interface DocumentProfile {
  version: 1;
  profile: string;
  title?: string;
  titleSuffix?: string;
  globalRules: string[];
  evidenceRules: string[];
  sectionRules: DocumentProfileSectionRule[];
}

const BASE_DOCUMENT_RULES = [
  "只把已确认事实写成确定表述；待补充/需确认的信息不得写成现状。",
  "模板提示、示例资料、标准要求和推断只能作为写作约束，不能冒充项目事实。",
  "每个章节组内部只处理该章节组对应的材料、事实和约束，不跨章节组拼接结论。",
  "未知内容必须显式标记为待补充/需确认。"
];

export function getDocumentConfigPath(outputDir: string): string {
  return join(outputDir, "document-config.json");
}

export function getDocumentProfilePath(docsDir: string, profile: string): string {
  const safeProfile = sanitizeProfileFileName(profile.replace(/\.json$/i, ""));
  return join(docsDir, "document-profiles", `${safeProfile || "generic_document"}.json`);
}

export function loadDocumentConfig(filePath: string): DocumentConfig | undefined {
  if (!existsSync(filePath)) return undefined;
  try {
    const parsed = JSON.parse(readFileSync(filePath, "utf-8")) as unknown;
    return isDocumentConfig(parsed) ? parsed : undefined;
  } catch {
    return undefined;
  }
}

export function loadDocumentProfile(docsDir: string, profile: string): DocumentProfile | undefined {
  const profilePath = getDocumentProfilePath(docsDir, profile);
  return loadDocumentProfileFromFile(profilePath, profile);
}

export function loadDocumentProfileFromFile(profilePath: string, fallbackProfile = "generic_document"): DocumentProfile | undefined {
  if (!existsSync(profilePath)) return undefined;
  try {
    return parseDocumentProfileJson(readFileSync(profilePath, "utf-8"), fallbackProfile);
  } catch {
    return undefined;
  }
}

export function parseDocumentProfileJson(raw: string, fallbackProfile = "generic_document"): DocumentProfile | undefined {
  try {
    const parsed = JSON.parse(raw) as unknown;
    return normalizeDocumentProfile(parsed, fallbackProfile);
  } catch {
    return undefined;
  }
}

export function buildDocumentConfigFromProfile(input: {
  profile: DocumentProfile;
  title: string;
  sourceSummary: string;
  memory?: string;
  facts?: DocumentConfigFact[];
  gaps?: string[];
  planning?: Partial<DocumentGenerationPlanning>;
  sectionGroupPlans?: DocumentConfigSectionGroupPlanInput[];
  templateSections?: DocumentConfigSectionLike[];
  wordTemplate?: DocumentWordTemplateConfig;
}): DocumentConfig {
  const profileSectionGroups = buildSectionGroupsFromSections(input.templateSections ?? []);
  const sectionGroups = applyDocumentSectionGroupPlans(
    applyGlobalEvidenceRules(
      applyProfileSectionRules(profileSectionGroups, input.profile.sectionRules),
      input.profile.evidenceRules
    ),
    input.sectionGroupPlans ?? []
  );
  const planning = normalizeDocumentPlanning(input.planning);
  return {
    version: 1,
    profile: input.profile.profile,
    title: input.title || input.profile.title || "文档",
    generatedAt: new Date().toISOString(),
    sourceSummary: compactConfigText(input.sourceSummary, 12000),
    globalRules: uniqueStrings([...BASE_DOCUMENT_RULES, ...input.profile.globalRules]),
    facts: input.facts ?? extractFactsFromMemory(input.memory || input.sourceSummary),
    gaps: input.gaps ?? extractGapsFromMemory(input.memory || input.sourceSummary),
    ...(planning ? { planning } : {}),
    sectionGroups,
    ...(input.wordTemplate ? { wordTemplate: input.wordTemplate } : {})
  };
}

export function formatDocumentConfigSummary(config: DocumentConfig): string {
  const lines = [
    `# ${config.title}`,
    "",
    `profile: ${config.profile}`,
    `generatedAt: ${config.generatedAt}`,
    "",
    "globalRules:",
    ...config.globalRules.map((rule) => `- ${rule}`),
    "",
    "facts:",
    ...(config.facts.length ? config.facts.map((fact) => `- ${fact.key}：${fact.value}${fact.source ? `（来源：${fact.source}）` : ""}`) : ["- 无"]),
    "",
    "gaps:",
    ...(config.gaps.length ? config.gaps.map((gap) => `- ${gap}`) : ["- 无"]),
    "",
    "planning:",
    ...(config.planning
      ? [
          `- summary：${config.planning.summary}`,
          ...(config.planning.assumptions.length ? config.planning.assumptions.map((item) => `- assumption：${item}`) : []),
          ...(config.planning.risks.length ? config.planning.risks.map((item) => `- risk：${item}`) : [])
        ]
      : ["- 无"]),
    "",
    "wordTemplate:",
    ...(config.wordTemplate
      ? [
          `- templatePath：${config.wordTemplate.templatePath || "默认内置模板"}`,
          `- templateJsonPath：${config.wordTemplate.templateJsonPath || "自动从 docx 解析锚点"}`,
          `- renderMode：${config.wordTemplate.renderMode || "append"}`,
          `- outputNameSuffix：${config.wordTemplate.outputNameSuffix || "无"}`
        ]
      : ["- 无"]),
    "",
    "sectionGroups:"
  ];

  for (const group of config.sectionGroups) {
    lines.push(
      `- ${group.id} | ${group.title} | sections: ${group.sectionTitles.join("、")}`,
      ...(group.plan
        ? [
            `  plan: ${group.plan.objective}`,
            ...group.plan.outline.map((item) => `  outline: ${item}`),
            ...group.plan.keyPoints.map((item) => `  keyPoint: ${item}`),
            ...group.plan.evidenceNeeds.map((item) => `  evidenceNeed: ${item}`),
            ...group.plan.openQuestions.map((item) => `  openQuestion: ${item}`)
          ]
        : []),
      ...group.writingRules.map((rule) => `  rule: ${rule}`)
    );
  }

  return lines.join("\n");
}

export function findDocumentConfigSectionGroupForSection(
  config: DocumentConfig | undefined,
  sectionIdOrNumber: string
): DocumentConfigSectionGroup | undefined {
  if (!config) return undefined;
  const normalized = normalizeLookup(sectionIdOrNumber);
  return config.sectionGroups.find((group) =>
    [
      group.id,
      ...group.sectionIds,
      ...group.sectionNumbers,
      ...group.sectionTitles
    ].some((candidate) => normalizeLookup(candidate) === normalized)
  );
}

function normalizeDocumentProfile(value: unknown, fallbackProfile: string): DocumentProfile | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const record = value as Record<string, unknown>;
  const sectionRules = normalizeProfileSectionRules(record.sectionRules);
  if (!sectionRules.length) return undefined;
  return {
    version: 1,
    profile: readProfileString(record, "profile") || sanitizeProfileFileName(fallbackProfile) || "generic_document",
    title: readProfileString(record, "title") || undefined,
    titleSuffix: readProfileString(record, "titleSuffix") || undefined,
    globalRules: readProfileStringArray(record, "globalRules") ?? [],
    evidenceRules: readProfileStringArray(record, "evidenceRules") ?? [],
    sectionRules
  };
}

function normalizeProfileSectionRules(value: unknown): DocumentProfileSectionRule[] {
  if (!Array.isArray(value)) return [];
  const rules: DocumentProfileSectionRule[] = [];
  value.forEach((item, index) => {
    if (!item || typeof item !== "object" || Array.isArray(item)) return;
    const record = item as Record<string, unknown>;
    const match = readProfileString(record, "match");
    const title = readProfileString(record, "title");
    const id = readProfileString(record, "id") || (match || title ? `section_group_rule_${index + 1}` : "");
    if (!match && !title && !id) return;
    const rulesList = readProfileStringArray(record, "rules") ?? [];
    rules.push({
      id: id || undefined,
      match: match || undefined,
      title: title || undefined,
      order: readProfileNumber(record, "order", Number.NaN),
      rules: rulesList,
      writingRules: readProfileStringArray(record, "writingRules") ?? [],
      evidenceRules: readProfileStringArray(record, "evidenceRules") ?? [],
      requiredFacts: readProfileStringArray(record, "requiredFacts") ?? []
    });
  });
  return rules;
}

function applyProfileSectionRules(
  sectionGroups: DocumentConfigSectionGroup[],
  rules: DocumentProfileSectionRule[]
): DocumentConfigSectionGroup[] {
  if (!rules.length) return sectionGroups;
  return sectionGroups.map((group) => {
    const matchedRules = rules.filter((rule) => matchesProfileSectionRule(group, rule));
    if (!matchedRules.length) return group;
    const titleOverride = matchedRules.map((rule) => rule.title).find((title): title is string => Boolean(title));
    return {
      ...group,
      title: titleOverride || group.title,
      writingRules: uniqueStrings([
        ...group.writingRules,
        ...matchedRules.flatMap((rule) => [...(rule.rules ?? []), ...(rule.writingRules ?? [])])
      ]),
      evidenceRules: uniqueStrings([
        ...group.evidenceRules,
        ...matchedRules.flatMap((rule) => rule.evidenceRules ?? [])
      ]),
      requiredFacts: uniqueStrings([
        ...group.requiredFacts,
        ...matchedRules.flatMap((rule) => rule.requiredFacts ?? [])
      ])
    };
  });
}

function applyGlobalEvidenceRules(
  sectionGroups: DocumentConfigSectionGroup[],
  evidenceRules: string[]
): DocumentConfigSectionGroup[] {
  if (!evidenceRules.length) return sectionGroups;
  return sectionGroups.map((group) => ({
    ...group,
    evidenceRules: uniqueStrings([...group.evidenceRules, ...evidenceRules])
  }));
}

function applyDocumentSectionGroupPlans(
  sectionGroups: DocumentConfigSectionGroup[],
  plans: DocumentConfigSectionGroupPlanInput[]
): DocumentConfigSectionGroup[] {
  const normalizedPlans = plans.map(normalizeDocumentSectionGroupPlanInput).filter(isNonEmptyDocumentSectionGroupPlanInput);
  if (!normalizedPlans.length) return sectionGroups;
  return sectionGroups.map((group) => {
    const matchedPlans = normalizedPlans.filter((plan) => matchesDocumentSectionGroupPlan(group, plan));
    const plan = mergeDocumentSectionGroupPlans(matchedPlans);
    return plan ? { ...group, plan } : group;
  });
}

function normalizeDocumentPlanning(value?: Partial<DocumentGenerationPlanning>): DocumentGenerationPlanning | undefined {
  if (!value) return undefined;
  const summary = compactConfigText(typeof value.summary === "string" ? value.summary : "", 2000);
  const assumptions = uniqueStrings(Array.isArray(value.assumptions) ? value.assumptions : []);
  const risks = uniqueStrings(Array.isArray(value.risks) ? value.risks : []);
  if (!summary && !assumptions.length && !risks.length) return undefined;
  return {
    summary: summary || "按当前模板结构和已确认事实分章节组生成正文。",
    assumptions,
    risks
  };
}

function normalizeDocumentSectionGroupPlanInput(plan: DocumentConfigSectionGroupPlanInput): DocumentConfigSectionGroupPlanInput {
  return {
    id: normalizePlanText(plan.id),
    sectionGroupId: normalizePlanText(plan.sectionGroupId),
    match: normalizePlanText(plan.match),
    title: normalizePlanText(plan.title),
    order: Number.isFinite(plan.order) ? Math.trunc(plan.order ?? 0) : undefined,
    objective: compactConfigText(normalizePlanText(plan.objective), 1200),
    outline: uniqueStrings(plan.outline ?? []).map((item) => compactConfigText(item, 800)),
    keyPoints: uniqueStrings(plan.keyPoints ?? []).map((item) => compactConfigText(item, 800)),
    evidenceNeeds: uniqueStrings(plan.evidenceNeeds ?? []).map((item) => compactConfigText(item, 800)),
    openQuestions: uniqueStrings(plan.openQuestions ?? []).map((item) => compactConfigText(item, 800))
  };
}

function normalizePlanText(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

function isNonEmptyDocumentSectionGroupPlanInput(plan: DocumentConfigSectionGroupPlanInput): boolean {
  return Boolean(
    plan.id ||
      plan.sectionGroupId ||
      plan.match ||
      plan.title ||
      Number.isFinite(plan.order) ||
      plan.objective ||
      plan.outline?.length ||
      plan.keyPoints?.length ||
      plan.evidenceNeeds?.length ||
      plan.openQuestions?.length
  );
}

function matchesDocumentSectionGroupPlan(group: DocumentConfigSectionGroup, plan: DocumentConfigSectionGroupPlanInput): boolean {
  if (plan.id && normalizeLookup(plan.id) === normalizeLookup(group.id)) return true;
  if (plan.sectionGroupId && normalizeLookup(plan.sectionGroupId) === normalizeLookup(group.id)) return true;
  if (Number.isFinite(plan.order) && Math.trunc(plan.order ?? 0) === group.order) return true;
  const match = normalizeLookup(plan.match || plan.title || "");
  if (!match) return false;
  const candidates = [
    group.title,
    group.id,
    ...group.sectionIds,
    ...group.sectionNumbers,
    ...group.sectionTitles
  ].map(normalizeLookup);
  return candidates.some((candidate) => candidate === match || candidate.includes(match) || match.includes(candidate));
}

function mergeDocumentSectionGroupPlans(plans: DocumentConfigSectionGroupPlanInput[]): DocumentConfigSectionGroupPlan | undefined {
  if (!plans.length) return undefined;
  const objective = plans.map((plan) => plan.objective).find((value): value is string => Boolean(value));
  const outline = uniqueStrings(plans.flatMap((plan) => plan.outline ?? []));
  const keyPoints = uniqueStrings(plans.flatMap((plan) => plan.keyPoints ?? []));
  const evidenceNeeds = uniqueStrings(plans.flatMap((plan) => plan.evidenceNeeds ?? []));
  const openQuestions = uniqueStrings(plans.flatMap((plan) => plan.openQuestions ?? []));
  if (!objective && !outline.length && !keyPoints.length && !evidenceNeeds.length && !openQuestions.length) return undefined;
  return {
    objective: objective || "按本章节组结构形成可交付正文。",
    outline,
    keyPoints,
    evidenceNeeds,
    openQuestions
  };
}

function matchesProfileSectionRule(group: DocumentConfigSectionGroup, rule: DocumentProfileSectionRule): boolean {
  if (rule.id && normalizeLookup(rule.id) === normalizeLookup(group.id)) return true;
  if (Number.isFinite(rule.order) && Math.trunc(rule.order ?? 0) === group.order) return true;
  const match = normalizeLookup(rule.match || rule.title || "");
  if (!match) return false;
  const candidates = [
    group.title,
    group.id,
    ...group.sectionIds,
    ...group.sectionNumbers,
    ...group.sectionTitles
  ].map(normalizeLookup);
  return candidates.some((candidate) => candidate === match || candidate.includes(match) || match.includes(candidate));
}

function readProfileString(record: Record<string, unknown>, key: string): string {
  const value = record[key];
  return typeof value === "string" ? value.trim() : "";
}

function readProfileStringArray(record: Record<string, unknown>, key: string): string[] | undefined {
  const value = record[key];
  if (!Array.isArray(value)) return undefined;
  const items = value.map((item) => (typeof item === "string" ? item.trim() : "")).filter(Boolean);
  return items.length ? items : undefined;
}

function readProfileNumber(record: Record<string, unknown>, key: string, fallback: number): number {
  const value = record[key];
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string" && value.trim()) {
    const parsed = Number(value.trim());
    return Number.isFinite(parsed) ? parsed : fallback;
  }
  return fallback;
}

function sanitizeProfileFileName(value: string): string {
  return value.replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_").replace(/\s+/g, "_").trim();
}

function buildSectionGroupsFromSections(sections: DocumentConfigSectionLike[]): DocumentConfigSectionGroup[] {
  const grouped = new Map<string, DocumentConfigSectionLike[]>();
  for (const section of sections) {
    const root = getTopLevelSectionNumber(section.number);
    const items = grouped.get(root) ?? [];
    items.push(section);
    grouped.set(root, items);
  }

  return Array.from(grouped.entries())
    .sort((left, right) => Number(left[0]) - Number(right[0]))
    .map(([root, items], index) => {
      const ordered = [...items].sort((left, right) => compareSectionNumber(left.number, right.number));
      const rootSection = ordered.find((section) => section.number === root) ?? ordered[0];
      const sectionIds = ordered.map((section) => section.id);
      const sectionNumbers = ordered.map((section) => section.number);
      const sectionTitles = ordered.map((section) => section.title);
      const writingRules = uniqueStrings([
        `章节组主题：${rootSection?.title || `第 ${root} 章`}`,
        "只写本章节组对应的已确认事实，不能把模板提示、标准要求、示例资料或推断写成项目现状。",
        "没有证据的建设单位、系统名称、设备型号、密码产品、部署位置、接口和责任主体必须写待补充/需确认。",
        ...(ordered.flatMap((section) => section.writingHint ? [section.writingHint] : [])),
        ...(ordered.flatMap((section) => section.paragraphTasks ?? []).slice(0, 6))
      ]);
      const evidenceRules = uniqueStrings([
        "仅使用用户明确提供、附件工具读到或项目档案已确认的事实。",
        "项目档案中的待补充信息不得转写为事实。"
      ]);
      const requiredFacts = uniqueStrings([
        ...ordered.flatMap((section) => inferSectionRequiredFacts(section)),
        rootSection?.title ? `${rootSection.title}相关事实` : ""
      ]);

      return {
        id: `section_group_${root}`,
        order: Number(root) || index + 1,
        title: rootSection?.title || `第 ${root} 章节组`,
        sectionIds,
        sectionNumbers,
        sectionTitles,
        writingRules,
        evidenceRules,
        requiredFacts
      };
    });
}

function inferSectionRequiredFacts(section: DocumentConfigSectionLike): string[] {
  const text = `${section.title} ${section.writingHint || ""} ${section.placeholders?.join(" ") || ""}`;
  const facts: string[] = [];
  if (/系统|平台|应用/.test(text)) facts.push("应用系统名称和边界");
  if (/单位|建设/.test(text)) facts.push("建设单位与责任主体");
  if (/等级|等保/.test(text)) facts.push("等保级别或合规要求");
  if (/机房|云|部署|环境/.test(text)) facts.push("部署位置与环境");
  if (/产品|密码|算法|证书|密钥/.test(text)) facts.push("密码产品、算法和密钥管理事实");
  if (/接口|联通|访问|通信/.test(text)) facts.push("外部接口和访问路径");
  return facts;
}

function extractFactsFromMemory(text: string): DocumentConfigFact[] {
  const facts: DocumentConfigFact[] = [];
  const lines = text.split(/\r?\n/);
  let inFactBlock = false;
  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) continue;
    if (/^已确认事实[:：]$/.test(line)) {
      inFactBlock = true;
      continue;
    }
    if (/^待补充信息[:：]$/.test(line)) {
      inFactBlock = false;
      continue;
    }
    if (!inFactBlock) continue;
    const match = line.match(/^-?\s*([^：:]+)[:：]\s*(.+?)(?:\s*[（(]来源[:：]\s*([^)）]+)[)）])?\s*$/);
    if (!match) continue;
    facts.push({
      key: match[1].trim(),
      value: match[2].trim(),
      ...(match[3] ? { source: match[3].trim() } : {})
    });
  }
  return uniqueFacts(facts);
}

function extractGapsFromMemory(text: string): string[] {
  const gaps: string[] = [];
  const lines = text.split(/\r?\n/);
  let inGapBlock = false;
  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) continue;
    if (/^待补充信息[:：]$/.test(line)) {
      inGapBlock = true;
      continue;
    }
    if (inGapBlock && /^生成就绪[:：]/.test(line)) inGapBlock = false;
    if (inGapBlock && /^##\s+/.test(line)) inGapBlock = false;
    if (!inGapBlock) continue;
    const match = line.match(/^-?\s*(.+)$/);
    if (match) gaps.push(match[1].trim());
  }
  return uniqueStrings(gaps);
}

function uniqueFacts(facts: DocumentConfigFact[]): DocumentConfigFact[] {
  const seen = new Set<string>();
  const unique: DocumentConfigFact[] = [];
  for (const fact of facts) {
    const key = normalizeLookup(fact.key);
    const value = normalizeLookup(fact.value);
    const signature = `${key}:${value}`;
    if (seen.has(signature)) continue;
    seen.add(signature);
    unique.push(fact);
  }
  return unique;
}

function uniqueStrings(values: string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const value of values.map((item) => item.trim()).filter(Boolean)) {
    const key = normalizeLookup(value);
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(value);
  }
  return result;
}

function compareSectionNumber(left: string, right: string): number {
  const leftParts = left.split(".").map((part) => Number(part));
  const rightParts = right.split(".").map((part) => Number(part));
  const length = Math.max(leftParts.length, rightParts.length);
  for (let index = 0; index < length; index += 1) {
    const leftValue = leftParts[index] ?? -1;
    const rightValue = rightParts[index] ?? -1;
    if (leftValue !== rightValue) return leftValue - rightValue;
  }
  return left.localeCompare(right, "zh-CN");
}

function getTopLevelSectionNumber(value: string): string {
  return value.split(".")[0] || value;
}

function normalizeLookup(value: string): string {
  return value
    .toLowerCase()
    .replace(/[第章节\s.。．、,，:：()（）\[\]【】《》"'“”‘’_-]+/g, "")
    .trim();
}

function compactConfigText(value: string, maxChars: number): string {
  const normalized = value.replace(/\r\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  if (normalized.length <= maxChars) return normalized;
  return `${normalized.slice(0, maxChars)}\n\n[内容过长，已截断 ${normalized.length - maxChars} 字符]`;
}

function isDocumentConfig(value: unknown): value is DocumentConfig {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const record = value as Record<string, unknown>;
  return record.version === 1 && typeof record.profile === "string" && Array.isArray(record.sectionGroups);
}
