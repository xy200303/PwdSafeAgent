import { copyFile, mkdir, readFile, writeFile } from "node:fs/promises";
import { basename, dirname, extname } from "node:path";
import { createCanvas, loadImage } from "@napi-rs/canvas";
import { createReport } from "docx-templates";
import PizZip from "pizzip";
import { compactText, getCurrentTimeText, sanitizeFileName } from "./agentTools";

export interface SchemeDocumentInput {
  prompt: string;
  memory: string;
  generatedMarkdown: string;
  fields?: SchemeTemplateFieldInput;
  templateFields?: SchemeTemplateFieldInput;
  diagrams?: SchemeDiagramAsset[];
  renderMode?: SchemeDocumentRenderMode;
  templateJsonPath?: string;
  templateCells?: TemplateCellReplacementInput[];
  templateTables?: TemplateTableReplacementInput[];
  contentControls?: ContentControlReplacementInput[];
}

export type SchemeTemplateFieldInput = Record<string, unknown> | Array<Record<string, unknown>>;
export type SchemeDocumentRenderMode = "append" | "full_document" | "template_sections" | "template" | "section";

export interface SchemeDiagramAsset {
  label: string;
  kind?: string;
  path: string;
  figureId?: string;
}

export interface SchemeDocumentResult {
  outputPath: string;
  fileName: string;
  filledFields: string[];
  missingFields: string[];
  facts: SchemeFactModel;
  appendedMarkdown: boolean;
  embeddedDiagrams: string[];
  templateReplacementCount: number;
  templateTableReplacementCount: number;
  templateCellReplacementCount: number;
  contentControlReplacementCount: number;
  renderMode: SchemeDocumentRenderMode;
  templateAnchorsUsed: string[];
}

export interface TemplateWordDocumentInput {
  fields?: SchemeTemplateFieldInput;
  templateFields?: SchemeTemplateFieldInput;
  templateJsonPath?: string;
  templateCells?: TemplateCellReplacementInput[];
  templateTables?: TemplateTableReplacementInput[];
  contentControls?: ContentControlReplacementInput[];
}

export interface TemplateWordDocumentResult {
  outputPath: string;
  fileName: string;
  filledFields: string[];
  templateReplacementCount: number;
  templateTableReplacementCount: number;
  templateCellReplacementCount: number;
  contentControlReplacementCount: number;
}

export interface WordTemplateUpdateInput extends TemplateWordDocumentInput {
  diagrams?: SchemeDiagramAsset[];
}

export interface WordTemplateUpdateResult extends TemplateWordDocumentResult {
  embeddedDiagrams: string[];
}

export interface WordSectionReplacementInput extends TemplateWordDocumentInput {
  section: string;
  content: string;
  diagrams?: SchemeDiagramAsset[];
  templateJsonPath?: string;
}

export interface WordSectionBatchReplacementInput extends TemplateWordDocumentInput {
  sections: WordSectionContentInput[];
  diagrams?: SchemeDiagramAsset[];
  templateJsonPath?: string;
}

export interface WordSectionContentInput {
  section: string;
  content: string;
}

export interface WordSectionReplacementResult {
  outputPath: string;
  fileName: string;
  section: string;
  matchedHeading: string;
  replacementCount: number;
  templateReplacementCount: number;
  templateTableReplacementCount: number;
  templateCellReplacementCount: number;
  contentControlReplacementCount: number;
  embeddedDiagrams: string[];
  templateAnchorId?: string;
}

export interface WordSectionBatchReplacementResult {
  outputPath: string;
  fileName: string;
  sections: Array<{
    section: string;
    matchedHeading: string;
    replacementCount: number;
    templateAnchorId?: string;
  }>;
  templateReplacementCount: number;
  templateTableReplacementCount: number;
  templateCellReplacementCount: number;
  contentControlReplacementCount: number;
  embeddedDiagrams: string[];
}

export interface TemplateCellReplacementInput {
  tableId?: string;
  caption?: string;
  rowIndex: number;
  cellIndex?: number;
  columnIndex?: number;
  value: string;
}

export interface TemplateTableReplacementInput {
  tableId?: string;
  caption?: string;
  markdown?: string;
  rows?: string[][];
}

export interface ContentControlReplacementInput {
  tag: string;
  value: string;
}

export interface SchemeFactModel {
  systemName: string;
  organizationName: string;
  province: string;
  address: string;
  zipCode: string;
  securityLevel: string;
  machineRooms: Array<{
    name: string;
    owner: string;
    address: string;
  }>;
  subsystems: string[];
  keyDataTypes: string[];
  users: string[];
  deploymentMode: string;
  cloudPlatform: string;
  databases: string[];
  cryptoProducts: string[];
  missingFields: string[];
}

export interface WordTemplateJson {
  sections?: WordTemplateSection[];
  tables?: WordTemplateTable[];
  figures?: WordTemplateFigure[];
  fieldBlocks?: WordTemplateFieldBlock[];
  textBlocks?: WordTemplateTextBlock[];
}

export interface WordTemplateSection {
  id: string;
  number: string;
  title: string;
  headingBlock: number;
  bodyRange: [number, number];
  directBodyRange?: [number, number];
  headingLevel?: number;
  childSections?: string[];
  anchors?: {
    body?: WordTemplateAnchor;
  };
}

export interface WordTemplateTable {
  id: string;
  block: number;
  section?: string;
  sectionNumber?: string;
  caption?: string;
  captionBlock?: number;
  anchors?: {
    table?: WordTemplateAnchor;
    caption?: WordTemplateAnchor;
  };
  rows?: Array<{
    index: number;
    cells: Array<{
      cellIndex: number;
      columnIndex: number;
      text?: string;
    }>;
  }>;
}

export interface WordTemplateFigure {
  id: string;
  sectionNumber?: string;
  imageBlock?: number;
  captionBlock: number;
  caption: string;
  anchorKind?: string;
  mode?: string;
  anchors?: {
    image?: WordTemplateAnchor;
    caption?: WordTemplateAnchor;
  };
}

export interface WordTemplateAnchor {
  tag: string;
  alias?: string;
}

interface DynamicSdtElement {
  tag: string;
  alias?: string;
  xml: string;
  text: string;
  start: number;
  end: number;
}

const PLACEHOLDER_KEYS = [
  "应用系统",
  "建设单位",
  "单位省份",
  "单位地址",
  "单位邮编",
  "等保级别",
  "应用子系统 1",
  "应用子系统1",
  "应用子系统 2",
  "应用子系统2",
  "物理机房 1",
  "物理机房 1 管理单位",
  "物理机房 1 地址",
  "物理机房1",
  "物理机房1管理单位",
  "物理机房1地址",
  "物理机房2",
  "物理机房2管理单位",
  "物理机房2地址",
  "应用安全网关",
  "密码安全设备",
  "网络安全设备",
  "运维网关1",
  "运维网关2",
  "安全接入网关1",
  "安全接入网关2",
  "运维安全网关2",
  "安全认证网关",
  "服务器密码机",
  "签名验签服务器",
  "数字证书认证系统",
  "协同签名系统",
  "密码服务管理平台",
  "密码系统产品",
  "云平台",
  "网络安全产品",
  "密码安全产品",
  "编制日期"
] as const;

type PlaceholderKey = (typeof PLACEHOLDER_KEYS)[number];
type TemplateData = Record<string, string>;

export function extractSchemeFacts(input: SchemeDocumentInput): SchemeFactModel {
  const source = `${renderExplicitTemplateFieldSource(input)}\n\n${input.prompt}\n\n${input.memory}\n\n${input.generatedMarkdown}`;
  const appName = pickValue(source, [
    /(?:系统名称|项目名称|应用系统)\s*[：:]\s*([^\n，。；;]+)/,
    /([^\n，。；;]{2,40})(?:密码应用方案|系统密码应用方案)/
  ]);
  const orgName = pickValue(source, [/(?:项目建设单位|建设单位|单位名称)\s*[：:]\s*([^\n，。；;]+)/]);
  const province = pickValue(source, [/(?:单位省份|所属省份|所在省份|省份)\s*[：:]\s*([^\n，。；;]+)/]) || inferProvince(source);
  const address = pickValue(source, [/(?:单位地址|建设单位地址|地址)\s*[：:]\s*([^\n]+)/]);
  const zipCode = pickValue(source, [/(?:邮政编码|单位邮编|邮编)\s*[：:]\s*(\d{6})/]);
  const level = pickValue(source, [/(?:等保级别|安全保护等级|等级保护级别|密评依据).*?([一二三四1234])\s*级/]);
  const machineRoom = pickValue(source, [/(?:主机房|物理机房\s*1|机房名称)\s*[：:]\s*([^\n，。；;]+)/]);
  const machineRoomAddress = pickValue(source, [/(?:物理机房\s*1\s*地址|主机房地址|机房地址)\s*[：:]\s*([^\n，。；;]+)/]);
  const deploymentMode = pickValue(source, [/(?:部署模式|系统架构|架构模式)\s*[：:]\s*([^\n，。；;]+)/]) || inferDeploymentMode(source);
  const cloudPlatform = pickValue(source, [/(?:云平台名称|云平台)\s*[：:]\s*([^\n，。；;]+)/]);
  const facts: SchemeFactModel = {
    systemName: appName,
    organizationName: orgName,
    province,
    address,
    zipCode,
    securityLevel: normalizeLevel(level),
    machineRooms: [
      {
        name: machineRoom,
        owner: orgName,
        address: machineRoomAddress
      }
    ],
    subsystems: inferList(source, [
      /(?:子系统|应用模块|业务系统)\s*[：:]\s*([^\n]+)/,
      /主要包含[：:]?\s*([^\n。；;]+)/
    ]),
    keyDataTypes: inferKeywordList(source, [
      "身份鉴别数据",
      "重要业务数据",
      "重要审计数据",
      "个人信息",
      "电子证照数据",
      "交易数据",
      "日志数据",
      "密钥数据"
    ]),
    users: inferKeywordList(source, ["业务用户", "管理用户", "运维用户", "互联网用户", "第三方系统", "移动端用户"]),
    deploymentMode,
    cloudPlatform,
    databases: inferKeywordList(source, ["MySQL", "Oracle", "PostgreSQL", "SQL Server", "达梦", "人大金仓", "Redis"]),
    cryptoProducts: inferKeywordList(source, [
      "密码服务管理平台",
      "服务器密码机",
      "签名验签服务器",
      "安全认证网关",
      "数字证书认证系统",
      "协同签名系统",
      "时间戳服务器",
      "国密浏览器",
      "VPN"
    ]),
    missingFields: []
  };

  facts.missingFields = buildMissingFactList(facts);
  return facts;
}

export function buildSchemeTemplateData(input: SchemeDocumentInput): TemplateData {
  const facts = extractSchemeFacts(input);
  const appName = facts.systemName;
  const orgName = facts.organizationName;
  const province = facts.province;
  const address = facts.address;
  const zipCode = facts.zipCode;
  const level = facts.securityLevel;
  const machineRoom = facts.machineRooms[0]?.name ?? "";
  const machineRoomAddress = facts.machineRooms[0]?.address ?? "";

  const base: Record<PlaceholderKey, string> = {
    应用系统: appName || "待补充应用系统",
    建设单位: orgName || "待补充建设单位",
    单位省份: province || "待补充省份",
    单位地址: address || "待补充单位地址",
    单位邮编: zipCode || "待补充邮编",
    等保级别: level || "三级",
    "应用子系统 1": facts.subsystems[0] || (appName ? `${appName}业务子系统` : "待补充应用子系统1"),
    应用子系统1: facts.subsystems[0] || (appName ? `${appName}业务子系统` : "待补充应用子系统1"),
    "应用子系统 2": facts.subsystems[1] || (appName ? `${appName}管理子系统` : "待补充应用子系统2"),
    应用子系统2: facts.subsystems[1] || (appName ? `${appName}管理子系统` : "待补充应用子系统2"),
    "物理机房 1": machineRoom || "主机房",
    "物理机房 1 管理单位": orgName || "待补充机房管理单位",
    "物理机房 1 地址": machineRoomAddress || "待补充机房地址",
    物理机房1: machineRoom || "主机房",
    物理机房1管理单位: orgName || "待补充机房管理单位",
    物理机房1地址: machineRoomAddress || "待补充机房地址",
    物理机房2: "灾备机房",
    物理机房2管理单位: orgName || "待补充机房管理单位",
    物理机房2地址: "待补充灾备机房地址",
    应用安全网关: "安全认证网关",
    密码安全设备: "服务器密码机",
    网络安全设备: "防火墙",
    运维网关1: "运维安全网关",
    运维网关2: "运维安全网关",
    安全接入网关1: "安全接入网关",
    安全接入网关2: "安全接入网关",
    运维安全网关2: "运维安全网关",
    安全认证网关: "安全认证网关",
    服务器密码机: "服务器密码机",
    签名验签服务器: "签名验签服务器",
    数字证书认证系统: "数字证书认证系统",
    协同签名系统: "协同签名系统",
    密码服务管理平台: "密码服务管理平台",
    密码系统产品: facts.cryptoProducts.join("、") || "密码服务管理平台、服务器密码机、签名验签服务器",
    云平台: facts.cloudPlatform || "待补充云平台",
    网络安全产品: "防火墙、入侵检测系统、日志审计系统",
    密码安全产品: facts.cryptoProducts.join("、") || "密码服务管理平台、服务器密码机、签名验签服务器",
    编制日期: formatChineseDate(new Date())
  };

  const data: TemplateData = {};
  for (const key of PLACEHOLDER_KEYS) {
    addTemplateValue(data, key, base[key]);
  }
  for (const [key, value] of Object.entries(buildExplicitTemplateFieldOverrides(input))) {
    addTemplateValue(data, key, value);
  }

  data["方案生成正文"] = input.generatedMarkdown;
  data["结构化事实摘要"] = renderFactSummaryMarkdown(facts);
  data["当前时间"] = getCurrentTimeText();
  return data;
}

async function loadWordTemplateJson(templateJsonPath?: string, templateDocxPath?: string): Promise<WordTemplateJson | undefined> {
  if (templateJsonPath) {
    try {
      const parsed = JSON.parse(await readFile(templateJsonPath, "utf-8")) as WordTemplateJson;
      if (Array.isArray(parsed.sections)) return parsed;
    } catch {
      // Missing or malformed template JSON only disables anchor-based template operations.
    }
  }

  if (templateDocxPath) {
    try {
      return await buildWordTemplateJsonFromDocx(templateDocxPath);
    } catch {
      // Dynamic anchor parsing is best-effort; callers can still fall back to heading-based rendering.
    }
  }

  return undefined;
}

async function buildWordTemplateJsonFromDocx(templateDocxPath: string): Promise<WordTemplateJson | undefined> {
  const content = await readFile(templateDocxPath, "binary");
  const zip = new PizZip(content);
  return buildWordTemplateJsonFromZip(zip);
}

export async function parseWordTemplateAnchorsFromDocx(templateDocxPath: string): Promise<WordTemplateJson | undefined> {
  return buildWordTemplateJsonFromDocx(templateDocxPath);
}

function buildWordTemplateJsonFromZip(zip: PizZip): WordTemplateJson | undefined {
  const documentXml = zip.file("word/document.xml")?.asText();
  if (!documentXml) return undefined;

  const bodyOpen = documentXml.match(/<w:body\b[^>]*>/);
  const bodyEnd = documentXml.lastIndexOf("</w:body>");
  if (!bodyOpen || bodyOpen.index === undefined || bodyEnd < 0) return undefined;

  const bodyStart = bodyOpen.index + bodyOpen[0].length;
  const bodyXml = documentXml.slice(bodyStart, bodyEnd);
  const blocks = collectWordBodyBlocks(bodyXml, buildStyleHeadingLevels(zip));
  const sdtElements = collectDynamicSdtElements(bodyXml);
  const sections = buildDynamicTemplateSections(blocks);
  const tables = buildDynamicTemplateTables(sdtElements);
  const figures = buildDynamicTemplateFigures(sdtElements);
  const fieldBlocks = buildDynamicTemplateFieldBlocks(sdtElements);
  const textBlocks = buildDynamicTemplateTextBlocks(sdtElements);
  const templateJson: WordTemplateJson = {
    ...(sections.length ? { sections } : {}),
    ...(tables.length ? { tables } : {}),
    ...(figures.length ? { figures } : {}),
    ...(fieldBlocks.length ? { fieldBlocks } : {}),
    ...(textBlocks.length ? { textBlocks } : {})
  };

  return Object.keys(templateJson).length ? templateJson : undefined;
}

function collectDynamicSdtElements(xml: string, baseOffset = 0): DynamicSdtElement[] {
  const elements: DynamicSdtElement[] = [];
  let cursor = 0;

  while (cursor < xml.length) {
    const start = findNextElementStart(xml, "w:sdt", cursor);
    if (start < 0) break;

    const element = readBalancedWordElement(xml, start, "w:sdt");
    if (!element) break;
    const tag = extractSdtTag(element.xml);
    if (tag) {
      elements.push({
        tag,
        alias: extractSdtAlias(element.xml),
        xml: element.xml,
        text: extractVisibleWordText(element.xml),
        start: baseOffset + element.start,
        end: baseOffset + element.end
      });
    }

    const contentRange = getSdtContentRange(element.xml);
    if (contentRange) {
      elements.push(
        ...collectDynamicSdtElements(
          element.xml.slice(contentRange.start, contentRange.end),
          baseOffset + element.start + contentRange.start
        )
      );
    }
    cursor = element.end;
  }

  return elements;
}

function buildDynamicTemplateSections(blocks: WordBodyBlock[]): WordTemplateSection[] {
  const sections: WordTemplateSection[] = [];
  for (const block of blocks) {
    if (block.tagName !== "sdt" || !block.sdtTag) continue;
    const tag = block.sdtTag;
    const match = tag.match(/^ps:section:(sec_\d+(?:_\d+)*):body$/i);
    if (!match?.[1]) continue;
    const id = match[1];
    const number = parseSectionNumberFromTemplateId(id);
    const headingLevel = number.split(".").filter(Boolean).length || 1;
    const heading = findPreviousHeadingBlock(blocks, block.index, headingLevel);
    const title = heading ? stripHeadingNumberPrefix(heading.text) : block.text || number || id;
    sections.push({
      id,
      number,
      title,
      headingBlock: heading?.index ?? block.index,
      bodyRange: [block.index, block.index + 1],
      directBodyRange: [block.index, block.index + 1],
      headingLevel,
      anchors: {
        body: buildDynamicAnchor(tag)
      }
    });
  }

  for (const section of sections) {
    const childSections = sections
      .filter((candidate) => candidate.number.startsWith(`${section.number}.`) && candidate.number !== section.number)
      .map((candidate) => candidate.id);
    if (childSections.length) section.childSections = childSections;
  }

  return sections;
}

function buildDynamicTemplateTables(sdtElements: DynamicSdtElement[]): WordTemplateTable[] {
  const captionById = new Map<string, DynamicSdtElement>();
  const indexByTag = new Map<string, number>();
  sdtElements.forEach((element, index) => {
    indexByTag.set(element.tag, index);
    const captionMatch = element.tag.match(/^ps:table:([^:]+):caption$/i);
    if (captionMatch?.[1]) captionById.set(captionMatch[1], element);
  });

  const tables: WordTemplateTable[] = [];
  sdtElements.forEach((element, index) => {
    const tag = element.tag;
    const match = tag.match(/^ps:table:([^:]+)$/i);
    if (!match?.[1]) return;
    const id = match[1];
    const caption = captionById.get(id);
    const table: WordTemplateTable = {
      id,
      block: index,
      anchors: {
        table: buildDynamicAnchor(tag)
      }
    };
    if (caption) {
      table.caption = caption.text;
      table.captionBlock = indexByTag.get(caption.tag) ?? index;
      table.anchors = {
        ...table.anchors,
        caption: buildDynamicAnchor(caption.tag, caption.alias)
      };
    }
    const rows = buildDynamicTemplateTableRows(element.xml);
    if (rows) table.rows = rows;
    tables.push(table);
  });
  return tables;
}

function buildDynamicTemplateTableRows(xml: string): WordTemplateTable["rows"] {
  const tableXml = xml.match(/<w:tbl\b[\s\S]*?<\/w:tbl>/)?.[0] ?? "";
  if (!tableXml) return undefined;
  const rows = Array.from(tableXml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)).map((rowMatch, rowIndex) => {
    const cells = Array.from(rowMatch[0].matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)).map((cellMatch, cellIndex) => ({
      cellIndex,
      columnIndex: cellIndex,
      text: extractVisibleWordText(cellMatch[0])
    }));
    return { index: rowIndex, cells };
  });
  return rows.length ? rows : undefined;
}

function buildDynamicTemplateFigures(sdtElements: DynamicSdtElement[]): WordTemplateFigure[] {
  const figures = new Map<string, WordTemplateFigure>();
  for (const element of sdtElements) {
    const match = element.tag.match(/^ps:figure:([^:]+):(image|caption)$/i);
    if (!match?.[1] || !match[2]) continue;
    const id = match[1];
    const kind = match[2].toLowerCase();
    const figure = figures.get(id) ?? {
      id,
      captionBlock: 0,
      caption: "",
      anchors: {}
    };
    if (kind === "image") {
      figure.imageBlock = 0;
      figure.anchors = {
        ...figure.anchors,
        image: buildDynamicAnchor(element.tag, element.alias)
      };
    } else {
      figure.caption = element.text;
      figure.captionBlock = 0;
      figure.anchors = {
        ...figure.anchors,
        caption: buildDynamicAnchor(element.tag, element.alias)
      };
    }
    figures.set(id, figure);
  }
  return Array.from(figures.values());
}

function buildDynamicTemplateFieldBlocks(sdtElements: DynamicSdtElement[]): WordTemplateFieldBlock[] {
  const fieldBlocks: WordTemplateFieldBlock[] = [];
  sdtElements.forEach((element, index) => {
    const match = element.tag.match(/^ps:field-block:([^:]+)$/i);
    if (!match?.[1]) return;
    fieldBlocks.push({
      id: match[1],
      block: index,
      text: element.text,
      anchors: {
        block: buildDynamicAnchor(element.tag, element.alias)
      }
    });
  });
  return fieldBlocks;
}

function buildDynamicTemplateTextBlocks(sdtElements: DynamicSdtElement[]): WordTemplateTextBlock[] {
  const textBlocks: WordTemplateTextBlock[] = [];
  sdtElements.forEach((element, index) => {
    const match = element.tag.match(/^ps:section:(sec_\d+(?:_\d+)*):text:(\d+)$/i);
    if (!match?.[1] || !match[2]) return;
    const order = Number(match[2]);
    const section = match[1];
    textBlocks.push({
      id: `${section}_text_${order}`,
      section,
      sectionNumber: parseSectionNumberFromTemplateId(section),
      order,
      blockRange: [index, index + 1],
      text: element.text,
      anchors: {
        block: buildDynamicAnchor(element.tag, element.alias)
      }
    });
  });
  return textBlocks;
}

function buildDynamicAnchor(tag: string, alias?: string): WordTemplateAnchor {
  return alias ? { tag, alias } : { tag };
}

export async function createWordDocxFromTemplate(
  templatePath: string,
  outputPath: string,
  input: TemplateWordDocumentInput = {}
): Promise<TemplateWordDocumentResult> {
  await mkdir(dirname(outputPath), { recursive: true });
  const templateCells = input.templateCells ?? [];
  const templateTables = input.templateTables ?? [];
  const contentControls = input.contentControls ?? [];
  const explicitFields = buildExplicitTemplateFieldOverrides({
    prompt: "",
    memory: "",
    generatedMarkdown: "",
    fields: input.fields,
    templateFields: input.templateFields
  });
  const filledFields = Object.keys(explicitFields).filter((key) => !key.startsWith("$"));

  if (!filledFields.length && !templateCells.length && !templateTables.length && !contentControls.length) {
    await copyFile(templatePath, outputPath);
    return {
      outputPath,
      fileName: sanitizeFileName(outputPath.split(/[\\/]/).at(-1) || "密码应用方案.docx"),
      filledFields,
      templateReplacementCount: 0,
      templateTableReplacementCount: 0,
      templateCellReplacementCount: 0,
      contentControlReplacementCount: 0
    };
  }

  const content = await readFile(templatePath, "binary");
  const zip = new PizZip(content);
  const templateJson = await loadWordTemplateJson(input.templateJsonPath, templatePath);
  const templateReplacementCount = await replaceTemplatePlaceholders(zip, explicitFields);
  const templateTableReplacementCount = replaceTemplateTables(zip, templateJson, templateTables);
  const templateCellReplacementCount = replaceTemplateTableCells(zip, templateJson, templateCells);
  const contentControlReplacementCount = replaceContentControlsByTag(zip, contentControls, templateJson);
  const buffer = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
  await writeFile(outputPath, buffer);

  return {
    outputPath,
    fileName: sanitizeFileName(outputPath.split(/[\\/]/).at(-1) || "密码应用方案.docx"),
    filledFields,
    templateReplacementCount,
    templateTableReplacementCount,
    templateCellReplacementCount,
    contentControlReplacementCount
  };
}

export interface WordTemplateFieldBlock {
  id: string;
  block: number;
  text: string;
  placeholders?: string[];
  section?: string;
  sectionNumber?: string;
  anchors?: {
    block?: WordTemplateAnchor;
  };
}

export interface WordTemplateTextBlock {
  id: string;
  section: string;
  sectionNumber?: string;
  order: number;
  blockRange: [number, number];
  text?: string;
  anchors?: {
    block?: WordTemplateAnchor;
  };
}

export async function updateWordTemplateContent(
  sourcePath: string,
  outputPath: string,
  input: WordTemplateUpdateInput = {}
): Promise<WordTemplateUpdateResult> {
  const content = await readFile(sourcePath, "binary");
  const zip = new PizZip(content);
  const explicitFields = buildExplicitTemplateFieldOverrides({
    prompt: "",
    memory: "",
    generatedMarkdown: "",
    fields: input.fields,
    templateFields: input.templateFields
  });
  const templateJson = await loadWordTemplateJson(input.templateJsonPath, sourcePath);
  const templateReplacementCount = await replaceTemplatePlaceholders(zip, explicitFields);
  const templateTableReplacementCount = replaceTemplateTables(zip, templateJson, input.templateTables ?? []);
  const templateCellReplacementCount = replaceTemplateTableCells(zip, templateJson, input.templateCells ?? []);
  const contentControlReplacementCount = replaceContentControlsByTag(zip, input.contentControls ?? [], templateJson);
  const embeddedDiagrams = await appendGeneratedDiagrams(zip, input.diagrams ?? [], templateJson);

  await mkdir(dirname(outputPath), { recursive: true });
  const buffer = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
  await writeFile(outputPath, buffer);

  return {
    outputPath,
    fileName: sanitizeFileName(outputPath.split(/[\\/]/).at(-1) || "密码应用方案.docx"),
    filledFields: Object.keys(explicitFields).filter((key) => !key.startsWith("$")),
    templateReplacementCount,
    templateTableReplacementCount,
    templateCellReplacementCount,
    contentControlReplacementCount,
    embeddedDiagrams
  };
}

export async function replaceWordSectionContent(
  sourcePath: string,
  outputPath: string,
  input: WordSectionReplacementInput
): Promise<WordSectionReplacementResult> {
  const content = await readFile(sourcePath, "binary");
  const zip = new PizZip(content);
  const templateJson = await loadWordTemplateJson(input.templateJsonPath, sourcePath);
  const templateReplacementCount = await replaceTemplatePlaceholders(
    zip,
    buildExplicitTemplateFieldOverrides({
      prompt: "",
      memory: "",
      generatedMarkdown: "",
      fields: input.fields,
      templateFields: input.templateFields
    })
  );
  const replacement = replaceDocumentSectionWithMarkdown(zip, input.section, input.content, templateJson);
  if (!replacement) {
    throw new Error(`未找到 Word 章节：${input.section}`);
  }
  const templateTableReplacementCount = replaceTemplateTables(zip, templateJson, input.templateTables ?? []);
  const templateCellReplacementCount = replaceTemplateTableCells(zip, templateJson, input.templateCells ?? []);
  const contentControlReplacementCount = replaceContentControlsByTag(zip, input.contentControls ?? [], templateJson);
  const embeddedDiagrams = await appendGeneratedDiagrams(zip, input.diagrams ?? [], templateJson);

  await mkdir(dirname(outputPath), { recursive: true });
  const buffer = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
  await writeFile(outputPath, buffer);

  return {
    outputPath,
    fileName: sanitizeFileName(outputPath.split(/[\\/]/).at(-1) || "密码应用方案.docx"),
    section: input.section,
    matchedHeading: replacement.matchedHeading,
    replacementCount: replacement.replacementCount,
    templateReplacementCount,
    templateTableReplacementCount,
    templateCellReplacementCount,
    contentControlReplacementCount,
    embeddedDiagrams,
    templateAnchorId: replacement.templateAnchorId
  };
}

export async function replaceWordSectionsContent(
  sourcePath: string,
  outputPath: string,
  input: WordSectionBatchReplacementInput
): Promise<WordSectionBatchReplacementResult> {
  const sections = input.sections
    .map((section) => ({ section: section.section.trim(), content: section.content }))
    .filter((section) => section.section && section.content.trim());
  if (!sections.length) {
    throw new Error("未提供可写入的 Word 章节");
  }

  const content = await readFile(sourcePath, "binary");
  const zip = new PizZip(content);
  const templateJson = await loadWordTemplateJson(input.templateJsonPath, sourcePath);
  const templateReplacementCount = await replaceTemplatePlaceholders(
    zip,
    buildExplicitTemplateFieldOverrides({
      prompt: "",
      memory: "",
      generatedMarkdown: "",
      fields: input.fields,
      templateFields: input.templateFields
    })
  );
  const replacements = replaceDocumentSectionsWithMarkdown(zip, sections, templateJson);
  const templateTableReplacementCount = replaceTemplateTables(zip, templateJson, input.templateTables ?? []);
  const templateCellReplacementCount = replaceTemplateTableCells(zip, templateJson, input.templateCells ?? []);
  const contentControlReplacementCount = replaceContentControlsByTag(zip, input.contentControls ?? [], templateJson);

  const embeddedDiagrams = await appendGeneratedDiagrams(zip, input.diagrams ?? [], templateJson);

  await mkdir(dirname(outputPath), { recursive: true });
  const buffer = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
  await writeFile(outputPath, buffer);

  return {
    outputPath,
    fileName: sanitizeFileName(outputPath.split(/[\\/]/).at(-1) || "密码应用方案.docx"),
    sections: replacements,
    templateReplacementCount,
    templateTableReplacementCount,
    templateCellReplacementCount,
    contentControlReplacementCount,
    embeddedDiagrams
  };
}

export async function writeSchemeDocxFromTemplate(
  templatePath: string,
  outputPath: string,
  input: SchemeDocumentInput
): Promise<SchemeDocumentResult> {
  const facts = extractSchemeFacts(input);
  const data = buildSchemeTemplateData(input);
  const content = await readFile(templatePath, "binary");
  const zip = new PizZip(content);
  const templateJson = await loadWordTemplateJson(input.templateJsonPath, templatePath);
  const templateReplacementCount = await replaceTemplatePlaceholders(zip, data);
  const renderedZip = zip;
  const renderMode = input.renderMode ?? "append";
  let appendedMarkdown = false;
  let templateAnchorsUsed: string[] = [];
  if (renderMode === "full_document") {
    if (shouldPreserveTemplateAnchorsForFullDocument(input, templateJson)) {
      const replacement = replaceMarkdownDocumentSections(renderedZip, input.generatedMarkdown, templateJson);
      appendedMarkdown = replacement.replacementCount > 0;
      templateAnchorsUsed = replacement.templateAnchorIds;
      if (!appendedMarkdown) {
        throw new Error(
          "render_mode=full_document 在当前写入中需要保留模板图位或表格锚点，但 Markdown 未匹配到模板章节。请改用带编号的章节正文，或直接使用 sections/template_sections 写入。"
        );
      }
    } else {
      appendedMarkdown = replaceDocumentBodyWithGeneratedMarkdown(renderedZip, input.generatedMarkdown, facts);
    }
  } else if (renderMode === "template_sections") {
    const replacement = replaceMarkdownDocumentSections(renderedZip, input.generatedMarkdown, templateJson);
    appendedMarkdown = replacement.replacementCount > 0;
    templateAnchorsUsed = replacement.templateAnchorIds;
    if (!appendedMarkdown) {
      appendedMarkdown = appendGeneratedMarkdown(renderedZip, input.generatedMarkdown, facts);
    }
  } else {
    appendedMarkdown = appendGeneratedMarkdown(renderedZip, input.generatedMarkdown, facts);
  }
  const templateTableReplacementCount = replaceTemplateTables(renderedZip, templateJson, input.templateTables ?? []);
  const templateCellReplacementCount = replaceTemplateTableCells(renderedZip, templateJson, input.templateCells ?? []);
  const contentControlReplacementCount = replaceContentControlsByTag(renderedZip, input.contentControls ?? [], templateJson);
  const embeddedDiagrams = await appendGeneratedDiagrams(renderedZip, input.diagrams ?? [], templateJson);
  replaceResidualTemplateMarkers(renderedZip);
  await mkdir(dirname(outputPath), { recursive: true });
  const buffer = renderedZip.generate({ type: "nodebuffer", compression: "DEFLATE" });
  await writeFile(outputPath, buffer);

  const filledFields = PLACEHOLDER_KEYS.filter((key) => !data[key].startsWith("待补充"));
  const missingFields = PLACEHOLDER_KEYS.filter((key) => data[key].startsWith("待补充"));
  return {
    outputPath,
    fileName: sanitizeFileName(outputPath.split(/[\\/]/).at(-1) || "密码应用方案.docx"),
    filledFields,
    missingFields,
    facts,
    appendedMarkdown,
    embeddedDiagrams,
    templateReplacementCount,
    templateTableReplacementCount,
    templateCellReplacementCount,
    contentControlReplacementCount,
    renderMode,
    templateAnchorsUsed
  };
}

function shouldPreserveTemplateAnchorsForFullDocument(
  input: Pick<SchemeDocumentInput, "diagrams" | "templateCells" | "templateTables">,
  templateJson?: WordTemplateJson
): boolean {
  if (!templateJson?.sections?.length) return false;
  return Boolean((input.diagrams?.length ?? 0) || (input.templateTables?.length ?? 0) || (input.templateCells?.length ?? 0));
}

export async function replaceTemplatePlaceholders(zip: PizZip, data: TemplateData): Promise<number> {
  const docxTemplateData = buildDocxTemplatesData(data);
  return Object.keys(docxTemplateData).length ? await renderDocxTemplateFields(zip, docxTemplateData) : 0;
}

function replaceResidualTemplateMarkers(zip: PizZip): number {
  const replacements: PlaceholderReplacement[] = [
    { placeholder: "【正文占位】", value: "本节内容待补充/需确认。" },
    { placeholder: "【图片占位】", value: "图示待生成/需确认。" },
    { placeholder: "【待填写】", value: "待补充/需确认" }
  ];
  let count = 0;
  for (const fileName of getTemplateXmlFileNames(zip)) {
    const file = zip.file(fileName);
    const xml = file?.asText();
    if (!xml) continue;
    const result = replaceTextNodePlaceholders(xml, replacements);
    if (result.count > 0) {
      zip.file(fileName, result.xml);
      count += result.count;
    }
  }
  return count;
}

function buildExplicitTemplateFieldOverrides(input: SchemeDocumentInput): TemplateData {
  const overrides: TemplateData = {};
  applyTemplateFieldInput(overrides, input.fields);
  applyTemplateFieldInput(overrides, input.templateFields);
  return overrides;
}

function renderExplicitTemplateFieldSource(input: SchemeDocumentInput): string {
  const fields = buildExplicitTemplateFieldOverrides(input);
  return Object.entries(fields)
    .filter(([key]) => !key.startsWith("$"))
    .map(([key, value]) => `${key}：${value}`)
    .join("\n");
}

function applyTemplateFieldInput(target: TemplateData, input: SchemeTemplateFieldInput | undefined): void {
  if (!input) return;

  if (Array.isArray(input)) {
    for (const item of input) {
      if (!item || typeof item !== "object" || Array.isArray(item)) continue;
      const record = item as Record<string, unknown>;
      applyTemplateField(target, record.key, record.value);
    }
    return;
  }

  if (typeof input !== "object") return;
  applyStructuredTemplateFieldGroups(target, input);
  for (const [key, value] of Object.entries(input)) {
    applyTemplateField(target, key, value);
  }
}

function applyStructuredTemplateFieldGroups(target: TemplateData, input: Record<string, unknown>): void {
  const subsystems = readStringArrayField(input, ["subsystems", "applicationSubsystems", "application_subsystems"]);
  if (subsystems[0]) applyTemplateField(target, "应用子系统1", subsystems[0]);
  if (subsystems[1]) applyTemplateField(target, "应用子系统2", subsystems[1]);

  const cryptoProducts = readStringArrayField(input, ["cryptoProducts", "crypto_products", "passwordProducts", "password_products"]);
  if (cryptoProducts.length) {
    const value = cryptoProducts.join("、");
    applyTemplateField(target, "密码系统产品", value);
    applyTemplateField(target, "密码安全产品", value);
  }

  const networkProducts = readStringArrayField(input, ["networkSecurityProducts", "network_security_products"]);
  if (networkProducts.length) applyTemplateField(target, "网络安全产品", networkProducts.join("、"));

  const machineRooms = input.machineRooms ?? input.machine_rooms;
  if (Array.isArray(machineRooms)) {
    for (const [index, room] of machineRooms.slice(0, 2).entries()) {
      if (!room || typeof room !== "object" || Array.isArray(room)) continue;
      const record = room as Record<string, unknown>;
      const roomNumber = index + 1;
      applyTemplateField(target, `物理机房${roomNumber}`, record.name);
      applyTemplateField(target, `物理机房${roomNumber}管理单位`, record.owner ?? record.manager ?? record.organization);
      applyTemplateField(target, `物理机房${roomNumber}地址`, record.address);
    }
  }
}

function readStringArrayField(input: Record<string, unknown>, keys: string[]): string[] {
  for (const key of keys) {
    const value = input[key];
    if (!Array.isArray(value)) continue;
    return value.map(stringifyTemplateFieldValue).filter(Boolean);
  }
  return [];
}

function applyTemplateField(target: TemplateData, rawKey: unknown, rawValue: unknown): void {
  if (typeof rawKey !== "string") return;
  const key = normalizeTemplateFieldKey(rawKey);
  const value = stringifyTemplateFieldValue(rawValue);
  if (!key || !value) return;
  addTemplateValue(target, key, value);
}

function normalizeTemplateFieldKey(rawKey: string): string {
  const stripped = stripTemplateDelimiters(rawKey);
  const aliasKey = stripped.replace(/[\s_\-]+/g, "").toLowerCase();
  return TEMPLATE_FIELD_ALIASES[aliasKey] ?? stripped;
}

function stripTemplateDelimiters(rawKey: string): string {
  const key = rawKey.trim();
  if (key.startsWith("{") && key.endsWith("}") && !key.startsWith("${")) {
    return key.slice(1, -1).trim();
  }
  return key;
}

function stringifyTemplateFieldValue(value: unknown): string {
  if (typeof value === "string") return cleanValue(value);
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (Array.isArray(value)) return value.map(stringifyTemplateFieldValue).filter(Boolean).join("、");
  if (!value || typeof value !== "object") return "";

  const record = value as Record<string, unknown>;
  return (
    stringifyTemplateFieldValue(record.value) ||
    stringifyTemplateFieldValue(record.name) ||
    stringifyTemplateFieldValue(record.label) ||
    stringifyTemplateFieldValue(record.title)
  );
}

const TEMPLATE_FIELD_ALIASES: Record<string, PlaceholderKey> = {
  applicationsystem: "应用系统",
  systemname: "应用系统",
  projectname: "应用系统",
  constructionunit: "建设单位",
  organizationname: "建设单位",
  organisationname: "建设单位",
  orgname: "建设单位",
  unitname: "建设单位",
  province: "单位省份",
  unitprovince: "单位省份",
  address: "单位地址",
  unitaddress: "单位地址",
  zipcode: "单位邮编",
  postalcode: "单位邮编",
  securitylevel: "等保级别",
  protectionlevel: "等保级别",
  subsystem1: "应用子系统1",
  applicationsubsystem1: "应用子系统1",
  subsystem2: "应用子系统2",
  applicationsubsystem2: "应用子系统2",
  machineroom1: "物理机房1",
  machineroom1owner: "物理机房1管理单位",
  machineroom1manager: "物理机房1管理单位",
  machineroom1address: "物理机房1地址",
  machineroom2: "物理机房2",
  machineroom2owner: "物理机房2管理单位",
  machineroom2manager: "物理机房2管理单位",
  machineroom2address: "物理机房2地址",
  cloudplatform: "云平台",
  networksecurityproducts: "网络安全产品",
  cryptosecurityproducts: "密码安全产品",
  cryptoproducts: "密码系统产品",
  cryptosystemproducts: "密码系统产品",
  passwordproducts: "密码系统产品",
  compiledate: "编制日期",
  compilationdate: "编制日期"
};

function addTemplateValue(data: TemplateData, key: string, value: string): void {
  data[key] = value;
  data[key.replace(/\s+/g, "")] = value;
}

interface PlaceholderReplacement {
  placeholder: string;
  value: string;
}

interface TextNodeToken {
  start: number;
  end: number;
  openTag: string;
  closeTag: string;
  text: string;
  textStart: number;
  textEnd: number;
}

function buildDocxTemplatesData(data: TemplateData): TemplateData {
  const fields: TemplateData = {};
  for (const [key, value] of Object.entries(data)) {
    if (!isRenderableTemplateFieldKey(key)) continue;
    fields[key] = value;
    const compactKey = key.replace(/\s+/g, "");
    if (compactKey !== key) fields[compactKey] = value;
  }
  return fields;
}

async function renderDocxTemplateFields(zip: PizZip, data: TemplateData): Promise<number> {
  const before = countDocxTemplateFieldOccurrences(zip, data);
  if (!before) return 0;

  const preparedZip = new PizZip(zip.generate({ type: "nodebuffer", compression: "DEFLATE" }));
  prepareDocxTemplateCommands(preparedZip, data);
  const rendered = await createReport({
    template: preparedZip.generate({ type: "nodebuffer", compression: "DEFLATE" }),
    data: { data },
    cmdDelimiter: ["{", "}"],
    noSandbox: false,
    failFast: true,
    rejectNullish: false,
    processLineBreaks: true,
    errorHandler: (_error, rawCode) => `{${rawCode ?? ""}}`
  });
  replaceZipContents(zip, new PizZip(rendered));
  return before - countDocxTemplateFieldOccurrences(zip, data);
}

function prepareDocxTemplateCommands(zip: PizZip, data: TemplateData): void {
  const commandReplacements = buildDocxTemplateCommandReplacements(data);
  if (!commandReplacements.length) return;
  for (const fileName of getTemplateXmlFileNames(zip)) {
    const file = zip.file(fileName);
    const xml = file?.asText();
    if (!xml) continue;

    const result = replaceTextNodePlaceholders(xml, commandReplacements);
    if (result.count > 0) zip.file(fileName, result.xml);
  }
}

function buildDocxTemplateCommandReplacements(data: TemplateData): PlaceholderReplacement[] {
  const replacements = new Map<string, string>();
  for (const key of Object.keys(data)) {
    if (!isRenderableTemplateFieldKey(key)) continue;
    replacements.set(`{${key}}`, `{INS data[${JSON.stringify(key)}]}`);
  }
  return Array.from(replacements, ([placeholder, value]) => ({ placeholder, value })).sort(
    (left, right) => right.placeholder.length - left.placeholder.length
  );
}

function countDocxTemplateFieldOccurrences(zip: PizZip, data: TemplateData): number {
  let count = 0;
  for (const fileName of getTemplateXmlFileNames(zip)) {
    const text = extractTextFromXml(zip.file(fileName)?.asText() ?? "");
    if (!text) continue;
    for (const key of Object.keys(data)) {
      count += countOccurrences(text, `{${key}}`);
    }
  }
  return count;
}

function replaceZipContents(target: PizZip, source: PizZip): void {
  for (const fileName of Object.keys(target.files)) {
    target.remove(fileName);
  }

  for (const [fileName, file] of Object.entries(source.files)) {
    if (file.dir) {
      target.folder(fileName.replace(/\/$/, ""));
      continue;
    }
    target.file(fileName, file.asUint8Array(), { binary: true });
  }
}

function countOccurrences(source: string, needle: string): number {
  if (!source || !needle) return 0;
  let count = 0;
  let index = source.indexOf(needle);
  while (index >= 0) {
    count += 1;
    index = source.indexOf(needle, index + needle.length);
  }
  return count;
}

function isRenderableTemplateFieldKey(key: string): boolean {
  return Boolean(key) && !key.startsWith("$") && key !== "方案生成正文" && key !== "结构化事实摘要" && key !== "当前时间";
}

function getTemplateXmlFileNames(zip: PizZip): string[] {
  return Object.keys(zip.files).filter((fileName) =>
    /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes|comments).*\.xml$/i.test(fileName)
  );
}

function replaceTextNodePlaceholders(
  xml: string,
  replacements: PlaceholderReplacement[]
): { xml: string; count: number } {
  const tokens = collectTextNodeTokens(xml);
  if (!tokens.length) return { xml, count: 0 };

  const fullText = tokens.map((token) => token.text).join("");
  const matches = selectNonOverlappingMatches(fullText, replacements);
  if (!matches.length) return { xml, count: 0 };

  for (const match of matches.sort((left, right) => right.start - left.start)) {
    applyPlaceholderMatch(tokens, match);
  }

  let nextXml = xml;
  for (const token of [...tokens].reverse()) {
    const openTag = ensureTextNodePreservesSpaces(token.openTag, token.text);
    nextXml = `${nextXml.slice(0, token.start)}${openTag}${escapeXml(token.text)}${token.closeTag}${nextXml.slice(token.end)}`;
  }

  return { xml: nextXml, count: matches.length };
}

function collectTextNodeTokens(xml: string): TextNodeToken[] {
  const tokens: TextNodeToken[] = [];
  const pattern = /(<((?:w|a|m):t|w:instrText)\b[^>]*>)([\s\S]*?)(<\/\2>)/g;
  let textOffset = 0;
  let match: RegExpExecArray | null;

  while ((match = pattern.exec(xml))) {
    const text = decodeXmlText(match[3]);
    tokens.push({
      start: match.index,
      end: match.index + match[0].length,
      openTag: match[1],
      closeTag: match[4],
      text,
      textStart: textOffset,
      textEnd: textOffset + text.length
    });
    textOffset += text.length;
  }

  return tokens;
}

function selectNonOverlappingMatches(fullText: string, replacements: PlaceholderReplacement[]): PlaceholderMatch[] {
  const candidates: PlaceholderMatch[] = [];
  for (const replacement of replacements) {
    let index = fullText.indexOf(replacement.placeholder);
    while (index >= 0) {
      candidates.push({
        start: index,
        end: index + replacement.placeholder.length,
        value: replacement.value,
        placeholder: replacement.placeholder
      });
      index = fullText.indexOf(replacement.placeholder, index + 1);
    }
  }

  const selected: PlaceholderMatch[] = [];
  let lastEnd = -1;
  for (const candidate of candidates.sort(comparePlaceholderMatches)) {
    if (candidate.start < lastEnd) continue;
    selected.push(candidate);
    lastEnd = candidate.end;
  }
  return selected;
}

interface PlaceholderMatch {
  start: number;
  end: number;
  value: string;
  placeholder: string;
}

function comparePlaceholderMatches(left: PlaceholderMatch, right: PlaceholderMatch): number {
  if (left.start !== right.start) return left.start - right.start;
  return right.placeholder.length - left.placeholder.length;
}

function applyPlaceholderMatch(tokens: TextNodeToken[], match: PlaceholderMatch): void {
  const startToken = findTextTokenAt(tokens, match.start);
  const endToken = findTextTokenAt(tokens, match.end - 1);
  if (!startToken || !endToken) return;

  const startOffset = match.start - startToken.textStart;
  const endOffset = match.end - endToken.textStart;
  const startIndex = tokens.indexOf(startToken);
  const endIndex = tokens.indexOf(endToken);

  if (startIndex === endIndex) {
    startToken.text = `${startToken.text.slice(0, startOffset)}${match.value}${startToken.text.slice(endOffset)}`;
    return;
  }

  startToken.text = `${startToken.text.slice(0, startOffset)}${match.value}`;
  for (let index = startIndex + 1; index < endIndex; index += 1) {
    tokens[index].text = "";
  }
  endToken.text = endToken.text.slice(endOffset);
}

function findTextTokenAt(tokens: TextNodeToken[], offset: number): TextNodeToken | undefined {
  return tokens.find((token) => offset >= token.textStart && offset < token.textEnd);
}

function ensureTextNodePreservesSpaces(openTag: string, text: string): string {
  if (!/^\s|\s$/.test(text) || /\sxml:space=/.test(openTag)) return openTag;
  return openTag.replace(/>$/, ' xml:space="preserve">');
}

interface WordBodyBlock {
  index: number;
  start: number;
  end: number;
  tagName: string;
  xml: string;
  text: string;
  styleId: string;
  sdtTag?: string;
  headingLevel?: number;
  headingNumber?: string;
}

interface SectionReplacementSummary {
  matchedHeading: string;
  replacementCount: number;
  templateAnchorId?: string;
}

interface DocumentSectionsReplacementSummary {
  matchedHeadings: string[];
  replacementCount: number;
  templateAnchorIds: string[];
}

interface MarkdownSectionSegment {
  number: string;
  title: string;
  level: number;
  body: string;
  hasChildren: boolean;
}

interface SectionReplacementAnchor {
  target: WordBodyBlock;
  replacementStartIndex: number;
  replacementEndIndex: number;
  insertStartOffset?: number;
  insertEndOffset?: number;
  sdtBlock?: WordBodyBlock;
  templateSection?: WordTemplateSection;
  contentControlTag?: string;
  candidateContentControlTags?: string[];
}

function replaceDocumentSectionWithMarkdown(
  zip: PizZip,
  section: string,
  markdown: string,
  templateJson?: WordTemplateJson
): SectionReplacementSummary | undefined {
  const compactedMarkdown = compactText(markdown, 120000);
  if (!compactedMarkdown.trim()) return undefined;

  const documentFile = zip.file("word/document.xml");
  const documentXml = documentFile?.asText();
  if (!documentXml) return undefined;

  const bodyOpen = documentXml.match(/<w:body\b[^>]*>/);
  const bodyEnd = documentXml.lastIndexOf("</w:body>");
  if (!bodyOpen || bodyOpen.index === undefined || bodyEnd < 0) return undefined;

  const bodyStart = bodyOpen.index + bodyOpen[0].length;
  const bodyXml = documentXml.slice(bodyStart, bodyEnd);
  const preciseReplacements = replaceNumberedMarkdownSections(zip, section, markdown, templateJson);
  if (preciseReplacements.length) {
    return {
      matchedHeading: preciseReplacements.map((replacement) => replacement.matchedHeading).join("、"),
      replacementCount: preciseReplacements.reduce((total, replacement) => total + replacement.replacementCount, 0),
      templateAnchorId: preciseReplacements.map((replacement) => replacement.templateAnchorId).find(Boolean)
    };
  }

  const blocks = collectWordBodyBlocks(bodyXml, buildStyleHeadingLevels(zip));
  const anchor = findSectionReplacementAnchor(blocks, section, templateJson);
  if (!anchor) return undefined;

  const insertStart = anchor.insertStartOffset ?? blocks[anchor.replacementStartIndex]?.start ?? anchor.target.end;
  const insertEnd = anchor.insertEndOffset ?? blocks[anchor.replacementEndIndex]?.start ?? bodyXml.length;
  const markdownBody = stripLeadingMatchingMarkdownHeading(compactedMarkdown, section, anchor.target.text);
  const replacementXml = buildSectionReplacementXml(
    markdownBody,
    blocks,
    anchor.target,
    anchor.sdtBlock,
    anchor.templateSection
  );
  const anchoredReplacementXml = anchor.sdtBlock ? replaceSdtContentXml(anchor.sdtBlock.xml, replacementXml) : replacementXml;
  const nextBodyXml = `${bodyXml.slice(0, insertStart)}${anchoredReplacementXml}${bodyXml.slice(insertEnd)}`;
  const nextXml = `${documentXml.slice(0, bodyStart)}${nextBodyXml}${documentXml.slice(bodyEnd)}`;
  zip.file("word/document.xml", nextXml);

  return {
    matchedHeading: formatMatchedHeading(anchor.target),
    replacementCount: Math.max(0, anchor.replacementEndIndex - anchor.replacementStartIndex),
    templateAnchorId: anchor.templateSection?.id
  };
}

function replaceNumberedMarkdownSections(
  zip: PizZip,
  section: string,
  markdown: string,
  templateJson?: WordTemplateJson
): SectionReplacementSummary[] {
  const segments = selectPreciseMarkdownSectionSegments(section, extractMarkdownSectionSegments(markdown));
  const replacements: SectionReplacementSummary[] = [];

  for (const segment of segments.sort(compareSectionNumbersDescending)) {
    const replacement = replaceSingleDocumentSectionWithMarkdown(zip, segment.number, segment.body, templateJson);
    if (replacement) replacements.push(replacement);
  }

  return replacements.reverse();
}

function replaceMarkdownDocumentSections(
  zip: PizZip,
  markdown: string,
  templateJson?: WordTemplateJson
): DocumentSectionsReplacementSummary {
  const templateSectionReplacements = buildTemplateSectionMarkdownReplacements(markdown, templateJson);
  if (templateSectionReplacements.length) {
    try {
      const replacements = replaceDocumentSectionsWithMarkdown(zip, templateSectionReplacements, templateJson);
      if (replacements.length) {
        return {
          matchedHeadings: replacements.map((replacement) => replacement.matchedHeading),
          replacementCount: replacements.reduce((total, replacement) => total + replacement.replacementCount, 0),
          templateAnchorIds: replacements
            .map((replacement) => replacement.templateAnchorId)
            .filter((id): id is string => Boolean(id))
        };
      }
    } catch {
      // Fall back to the older leaf-section replacement path for templates whose JSON no longer matches the docx anchors.
    }
  }

  const segments = extractMarkdownSectionSegments(markdown).filter((segment) => segment.body.trim() && !segment.hasChildren);
  const replacements: SectionReplacementSummary[] = [];

  for (const segment of segments.sort(compareSectionNumbersDescending)) {
    const replacement = replaceSingleDocumentSectionWithMarkdown(zip, segment.number, segment.body, templateJson);
    if (replacement) replacements.push(replacement);
  }

  const orderedReplacements = replacements.reverse();
  return {
    matchedHeadings: orderedReplacements.map((replacement) => replacement.matchedHeading),
    replacementCount: orderedReplacements.reduce((total, replacement) => total + replacement.replacementCount, 0),
    templateAnchorIds: orderedReplacements
      .map((replacement) => replacement.templateAnchorId)
      .filter((id): id is string => Boolean(id))
  };
}

function buildTemplateSectionMarkdownReplacements(
  markdown: string,
  templateJson?: WordTemplateJson
): WordSectionContentInput[] {
  const templateSections = templateJson?.sections ?? [];
  if (!templateSections.length) return [];

  const segments = extractMarkdownSectionSegments(markdown);
  if (!segments.length) return [];

  const segmentByNumber = new Map<string, MarkdownSectionSegment>();
  const segmentsByTitle = new Map<string, MarkdownSectionSegment[]>();
  for (const segment of segments) {
    if (segment.number && !segmentByNumber.has(segment.number)) {
      segmentByNumber.set(segment.number, segment);
    }
    const normalizedTitle = normalizeHeadingLookup(segment.title);
    if (normalizedTitle) {
      segmentsByTitle.set(normalizedTitle, [...(segmentsByTitle.get(normalizedTitle) ?? []), segment]);
    }
  }

  let matchedCount = 0;
  const replacements = templateSections.map((section) => {
    const segment =
      segmentByNumber.get(section.number) ??
      selectUniqueMarkdownSegmentByTitle(segmentsByTitle, section.title) ??
      selectUniqueMarkdownSegmentByTitle(segmentsByTitle, `${section.number} ${section.title}`);
    if (segment) matchedCount += 1;
    const content = segment?.body.trim() || "本节内容待补充/需确认。";
    return {
      section: section.id,
      content
    };
  });

  return matchedCount > 0 ? replacements : [];
}

function selectUniqueMarkdownSegmentByTitle(
  segmentsByTitle: Map<string, MarkdownSectionSegment[]>,
  title: string
): MarkdownSectionSegment | undefined {
  const matches = segmentsByTitle.get(normalizeHeadingLookup(title));
  return matches?.length === 1 ? matches[0] : undefined;
}

function replaceDocumentSectionsWithMarkdown(
  zip: PizZip,
  sections: WordSectionContentInput[],
  templateJson?: WordTemplateJson
): WordSectionBatchReplacementResult["sections"] {
  const documentFile = zip.file("word/document.xml");
  const documentXml = documentFile?.asText();
  if (!documentXml) return [];

  const bodyOpen = documentXml.match(/<w:body\b[^>]*>/);
  const bodyEnd = documentXml.lastIndexOf("</w:body>");
  if (!bodyOpen || bodyOpen.index === undefined || bodyEnd < 0) return [];

  const bodyStart = bodyOpen.index + bodyOpen[0].length;
  const bodyXml = documentXml.slice(bodyStart, bodyEnd);
  const blocks = collectWordBodyBlocks(bodyXml, buildStyleHeadingLevels(zip));
  const xmlReplacements: Array<{
    section: string;
    start: number;
    end: number;
    xml: string;
    summary: WordSectionBatchReplacementResult["sections"][number];
  }> = [];

  for (const item of sections) {
    const compactedMarkdown = compactText(item.content, 120000);
    if (!compactedMarkdown.trim()) continue;

    const anchor = findSectionReplacementAnchor(blocks, item.section, templateJson);
    if (!anchor) throw new Error(`未找到 Word 章节锚点：${item.section}${getTemplateSectionAnchorTagHint(templateJson, item.section)}`);

    const insertStart = anchor.insertStartOffset ?? blocks[anchor.replacementStartIndex]?.start ?? anchor.target.end;
    const insertEnd = anchor.insertEndOffset ?? blocks[anchor.replacementEndIndex]?.start ?? bodyXml.length;
    const markdownBody = stripLeadingMatchingMarkdownHeading(compactedMarkdown, item.section, anchor.target.text);
    const replacementXml = buildSectionReplacementXml(
      markdownBody,
      blocks,
      anchor.target,
      anchor.sdtBlock,
      anchor.templateSection
    );
    const anchoredReplacementXml = anchor.sdtBlock ? replaceSdtContentXml(anchor.sdtBlock.xml, replacementXml) : replacementXml;
    xmlReplacements.push({
      section: item.section,
      start: insertStart,
      end: insertEnd,
      xml: anchoredReplacementXml,
      summary: {
        section: item.section,
        matchedHeading: formatMatchedHeading(anchor.target),
        replacementCount: Math.max(0, anchor.replacementEndIndex - anchor.replacementStartIndex),
        templateAnchorId: anchor.templateSection?.id
      }
    });
  }

  assertNonOverlappingSectionReplacements(xmlReplacements);
  let nextBodyXml = bodyXml;
  for (const replacement of [...xmlReplacements].sort((left, right) => right.start - left.start)) {
    nextBodyXml = `${nextBodyXml.slice(0, replacement.start)}${replacement.xml}${nextBodyXml.slice(replacement.end)}`;
  }

  const nextXml = `${documentXml.slice(0, bodyStart)}${nextBodyXml}${documentXml.slice(bodyEnd)}`;
  zip.file("word/document.xml", nextXml);
  return xmlReplacements.map((replacement) => replacement.summary);
}

function assertNonOverlappingSectionReplacements(
  replacements: Array<{ section: string; start: number; end: number }>
): void {
  const sorted = [...replacements].sort((left, right) => left.start - right.start);
  for (let index = 1; index < sorted.length; index += 1) {
    const previous = sorted[index - 1];
    const current = sorted[index];
    if (current.start < previous.end) {
      throw new Error(`Word 章节锚点范围重叠：${previous.section} / ${current.section}`);
    }
  }
}

function replaceSingleDocumentSectionWithMarkdown(
  zip: PizZip,
  section: string,
  markdown: string,
  templateJson?: WordTemplateJson
): SectionReplacementSummary | undefined {
  const compactedMarkdown = compactText(markdown, 120000);
  if (!compactedMarkdown.trim()) return undefined;

  const documentFile = zip.file("word/document.xml");
  const documentXml = documentFile?.asText();
  if (!documentXml) return undefined;

  const bodyOpen = documentXml.match(/<w:body\b[^>]*>/);
  const bodyEnd = documentXml.lastIndexOf("</w:body>");
  if (!bodyOpen || bodyOpen.index === undefined || bodyEnd < 0) return undefined;

  const bodyStart = bodyOpen.index + bodyOpen[0].length;
  const bodyXml = documentXml.slice(bodyStart, bodyEnd);
  const blocks = collectWordBodyBlocks(bodyXml, buildStyleHeadingLevels(zip));
  const anchor = findSectionReplacementAnchor(blocks, section, templateJson);
  if (!anchor) return undefined;

  const insertStart = anchor.insertStartOffset ?? blocks[anchor.replacementStartIndex]?.start ?? anchor.target.end;
  const insertEnd = anchor.insertEndOffset ?? blocks[anchor.replacementEndIndex]?.start ?? bodyXml.length;
  const markdownBody = stripLeadingMatchingMarkdownHeading(compactedMarkdown, section, anchor.target.text);
  const replacementXml = buildSectionReplacementXml(
    markdownBody,
    blocks,
    anchor.target,
    anchor.sdtBlock,
    anchor.templateSection
  );
  const anchoredReplacementXml = anchor.sdtBlock ? replaceSdtContentXml(anchor.sdtBlock.xml, replacementXml) : replacementXml;
  const nextBodyXml = `${bodyXml.slice(0, insertStart)}${anchoredReplacementXml}${bodyXml.slice(insertEnd)}`;
  const nextXml = `${documentXml.slice(0, bodyStart)}${nextBodyXml}${documentXml.slice(bodyEnd)}`;
  zip.file("word/document.xml", nextXml);

  return {
    matchedHeading: formatMatchedHeading(anchor.target),
    replacementCount: Math.max(0, anchor.replacementEndIndex - anchor.replacementStartIndex),
    templateAnchorId: anchor.templateSection?.id
  };
}

function extractMarkdownSectionSegments(markdown: string): MarkdownSectionSegment[] {
  const lines = markdown.split(/\r?\n/);
  const headings = lines
    .map((line, index) => ({ ...parseMarkdownSectionHeading(line), lineIndex: index }))
    .filter((heading): heading is ReturnType<typeof parseMarkdownSectionHeading> & { lineIndex: number } =>
      Boolean(heading.number)
    );
  const segments: MarkdownSectionSegment[] = [];

  for (const [index, heading] of headings.entries()) {
    const nextHeading = headings.slice(index + 1).find((candidate) => candidate.level <= heading.level);
    const nextAnyHeading = headings[index + 1];
    const bodyEnd = nextAnyHeading?.lineIndex ?? lines.length;
    const body = lines.slice(heading.lineIndex + 1, bodyEnd).join("\n").trim();
    const hasChildren = Boolean(nextHeading ? headings.slice(index + 1).some((candidate) => candidate.lineIndex < nextHeading.lineIndex && isChildSectionNumber(heading.number, candidate.number)) : headings.slice(index + 1).some((candidate) => isChildSectionNumber(heading.number, candidate.number)));
    segments.push({
      number: heading.number,
      title: heading.title,
      level: heading.level,
      body,
      hasChildren
    });
  }

  return segments;
}

function parseMarkdownSectionHeading(line: string): { number: string; title: string; level: number } {
  const trimmed = line.trim();
  const atx = trimmed.match(/^(#{1,6})\s+(.+?)\s*$/);
  const text = cleanMarkdownInline(atx?.[2] ?? trimmed).replace(/^#+\s*/, "");
  const match = text.match(/^(\d+(?:\.\d+)*)(?:[.．、]|\s+)(.+)$/);
  if (!match?.[1]) return { number: "", title: "", level: 0 };

  const number = match[1];
  const title = stripHeadingNumberPrefix(match[0]);
  const level = atx ? atx[1].length : number.split(".").length;
  if (!atx && !isLikelyDocumentSectionHeading(number, title)) return { number: "", title: "", level: 0 };
  return { number, title, level };
}

function isLikelyDocumentSectionHeading(number: string, title: string): boolean {
  if (number.includes(".")) return true;
  return new Set(["背景", "系统概述", "密码应用需求分析", "安全目标及设计原则", "密码应用设计", "安全管理方案", "安全与合规性分析", "实施保障方案"]).has(
    normalizeHeadingLookup(title)
  );
}

function selectPreciseMarkdownSectionSegments(
  section: string,
  segments: MarkdownSectionSegment[]
): MarkdownSectionSegment[] {
  const targetNumber = parseSectionNumber(section);
  if (!targetNumber) return [];

  const exact = segments.find((segment) => segment.number === targetNumber);
  const scopedSegments = segments.filter(
    (segment) =>
      segment.body.trim() &&
      !segment.hasChildren &&
      (segment.number === targetNumber || isChildSectionNumber(targetNumber, segment.number))
  );

  if (exact?.body.trim() && !exact.hasChildren) return [exact];
  return scopedSegments;
}

function isChildSectionNumber(parent: string, child: string): boolean {
  return child.startsWith(`${parent}.`) && child.length > parent.length + 1;
}

function compareSectionNumbersDescending(left: MarkdownSectionSegment, right: MarkdownSectionSegment): number {
  const leftParts = left.number.split(".").map(Number);
  const rightParts = right.number.split(".").map(Number);
  const length = Math.max(leftParts.length, rightParts.length);
  for (let index = 0; index < length; index += 1) {
    const diff = (rightParts[index] ?? -1) - (leftParts[index] ?? -1);
    if (diff !== 0) return diff;
  }
  return 0;
}

function collectWordBodyBlocks(bodyXml: string, headingLevels: Map<string, number>): WordBodyBlock[] {
  const blocks: WordBodyBlock[] = [];
  const counters: number[] = [];
  let cursor = 0;

  while (cursor < bodyXml.length) {
    const next = findNextBodyBlock(bodyXml, cursor);
    if (!next) break;

    const { xml, tagName, start, end } = next;
    const styleId = tagName === "p" ? extractParagraphStyleId(xml) : "";
    const headingLevel = tagName === "p" ? extractParagraphHeadingLevel(xml, styleId, headingLevels) : undefined;
    const sdtTag = tagName === "sdt" ? extractSdtTag(xml) : undefined;
    const block: WordBodyBlock = {
      index: blocks.length,
      start,
      end,
      tagName,
      xml,
      text: extractVisibleWordText(xml),
      styleId,
      ...(sdtTag ? { sdtTag } : {}),
      headingLevel
    };

    if (headingLevel) {
      counters[headingLevel - 1] = (counters[headingLevel - 1] ?? 0) + 1;
      counters.length = headingLevel;
      block.headingNumber = counters.slice(0, headingLevel).join(".");
    }

    blocks.push(block);
    cursor = end;
  }

  return blocks;
}

function findNextBodyBlock(
  bodyXml: string,
  cursor: number
): { xml: string; tagName: string; start: number; end: number } | undefined {
  const candidates = [
    { tagName: "p", start: findNextElementStart(bodyXml, "w:p", cursor) },
    { tagName: "tbl", start: findNextElementStart(bodyXml, "w:tbl", cursor) },
    { tagName: "sdt", start: findNextElementStart(bodyXml, "w:sdt", cursor) },
    { tagName: "sectPr", start: findNextElementStart(bodyXml, "w:sectPr", cursor) }
  ].filter((candidate) => candidate.start >= 0);

  if (!candidates.length) return undefined;
  const candidate = candidates.sort((left, right) => left.start - right.start)[0];
  const element = readBalancedWordElement(bodyXml, candidate.start, `w:${candidate.tagName}`);
  return element ? { ...element, tagName: candidate.tagName } : undefined;
}

function buildStyleHeadingLevels(zip: PizZip): Map<string, number> {
  const levels = new Map<string, number>();
  const stylesXml = zip.file("word/styles.xml")?.asText() ?? "";
  const pattern = /<w:style\b[\s\S]*?<\/w:style>/g;
  let match: RegExpExecArray | null;

  while ((match = pattern.exec(stylesXml))) {
    const styleXml = match[0];
    const styleId = styleXml.match(/\bw:styleId="([^"]+)"/)?.[1] ?? "";
    if (!styleId) continue;

    const outlineLevel = Number(styleXml.match(/<w:outlineLvl\b[^>]*\bw:val="(\d+)"/)?.[1]);
    if (Number.isFinite(outlineLevel) && outlineLevel >= 0 && outlineLevel <= 8) {
      levels.set(styleId, outlineLevel + 1);
      continue;
    }

    const styleName = styleXml.match(/<w:name\b[^>]*\bw:val="([^"]+)"/)?.[1] ?? "";
    const headingMatch = styleName.match(/heading\s*([1-9])/i);
    if (headingMatch?.[1]) levels.set(styleId, Number(headingMatch[1]));
  }

  return levels;
}

function findNextElementStart(xml: string, tagName: string, cursor: number): number {
  const pattern = new RegExp(`<${escapeRegExp(tagName)}(?=[\\s>/])`, "g");
  pattern.lastIndex = cursor;
  const match = pattern.exec(xml);
  return match ? match.index : -1;
}

function readBalancedWordElement(
  xml: string,
  start: number,
  tagName: string
): { xml: string; start: number; end: number } | undefined {
  const openTagEnd = xml.indexOf(">", start);
  if (openTagEnd < 0) return undefined;

  const openTag = xml.slice(start, openTagEnd + 1);
  if (openTag.endsWith("/>")) return { xml: openTag, start, end: openTagEnd + 1 };

  const closePattern = `</${tagName}>`;
  let cursor = openTagEnd + 1;
  let depth = 1;

  while (cursor < xml.length) {
    const nextOpen = findNextElementStart(xml, tagName, cursor);
    const nextClose = xml.indexOf(closePattern, cursor);
    if (nextClose < 0) return undefined;

    if (nextOpen >= 0 && nextOpen < nextClose) {
      const nestedOpenEnd = xml.indexOf(">", nextOpen);
      if (nestedOpenEnd < 0) return undefined;
      const nestedOpenTag = xml.slice(nextOpen, nestedOpenEnd + 1);
      if (!nestedOpenTag.endsWith("/>")) depth += 1;
      cursor = nestedOpenEnd + 1;
      continue;
    }

    depth -= 1;
    cursor = nextClose + closePattern.length;
    if (depth === 0) return { xml: xml.slice(start, cursor), start, end: cursor };
  }

  return undefined;
}

function extractSdtTag(xml: string): string {
  const contentStart = xml.indexOf("<w:sdtContent");
  const propertiesXml = contentStart >= 0 ? xml.slice(0, contentStart) : xml;
  const rawTag = propertiesXml.match(/<w:tag\b[^>]*\bw:val="([^"]+)"/)?.[1] ?? "";
  return normalizeContentControlTag(rawTag);
}

function extractSdtAlias(xml: string): string | undefined {
  const contentStart = xml.indexOf("<w:sdtContent");
  const propertiesXml = contentStart >= 0 ? xml.slice(0, contentStart) : xml;
  const rawAlias = propertiesXml.match(/<w:alias\b[^>]*\bw:val="([^"]+)"/)?.[1] ?? "";
  const alias = normalizeContentControlTag(rawAlias);
  return alias || undefined;
}

function findSdtElementByTags(
  xml: string,
  tags: string[],
  baseOffset = 0
): { xml: string; start: number; end: number; matchedTag: string } | undefined {
  const tagCandidates = normalizeContentControlTagCandidates(tags);
  if (!tagCandidates.length) return undefined;
  const tagSet = new Set(tagCandidates);
  const lowerTagSet = new Set(tagCandidates.map((item) => item.toLowerCase()));
  let cursor = 0;

  while (cursor < xml.length) {
    const start = findNextElementStart(xml, "w:sdt", cursor);
    if (start < 0) return undefined;

    const element = readBalancedWordElement(xml, start, "w:sdt");
    if (!element) return undefined;
    const elementTag = extractSdtTag(element.xml);
    if (tagSet.has(elementTag) || lowerTagSet.has(elementTag.toLowerCase())) {
      return {
        xml: element.xml,
        start: baseOffset + element.start,
        end: baseOffset + element.end,
        matchedTag: elementTag
      };
    }

    const contentRange = getSdtContentRange(element.xml);
    const nested = contentRange
      ? findSdtElementByTags(
          element.xml.slice(contentRange.start, contentRange.end),
          tagCandidates,
          baseOffset + element.start + contentRange.start
        )
      : undefined;
    if (nested) return nested;
    cursor = element.end;
  }

  return undefined;
}

function extractSdtContentXml(xml: string): string {
  const range = getSdtContentRange(xml);
  return range ? xml.slice(range.start, range.end) : "";
}

function replaceSdtContentXml(sdtXml: string, contentXml: string): string {
  const range = getSdtContentRange(sdtXml);
  if (!range) return contentXml;
  return `${sdtXml.slice(0, range.start)}${contentXml}${sdtXml.slice(range.end)}`;
}

function getSdtContentRange(xml: string): { start: number; end: number } | undefined {
  const match = xml.match(/<w:sdtContent\b[^>]*>/);
  const end = xml.lastIndexOf("</w:sdtContent>");
  if (!match || match.index === undefined || end < 0) return undefined;
  return { start: match.index + match[0].length, end };
}

function extractParagraphStyleId(paragraphXml: string): string {
  return paragraphXml.match(/<w:pStyle\b[^>]*\bw:val="([^"]+)"/)?.[1] ?? "";
}

function extractParagraphHeadingLevel(
  paragraphXml: string,
  styleId: string,
  headingLevels: Map<string, number>
): number | undefined {
  const directOutlineLevel = Number(paragraphXml.match(/<w:outlineLvl\b[^>]*\bw:val="(\d+)"/)?.[1]);
  if (Number.isFinite(directOutlineLevel) && directOutlineLevel >= 0 && directOutlineLevel <= 8) {
    return directOutlineLevel + 1;
  }
  return headingLevels.get(styleId);
}

function extractVisibleWordText(xml: string): string {
  return Array.from(xml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g))
    .map((match) => decodeXmlText(match[1]))
    .join("")
    .replace(/\s+/g, " ")
    .trim();
}

function findSectionReplacementAnchor(
  blocks: WordBodyBlock[],
  section: string,
  templateJson?: WordTemplateJson
): SectionReplacementAnchor | undefined {
  const templateSection = findTemplateSection(templateJson, section);
  return templateSection ? buildTemplateSectionReplacementAnchor(blocks, templateSection) : undefined;
}

function findTemplateSection(templateJson: WordTemplateJson | undefined, section: string): WordTemplateSection | undefined {
  const targetNumber = parseSectionNumber(section);
  const targetTitle = normalizeHeadingLookup(section);
  const targetAnchorId = normalizeTemplateAnchorId(section);
  const sections = templateJson?.sections ?? [];

  if (targetAnchorId) {
    const exact = sections.find((item) => normalizeTemplateAnchorId(item.id) === targetAnchorId);
    if (exact) return exact;
  }

  if (targetNumber) {
    const exact = sections.find((item) => item.number === targetNumber);
    if (exact) return exact;
  }

  if (!targetTitle) return undefined;
  const titleMatches = sections.filter((item) => {
    const title = normalizeHeadingLookup(item.title);
    const combined = normalizeHeadingLookup(`${item.number} ${item.title}`);
    const dotted = normalizeHeadingLookup(`${item.number}.${item.title}`);
    return title === targetTitle || combined === targetTitle || dotted === targetTitle;
  });
  return titleMatches.length === 1 ? titleMatches[0] : undefined;
}

function buildTemplateSectionReplacementAnchor(
  blocks: WordBodyBlock[],
  templateSection: WordTemplateSection
): SectionReplacementAnchor | undefined {
  const sectionNumber = parseSectionNumberFromTemplateId(templateSection.id) || templateSection.number;
  const headingLevel = templateSection.headingLevel ?? (sectionNumber.split(".").filter(Boolean).length || 1);
  const candidateTags = getSectionBodyAnchorTags(templateSection);
  const sdtBlock = findSdtBlockByTags(blocks, candidateTags);
  if (!sdtBlock) return undefined;

  const target = findPreviousHeadingBlock(blocks, sdtBlock.index, headingLevel) ?? {
    ...sdtBlock,
    tagName: "p",
    headingLevel,
    headingNumber: sectionNumber,
    text: templateSection.title
  };

  return {
    target: {
      ...target,
      headingNumber: sectionNumber,
      headingLevel,
      text: templateSection.title || target.text
    },
    replacementStartIndex: sdtBlock.index,
    replacementEndIndex: sdtBlock.index + 1,
    insertStartOffset: sdtBlock.start,
    insertEndOffset: sdtBlock.end,
    sdtBlock,
    templateSection,
    contentControlTag: sdtBlock.sdtTag,
    candidateContentControlTags: candidateTags
  };
}

function getSectionBodyAnchorTags(section: WordTemplateSection): string[] {
  return buildContentControlTagCandidates(section.anchors?.body, [`ps:section:${section.id}:body`]);
}

function findSdtBlockByTags(blocks: WordBodyBlock[], tags: string[]): WordBodyBlock | undefined {
  const tagCandidates = normalizeContentControlTagCandidates(tags);
  if (!tagCandidates.length) return undefined;
  const tagSet = new Set(tagCandidates);
  const lowerTagSet = new Set(tagCandidates.map((item) => item.toLowerCase()));
  return blocks.find((block) => {
    if (block.tagName !== "sdt" || !block.sdtTag) return false;
    return tagSet.has(block.sdtTag) || lowerTagSet.has(block.sdtTag.toLowerCase());
  });
}

function getTemplateSectionAnchorTagHint(templateJson: WordTemplateJson | undefined, section: string): string {
  const templateSection = findTemplateSection(templateJson, section);
  if (!templateSection) return "";
  const tags = getSectionBodyAnchorTags(templateSection);
  return tags.length ? `（尝试 Content Control tag：${tags.join("、")}）` : "";
}

function buildContentControlTagCandidates(anchor: WordTemplateAnchor | undefined, fallbackTags: string[]): string[] {
  return normalizeContentControlTagCandidates([
    anchor?.tag,
    anchor?.alias,
    ...fallbackTags
  ]);
}

function normalizeContentControlTagCandidates(tags: Array<string | undefined>): string[] {
  return unique(tags.map((tag) => normalizeContentControlTag(tag ?? "")).filter(Boolean));
}

function normalizeContentControlTag(tag: string): string {
  return decodeXmlText(tag).trim();
}

function findPreviousHeadingBlock(
  blocks: WordBodyBlock[],
  beforeIndex: number,
  headingLevel?: number
): WordBodyBlock | undefined {
  for (let index = beforeIndex - 1; index >= 0; index -= 1) {
    const block = blocks[index];
    if (!block?.headingLevel) continue;
    if (!headingLevel || block.headingLevel <= headingLevel) return block;
  }
  return undefined;
}

function replaceTemplateTableCells(
  zip: PizZip,
  templateJson: WordTemplateJson | undefined,
  replacements: TemplateCellReplacementInput[]
): number {
  if (!templateJson?.tables?.length || !replacements.length) return 0;

  let replacementCount = 0;
  for (const replacement of replacements) {
    if (!replacement.value.trim()) continue;
    if (replaceSingleTemplateTableCell(zip, templateJson, replacement)) replacementCount += 1;
  }
  return replacementCount;
}

function replaceTemplateTables(
  zip: PizZip,
  templateJson: WordTemplateJson | undefined,
  replacements: TemplateTableReplacementInput[]
): number {
  if (!templateJson?.tables?.length || !replacements.length) return 0;

  let replacementCount = 0;
  for (const replacement of replacements) {
    if (replaceSingleTemplateTable(zip, templateJson, replacement)) replacementCount += 1;
  }
  return replacementCount;
}

function replaceSingleTemplateTable(
  zip: PizZip,
  templateJson: WordTemplateJson,
  replacement: TemplateTableReplacementInput
): boolean {
  const table = findTemplateTable(templateJson, replacement);
  if (!table) {
    throw new Error(`未找到模板表格：${replacement.tableId || replacement.caption || "未提供 table_id/caption"}`);
  }

  const documentFile = zip.file("word/document.xml");
  const documentXml = documentFile?.asText();
  if (!documentXml) return false;

  const bodyOpen = documentXml.match(/<w:body\b[^>]*>/);
  const bodyEnd = documentXml.lastIndexOf("</w:body>");
  if (!bodyOpen || bodyOpen.index === undefined || bodyEnd < 0) return false;

  const bodyStart = bodyOpen.index + bodyOpen[0].length;
  const bodyXml = documentXml.slice(bodyStart, bodyEnd);
  const anchoredTable = findSdtElementByTags(bodyXml, getTableAnchorTags(table));
  if (!anchoredTable) {
    throw new Error(`未找到模板表格锚点：${table.id}（${table.caption || "未命名表格"}）`);
  }

  const replacementContentXml = buildTemplateTableReplacementContentXml(anchoredTable.xml, replacement, table);
  const replacementXml = replaceSdtContentXml(anchoredTable.xml, replacementContentXml);
  const nextBodyXml = `${bodyXml.slice(0, anchoredTable.start)}${replacementXml}${bodyXml.slice(anchoredTable.end)}`;
  const nextXml = `${documentXml.slice(0, bodyStart)}${nextBodyXml}${documentXml.slice(bodyEnd)}`;
  zip.file("word/document.xml", nextXml);
  return true;
}

function buildTemplateTableReplacementContentXml(
  sdtXml: string,
  replacement: TemplateTableReplacementInput,
  table: WordTemplateTable
): string {
  const rows = resolveTemplateTableReplacementRows(replacement, table);
  const anchorBlocks = collectWordBodyBlocks(extractSdtContentXml(sdtXml), new Map());
  const paragraphTemplate =
    extractSdtContentXml(sdtXml).match(/<w:p\b[\s\S]*?<\/w:p>/)?.[0] ??
    findBodyParagraphTemplate(anchorBlocks, -1) ??
    buildDefaultBodyParagraphTemplate();
  const tableTemplate = findTableTemplate(anchorBlocks, -1) ?? anchorBlocks.find((block) => block.tagName === "tbl")?.xml;
  return buildWordTableFromMarkdown(rows, paragraphTemplate, tableTemplate);
}

function resolveTemplateTableReplacementRows(
  replacement: TemplateTableReplacementInput,
  table: WordTemplateTable
): string[][] {
  const explicitRows = normalizeTemplateTableReplacementRows(replacement.rows);
  if (explicitRows) return explicitRows;

  const markdown = replacement.markdown?.trim() ?? "";
  if (!markdown) {
    throw new Error(`模板整表替换缺少内容：${table.id}（请提供 markdown 或 rows）`);
  }

  const lines = markdown.split(/\r?\n/);
  for (let index = 0; index < lines.length; index += 1) {
    const parsed = parseMarkdownTable(lines, index);
    if (parsed) return parsed.rows;
  }

  throw new Error(`模板整表替换未解析到 Markdown 表格：${table.id}`);
}

function normalizeTemplateTableReplacementRows(rows: string[][] | undefined): string[][] | undefined {
  if (!Array.isArray(rows) || !rows.length) return undefined;
  const normalizedRows = rows
    .map((row) =>
      Array.isArray(row)
        ? row.map((cell) => (cell == null ? "" : typeof cell === "string" ? cell.trim() : String(cell).trim()))
        : []
    )
    .filter((row) => row.length);
  if (!normalizedRows.length || normalizedRows.some((row) => !row.length)) return undefined;
  return normalizeMarkdownTableRows(normalizedRows);
}

function replaceContentControlsByTag(
  zip: PizZip,
  replacements: ContentControlReplacementInput[],
  templateJson?: WordTemplateJson
): number {
  const validReplacements = replacements
    .map((replacement) => ({
      tags: resolveContentControlReplacementTags(replacement, templateJson),
      value: replacement.value.trim()
    }))
    .filter((replacement) => replacement.tags.length && replacement.value);
  if (!validReplacements.length) return 0;

  const documentFile = zip.file("word/document.xml");
  const documentXml = documentFile?.asText();
  if (!documentXml) return 0;

  const bodyOpen = documentXml.match(/<w:body\b[^>]*>/);
  const bodyEnd = documentXml.lastIndexOf("</w:body>");
  if (!bodyOpen || bodyOpen.index === undefined || bodyEnd < 0) return 0;

  const bodyStart = bodyOpen.index + bodyOpen[0].length;
  let bodyXml = documentXml.slice(bodyStart, bodyEnd);
  let replacementCount = 0;

  for (const replacement of validReplacements) {
    const anchoredControl = findSdtElementByTags(bodyXml, replacement.tags);
    if (!anchoredControl) {
      throw new Error(`未找到 Word Content Control tag：${replacement.tags.join("、")}`);
    }
    const replacementXml = replaceSdtContentXml(
      anchoredControl.xml,
      buildContentControlReplacementContentXml(anchoredControl.xml, replacement.value)
    );
    bodyXml = `${bodyXml.slice(0, anchoredControl.start)}${replacementXml}${bodyXml.slice(anchoredControl.end)}`;
    replacementCount += 1;
  }

  zip.file("word/document.xml", `${documentXml.slice(0, bodyStart)}${bodyXml}${documentXml.slice(bodyEnd)}`);
  return replacementCount;
}

function resolveContentControlReplacementTags(
  replacement: ContentControlReplacementInput,
  templateJson?: WordTemplateJson
): string[] {
  const directTags = normalizeContentControlTagCandidates([replacement.tag]);
  if (!templateJson) return directTags;
  return unique([...directTags, ...resolveTemplateContentControlTags(templateJson, directTags)]);
}

function resolveTemplateContentControlTags(templateJson: WordTemplateJson, lookupTags: string[]): string[] {
  const requestedTags = normalizeContentControlTagCandidates(lookupTags);
  if (!requestedTags.length) return [];

  const requestedLookupSet = new Set(requestedTags.map((tag) => tag.toLowerCase()));
  const resolved: string[] = [];

  for (const anchor of collectTemplateContentControlAnchors(templateJson)) {
    if (anchor.lookupKeys.some((key) => requestedLookupSet.has(key.toLowerCase()))) {
      resolved.push(...anchor.tags);
    }
  }

  return unique(resolved);
}

function collectTemplateContentControlAnchors(templateJson: WordTemplateJson): Array<{ lookupKeys: string[]; tags: string[] }> {
  const anchors: Array<{ lookupKeys: string[]; tags: string[] }> = [];

  for (const section of templateJson.sections ?? []) {
    const sectionKeys = normalizeContentControlTagCandidates([
      section.id,
      section.number,
      section.title,
      `${section.number} ${section.title}`,
      `${section.number}.${section.title}`
    ]);
    anchors.push({
      lookupKeys: unique([...sectionKeys, ...getSectionBodyAnchorTags(section)]),
      tags: getSectionBodyAnchorTags(section)
    });
  }

  for (const table of templateJson.tables ?? []) {
    const tableKeys = normalizeContentControlTagCandidates([table.id, table.caption]);
    anchors.push({
      lookupKeys: unique([...tableKeys, ...getTableAnchorTags(table)]),
      tags: getTableAnchorTags(table)
    });

    if (table.anchors?.caption) {
      const captionTags = getTableCaptionAnchorTags(table);
      anchors.push({
        lookupKeys: unique([...tableKeys, `${table.caption ?? table.id} 题注`, ...captionTags]),
        tags: captionTags
      });
    }
  }

  for (const figure of templateJson.figures ?? []) {
    const figureKeys = normalizeContentControlTagCandidates([figure.id, figure.caption]);
    if (figure.anchors?.image) {
      anchors.push({
        lookupKeys: unique([...figureKeys, `${figure.caption} 图片`, ...getFigureImageAnchorTags(figure)]),
        tags: getFigureImageAnchorTags(figure)
      });
    }

    if (figure.anchors?.caption) {
      const captionTags = getFigureCaptionAnchorTags(figure);
      anchors.push({
        lookupKeys: unique([...figureKeys, `${figure.caption} 题注`, ...captionTags]),
        tags: captionTags
      });
    }
  }

  for (const fieldBlock of templateJson.fieldBlocks ?? []) {
    const fieldBlockTags = getFieldBlockAnchorTags(fieldBlock);
    anchors.push({
      lookupKeys: unique([fieldBlock.id, ...fieldBlockTags]),
      tags: fieldBlockTags
    });
  }

  for (const textBlock of templateJson.textBlocks ?? []) {
    const textBlockTags = getSectionTextBlockAnchorTags(textBlock);
    anchors.push({
      lookupKeys: unique([
        textBlock.id,
        `${textBlock.section}:text:${textBlock.order}`,
        textBlock.sectionNumber ? `${textBlock.sectionNumber}:text:${textBlock.order}` : "",
        ...textBlockTags
      ]),
      tags: textBlockTags
    });
  }

  return anchors;
}

function buildContentControlReplacementContentXml(sdtXml: string, value: string): string {
  const anchorBlocks = collectWordBodyBlocks(extractSdtContentXml(sdtXml), new Map());
  const paragraphTemplate = findBodyParagraphTemplate(anchorBlocks, -1) ?? buildDefaultBodyParagraphTemplate();
  const tableTemplate = findTableTemplate(anchorBlocks, -1);
  const blocks: string[] = [];
  const lines = value.split(/\r?\n/);
  let inCodeBlock = false;

  for (let lineIndex = 0; lineIndex < lines.length; lineIndex += 1) {
    const line = lines[lineIndex].trim();
    if (/^```/.test(line)) {
      inCodeBlock = !inCodeBlock;
      continue;
    }
    if (!line) continue;

    if (!inCodeBlock) {
      const table = parseMarkdownTable(lines, lineIndex);
      if (table) {
        blocks.push(buildWordTableFromMarkdown(table.rows, paragraphTemplate, tableTemplate));
        lineIndex = table.nextIndex - 1;
        continue;
      }
    }

    const text = line.replace(/^#{1,6}\s+/, "");
    blocks.push(buildWordParagraphFromTemplate(paragraphTemplate, cleanMarkdownInline(text)));
  }

  return blocks.length ? blocks.join("") : buildWordParagraphFromTemplate(paragraphTemplate, "");
}

function replaceSingleTemplateTableCell(
  zip: PizZip,
  templateJson: WordTemplateJson,
  replacement: TemplateCellReplacementInput
): boolean {
  const table = findTemplateTable(templateJson, replacement);
  if (!table) return false;

  const cellIndex = resolveTemplateCellIndex(table, replacement);
  if (cellIndex < 0 || replacement.rowIndex < 0) return false;

  const documentFile = zip.file("word/document.xml");
  const documentXml = documentFile?.asText();
  if (!documentXml) return false;

  const bodyOpen = documentXml.match(/<w:body\b[^>]*>/);
  const bodyEnd = documentXml.lastIndexOf("</w:body>");
  if (!bodyOpen || bodyOpen.index === undefined || bodyEnd < 0) return false;

  const bodyStart = bodyOpen.index + bodyOpen[0].length;
  const bodyXml = documentXml.slice(bodyStart, bodyEnd);
  const anchoredTable = findSdtElementByTags(bodyXml, getTableAnchorTags(table));
  if (!anchoredTable) return false;

  const nextTableXml = replaceTableCellXml(anchoredTable.xml, replacement.rowIndex, cellIndex, replacement.value);
  if (!nextTableXml || nextTableXml === anchoredTable.xml) return false;

  const nextBodyXml = `${bodyXml.slice(0, anchoredTable.start)}${nextTableXml}${bodyXml.slice(anchoredTable.end)}`;
  const nextXml = `${documentXml.slice(0, bodyStart)}${nextBodyXml}${documentXml.slice(bodyEnd)}`;
  zip.file("word/document.xml", nextXml);
  return true;
}

function findTemplateTable(
  templateJson: WordTemplateJson,
  replacement: Pick<TemplateCellReplacementInput, "tableId" | "caption"> | Pick<TemplateTableReplacementInput, "tableId" | "caption">
): WordTemplateTable | undefined {
  if (replacement.tableId) {
    const table = templateJson.tables?.find((item) => item.id === replacement.tableId);
    if (table) return table;
  }

  const caption = normalizeHeadingLookup(replacement.caption ?? "");
  if (!caption) return undefined;
  return templateJson.tables?.find((item) => normalizeHeadingLookup(item.caption ?? "") === caption);
}

function getTableAnchorTags(table: WordTemplateTable): string[] {
  return buildContentControlTagCandidates(table.anchors?.table, [`ps:table:${table.id}`]);
}

function getTableCaptionAnchorTags(table: WordTemplateTable): string[] {
  return buildContentControlTagCandidates(table.anchors?.caption, [`ps:table:${table.id}:caption`]);
}

function resolveTemplateCellIndex(table: WordTemplateTable, replacement: TemplateCellReplacementInput): number {
  if (typeof replacement.cellIndex === "number" && Number.isInteger(replacement.cellIndex)) return replacement.cellIndex;
  if (typeof replacement.columnIndex !== "number" || !Number.isInteger(replacement.columnIndex)) return -1;

  const row = table.rows?.find((item) => item.index === replacement.rowIndex);
  return row?.cells.find((cell) => cell.columnIndex === replacement.columnIndex)?.cellIndex ?? -1;
}

function replaceTableCellXml(tableXml: string, rowIndex: number, cellIndex: number, value: string): string | undefined {
  const rows = Array.from(tableXml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g));
  const rowMatch = rows[rowIndex];
  if (!rowMatch || rowMatch.index === undefined) return undefined;

  const rowXml = rowMatch[0];
  const cells = Array.from(rowXml.matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g));
  const cellMatch = cells[cellIndex];
  if (!cellMatch || cellMatch.index === undefined) return undefined;

  const nextCellXml = replaceWordTableCellText(cellMatch[0], value);
  const nextRowXml = `${rowXml.slice(0, cellMatch.index)}${nextCellXml}${rowXml.slice(cellMatch.index + cellMatch[0].length)}`;
  return `${tableXml.slice(0, rowMatch.index)}${nextRowXml}${tableXml.slice(rowMatch.index + rowXml.length)}`;
}

function replaceWordTableCellText(cellXml: string, value: string): string {
  const cellProperties = cellXml.match(/<w:tcPr\b[\s\S]*?<\/w:tcPr>/)?.[0] ?? "";
  const paragraphTemplate = cellXml.match(/<w:p\b[\s\S]*?<\/w:p>/)?.[0] ?? "<w:p><w:r><w:t></w:t></w:r></w:p>";
  return ["<w:tc>", cellProperties, buildWordParagraphFromTemplate(paragraphTemplate, value), "</w:tc>"].join("");
}

function parseSectionNumber(section: string): string {
  return section.trim().match(/^(\d+(?:\.\d+)*)(?:[.．、\s]|$)/)?.[1] ?? "";
}

function parseSectionNumberFromTemplateId(value: string): string {
  const match = value.trim().toLowerCase().match(/^sec_(\d+(?:_\d+)*)$/);
  return match?.[1]?.replace(/_/g, ".") ?? "";
}

function normalizeTemplateAnchorId(value: string): string {
  return value.trim().toLowerCase().match(/^sec_\d+(?:_\d+)*$/)?.[0] ?? "";
}

function normalizeHeadingLookup(value: string): string {
  return value
    .replace(/^#{1,6}\s*/, "")
    .replace(/^\d+(?:\.\d+)*[.．、]?\s*/, "")
    .replace(/[\s\t　:：,，.。;；、\-—_]/g, "")
    .toLowerCase()
    .trim();
}

function stripLeadingMatchingMarkdownHeading(markdown: string, section: string, matchedHeading: string): string {
  const lines = markdown.split(/\r?\n/);
  const firstContentIndex = lines.findIndex((line) => line.trim());
  if (firstContentIndex < 0) return "";

  const firstLine = lines[firstContentIndex].trim();
  if (!/^#{1,6}\s+/.test(firstLine)) return markdown;

  const headingText = firstLine.replace(/^#{1,6}\s+/, "");
  const firstNumber = parseSectionNumber(headingText);
  const sectionNumber = parseSectionNumber(section) || parseSectionNumber(matchedHeading);
  const firstTitle = normalizeHeadingLookup(headingText);
  const sectionTitle = normalizeHeadingLookup(section);
  const matchedTitle = normalizeHeadingLookup(matchedHeading);
  if (
    (firstNumber && sectionNumber && firstNumber === sectionNumber) ||
    firstTitle === sectionTitle ||
    firstTitle === matchedTitle
  ) {
    return [...lines.slice(0, firstContentIndex), ...lines.slice(firstContentIndex + 1)].join("\n").trim();
  }

  return markdown;
}

function buildSectionReplacementXml(
  markdown: string,
  blocks: WordBodyBlock[],
  target: WordBodyBlock,
  sdtBlock?: WordBodyBlock,
  templateSection?: WordTemplateSection
): string {
  const preservedStructuredSectionXml =
    sdtBlock && templateSection ? buildStructuredSectionReplacementXml(markdown, sdtBlock, templateSection.id) : undefined;
  if (preservedStructuredSectionXml !== undefined) {
    return preservedStructuredSectionXml;
  }

  const anchorBlocks = sdtBlock ? collectWordBodyBlocks(extractSdtContentXml(sdtBlock.xml), new Map()) : [];
  const headingStyleIds = collectHeadingStyleIds(blocks);
  const paragraphTemplate =
    findBodyParagraphTemplate(anchorBlocks, -1, headingStyleIds) ??
    findBodyParagraphTemplate(blocks, target.index, headingStyleIds) ??
    findBodyParagraphTemplate(blocks, -1, headingStyleIds) ??
    buildDefaultBodyParagraphTemplate();
  const tableTemplate = findTableTemplate(anchorBlocks, -1) ?? findTableTemplate(blocks, target.index);
  const headingTemplates = buildHeadingTemplates(blocks);
  const paragraphs: string[] = [];
  let inCodeBlock = false;
  const lines = markdown.split(/\r?\n/);

  for (let lineIndex = 0; lineIndex < lines.length; lineIndex += 1) {
    const rawLine = lines[lineIndex];
    const line = rawLine.trim();
    if (/^```/.test(line)) {
      inCodeBlock = !inCodeBlock;
      continue;
    }
    if (!line) continue;

    if (!inCodeBlock && /^#{1,6}\s+/.test(line)) {
      const headingText = line.replace(/^#{1,6}\s+/, "");
      const explicitLevel = parseSectionNumber(headingText).split(".").filter(Boolean).length;
      const fallbackLevel = Math.min((target.headingLevel ?? 1) + line.match(/^#+/)![0].length, 9);
      const headingLevel = explicitLevel || fallbackLevel;
      const template = headingTemplates.get(headingLevel) ?? headingTemplates.get(target.headingLevel ?? 1) ?? target.xml;
      paragraphs.push(buildWordParagraphFromTemplate(template, stripHeadingNumberPrefix(cleanMarkdownInline(headingText))));
      continue;
    }

    const table = !inCodeBlock ? parseMarkdownTable(lines, lineIndex) : undefined;
    if (table) {
      paragraphs.push(buildWordTableFromMarkdown(table.rows, paragraphTemplate, tableTemplate));
      lineIndex = table.nextIndex - 1;
      continue;
    }

    paragraphs.push(buildWordParagraphFromTemplate(paragraphTemplate, cleanMarkdownInline(line)));
  }

  return paragraphs.join("");
}

function buildStructuredSectionReplacementXml(
  markdown: string,
  sdtBlock: WordBodyBlock,
  sectionId: string
): string | undefined {
  const sectionContentXml = extractSdtContentXml(sdtBlock.xml);
  if (!sectionContentXml.trim()) return undefined;

  const sectionBlocks = collectWordBodyBlocks(sectionContentXml, new Map());
  const textBlocks = sectionBlocks.filter((block) => isSectionTextBlockForSection(block, sectionId));
  if (!textBlocks.length || !sectionBlocks.some((block) => isProtectedStructuredSectionBlock(block))) {
    return undefined;
  }

  const markdownChunks = splitMarkdownAcrossSectionTextBlocks(markdown, textBlocks.length);
  let nextContentXml = sectionContentXml;

  for (let index = 0; index < textBlocks.length; index += 1) {
    const textBlock = textBlocks[index];
    const tag = textBlock.sdtTag;
    if (!tag) continue;

    const anchoredBlock = findSdtElementByTags(nextContentXml, [tag]);
    if (!anchoredBlock) continue;

    const replacementContentXml = buildContentControlReplacementContentXml(anchoredBlock.xml, markdownChunks[index] ?? "");
    const replacementXml = replaceSdtContentXml(anchoredBlock.xml, replacementContentXml);
    nextContentXml = `${nextContentXml.slice(0, anchoredBlock.start)}${replacementXml}${nextContentXml.slice(anchoredBlock.end)}`;
  }

  return nextContentXml;
}

function isSectionTextBlockForSection(block: WordBodyBlock, sectionId: string): boolean {
  if (block.tagName !== "sdt" || !block.sdtTag) return false;
  const normalizedTag = block.sdtTag.toLowerCase();
  const normalizedSectionId = sectionId.toLowerCase();
  return normalizedTag.startsWith(`ps:section:${normalizedSectionId}:text:`);
}

function isProtectedStructuredSectionBlock(block: WordBodyBlock): boolean {
  if (block.tagName !== "sdt" || !block.sdtTag) return false;
  const normalizedTag = block.sdtTag.toLowerCase();
  return normalizedTag.startsWith("ps:table:") || normalizedTag.startsWith("ps:figure:");
}

function splitMarkdownAcrossSectionTextBlocks(markdown: string, blockCount: number): string[] {
  const normalizedMarkdown = markdown.trim();
  if (blockCount <= 1) return [normalizedMarkdown];
  if (!normalizedMarkdown) return Array.from({ length: blockCount }, () => "");

  const paragraphUnits = normalizedMarkdown
    .split(/\n{2,}/)
    .map((item) => item.trim())
    .filter(Boolean);
  const lineUnits = normalizedMarkdown
    .split(/\r?\n/)
    .map((item) => item.trim())
    .filter(Boolean);
  const units = paragraphUnits.length > 1 ? paragraphUnits : lineUnits.length ? lineUnits : [normalizedMarkdown];
  const separator = paragraphUnits.length > 1 ? "\n\n" : "\n";
  const chunks = Array.from({ length: blockCount }, () => "");
  const baseSize = Math.floor(units.length / blockCount);
  const remainder = units.length % blockCount;
  let cursor = 0;

  for (let index = 0; index < blockCount; index += 1) {
    const chunkSize = baseSize + (index < remainder ? 1 : 0);
    if (chunkSize <= 0) continue;
    chunks[index] = units.slice(cursor, cursor + chunkSize).join(separator).trim();
    cursor += chunkSize;
  }

  if (!chunks.some(Boolean)) chunks[0] = normalizedMarkdown;
  return chunks;
}

function findBodyParagraphTemplate(
  blocks: WordBodyBlock[],
  headingIndex: number,
  headingStyleIds = collectHeadingStyleIds(blocks)
): string | undefined {
  const localBlocks = blocks.filter((block, index) => index > headingIndex);
  const localTemplate = pickBestBodyParagraphTemplate(localBlocks, headingStyleIds, 40);
  if (localTemplate) return localTemplate;

  const globalTemplate = pickBestBodyParagraphTemplate(blocks, headingStyleIds, 40);
  if (globalTemplate) return globalTemplate;
  return undefined;
}

function collectHeadingStyleIds(blocks: WordBodyBlock[]): Set<string> {
  return new Set(blocks.filter((block) => block.headingLevel && block.styleId).map((block) => block.styleId));
}

function pickBestBodyParagraphTemplate(
  blocks: WordBodyBlock[],
  headingStyleIds: Set<string>,
  minScore = Number.NEGATIVE_INFINITY
): string | undefined {
  const candidates = blocks
    .flatMap((block) => collectBodyParagraphTemplateCandidates(block, headingStyleIds))
    .filter((candidate) => candidate.score >= minScore)
    .sort((left, right) => right.score - left.score);
  return candidates[0]?.xml;
}

function collectBodyParagraphTemplateCandidates(
  block: WordBodyBlock,
  headingStyleIds: Set<string>
): Array<{ xml: string; score: number }> {
  if (block.tagName === "p") {
    const score = scoreBodyParagraphTemplate(block.xml, headingStyleIds);
    return Number.isFinite(score) ? [{ xml: block.xml, score }] : [];
  }

  if (block.tagName !== "sdt") return [];
  return collectWordBodyBlocks(extractSdtContentXml(block.xml), new Map()).flatMap((nestedBlock) =>
    collectBodyParagraphTemplateCandidates(nestedBlock, headingStyleIds)
  );
}

function scoreBodyParagraphTemplate(paragraphXml: string, headingStyleIds: Set<string>): number {
  const text = extractVisibleWordText(paragraphXml);
  if (!text || /^\d+$/.test(text)) return Number.NEGATIVE_INFINITY;

  const styleId = extractParagraphStyleId(paragraphXml);
  if (styleId && headingStyleIds.has(styleId)) return Number.NEGATIVE_INFINITY;
  if (styleId === "12" || /^[表图]\s*\d/.test(text)) return Number.NEGATIVE_INFINITY;
  const paragraphProperties = paragraphXml.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/)?.[0] ?? "";
  if (/<w:tabs\b[\s\S]*?<w:tab\b[^>]*\bw:leader="dot"/.test(paragraphProperties)) {
    return Number.NEGATIVE_INFINITY;
  }

  let score = 0;
  if (styleId === "30") score += 100;
  if (styleId) score += 35;
  if (/<w:rFonts\b/.test(paragraphXml)) score += 8;
  if (/<w:sz(?:Cs)?\b/.test(paragraphXml)) score += 8;
  if (/<w:rPr\b/.test(paragraphProperties)) score -= 70;
  if (/<w:b(?:Cs)?\b/.test(paragraphXml)) score -= 90;
  if (/<w:i(?:Cs)?\b/.test(paragraphXml)) score -= 35;
  if (/<w:sz(?:Cs)?\b/.test(paragraphProperties)) score -= 45;
  if (/<w:numPr\b/.test(paragraphXml)) score -= 25;
  if (/<w:tbl\b/.test(paragraphXml)) score -= 50;
  return score;
}

function buildDefaultBodyParagraphTemplate(): string {
  return [
    "<w:p>",
    '<w:pPr><w:pStyle w:val="30"/><w:wordWrap w:val="0"/></w:pPr>',
    '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t></w:t></w:r>',
    "</w:p>"
  ].join("");
}

function findTableTemplate(blocks: WordBodyBlock[], headingIndex: number): string | undefined {
  const target = blocks[headingIndex];
  const sameSectionTable = blocks.find(
    (block, index) =>
      index > headingIndex &&
      block.tagName === "tbl" &&
      !blocks
        .slice(Math.max(0, headingIndex + 1), index)
        .some((item) => item.headingLevel && target?.headingLevel && item.headingLevel <= target.headingLevel)
  );
  return sameSectionTable?.xml ?? blocks.find((block) => block.tagName === "tbl")?.xml;
}

function buildHeadingTemplates(blocks: WordBodyBlock[]): Map<number, string> {
  const templates = new Map<number, string>();
  for (const block of blocks) {
    if (!block.headingLevel || templates.has(block.headingLevel)) continue;
    templates.set(block.headingLevel, block.xml);
  }
  return templates;
}

function buildWordParagraphFromTemplate(templateXml: string, text: string): string {
  const paragraphProperties = templateXml.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/)?.[0] ?? "";
  const firstRun = templateXml.match(/<w:r\b[\s\S]*?<\/w:r>/)?.[0] ?? "";
  const runProperties = firstRun.match(/<w:rPr\b[\s\S]*?<\/w:rPr>/)?.[0] ?? "";
  return buildWordParagraphXml(paragraphProperties, runProperties, text);
}

function buildWordParagraphXml(paragraphProperties: string, runProperties: string, text: string): string {
  return [
    "<w:p>",
    paragraphProperties,
    "<w:r>",
    runProperties,
    `<w:t xml:space="preserve">${escapeXml(text)}</w:t>`,
    "</w:r>",
    "</w:p>"
  ].join("");
}

interface ParsedMarkdownTable {
  rows: string[][];
  nextIndex: number;
}

function parseMarkdownTable(lines: string[], startIndex: number): ParsedMarkdownTable | undefined {
  const header = lines[startIndex]?.trim() ?? "";
  const separator = lines[startIndex + 1]?.trim() ?? "";
  if (!isMarkdownTableRow(header) || !isMarkdownTableSeparatorRow(separator)) return undefined;

  const rows = [splitMarkdownTableRow(header)];
  let nextIndex = startIndex + 2;
  while (nextIndex < lines.length && isMarkdownTableRow(lines[nextIndex]?.trim() ?? "")) {
    const row = splitMarkdownTableRow(lines[nextIndex]);
    if (row.length) rows.push(row);
    nextIndex += 1;
  }

  return rows.length ? { rows: normalizeMarkdownTableRows(rows), nextIndex } : undefined;
}

function isMarkdownTableRow(line: string): boolean {
  return /^\|.*\|$/.test(line.trim()) && splitMarkdownTableRow(line).length > 0;
}

function isMarkdownTableSeparatorRow(line: string): boolean {
  if (!isMarkdownTableRow(line)) return false;
  const cells = splitMarkdownTableRow(line);
  return cells.length > 0 && cells.every((cell) => /^:?-{3,}:?$/.test(cell.trim()));
}

function splitMarkdownTableRow(line: string): string[] {
  const trimmed = line.trim().replace(/^\|/, "").replace(/\|$/, "");
  const cells: string[] = [];
  let current = "";
  let escaped = false;

  for (const char of trimmed) {
    if (escaped) {
      current += char;
      escaped = false;
      continue;
    }
    if (char === "\\") {
      escaped = true;
      continue;
    }
    if (char === "|") {
      cells.push(cleanMarkdownInline(current.trim()));
      current = "";
      continue;
    }
    current += char;
  }
  cells.push(cleanMarkdownInline(current.trim()));
  return cells;
}

function normalizeMarkdownTableRows(rows: string[][]): string[][] {
  const columnCount = Math.max(...rows.map((row) => row.length));
  return rows.map((row) => Array.from({ length: columnCount }, (_, index) => row[index] ?? ""));
}

function buildWordTableFromMarkdown(rows: string[][], paragraphTemplate = "", tableTemplate?: string): string {
  const columnCount = Math.max(1, ...rows.map((row) => row.length));
  const columnWidth = Math.max(900, Math.floor(9000 / columnCount));
  const paragraphProperties = paragraphTemplate.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/)?.[0] ?? "";
  const firstRun = paragraphTemplate.match(/<w:r\b[\s\S]*?<\/w:r>/)?.[0] ?? "";
  const runProperties = firstRun.match(/<w:rPr\b[\s\S]*?<\/w:rPr>/)?.[0] ?? "";
  const tableProperties = tableTemplate?.match(/<w:tblPr\b[\s\S]*?<\/w:tblPr>/)?.[0] ?? buildDefaultTableProperties();
  const grid = [
    "<w:tblGrid>",
    ...Array.from({ length: columnCount }, () => `<w:gridCol w:w="${columnWidth}"/>`),
    "</w:tblGrid>"
  ].join("");
  const rowXml = rows
    .map((row, rowIndex) =>
      [
        "<w:tr>",
        ...Array.from({ length: columnCount }, (_, columnIndex) =>
          buildWordTableCell(row[columnIndex] ?? "", columnWidth, paragraphProperties, runProperties, rowIndex === 0)
        ),
        "</w:tr>"
      ].join("")
    )
    .join("");

  return ["<w:tbl>", tableProperties, grid, rowXml, "</w:tbl>"].join("");
}

function buildDefaultTableProperties(): string {
  return [
    "<w:tblPr>",
    '<w:tblW w:w="0" w:type="auto"/>',
    '<w:tblBorders>',
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>',
    '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>',
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>',
    '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>',
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>',
    '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>',
    '</w:tblBorders>',
    '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>',
    "</w:tblPr>"
  ].join("");
}

function buildWordTableCell(
  text: string,
  width: number,
  paragraphProperties: string,
  runProperties: string,
  header: boolean
): string {
  const cellProperties = [
    "<w:tcPr>",
    `<w:tcW w:w="${width}" w:type="dxa"/>`,
    header ? '<w:shd w:val="clear" w:color="auto" w:fill="D9EAF7"/>' : "",
    "</w:tcPr>"
  ]
    .filter(Boolean)
    .join("");
  const headerRunProperties = header ? mergeRunBold(runProperties) : runProperties;
  return ["<w:tc>", cellProperties, buildWordParagraphXml(paragraphProperties, headerRunProperties, text), "</w:tc>"].join("");
}

function mergeRunBold(runProperties: string): string {
  if (!runProperties) return "<w:rPr><w:b/></w:rPr>";
  if (/<w:b\b/.test(runProperties)) return runProperties;
  return runProperties.replace("</w:rPr>", "<w:b/></w:rPr>");
}

function cleanMarkdownInline(value: string): string {
  return value
    .replace(/^[-*+]\s+/, "· ")
    .replace(/^>\s*/, "引用：")
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .trim();
}

function stripHeadingNumberPrefix(value: string): string {
  return value.replace(/^\d+(?:\.\d+)*[.．、]?\s*/, "").trim();
}

function formatMatchedHeading(block: WordBodyBlock): string {
  return [block.headingNumber ? `${block.headingNumber}.` : "", block.text].filter(Boolean).join(" ").trim();
}

function replaceDocumentBodyWithGeneratedMarkdown(zip: PizZip, markdown: string, facts: SchemeFactModel): boolean {
  const compactedMarkdown = compactText(markdown, 120000);
  if (!compactedMarkdown.trim()) return false;

  const documentFile = zip.file("word/document.xml");
  const documentXml = documentFile?.asText();
  if (!documentXml) return false;

  const bodyOpen = documentXml.match(/<w:body\b[^>]*>/);
  const bodyEnd = documentXml.lastIndexOf("</w:body>");
  if (!bodyOpen || bodyOpen.index === undefined || bodyEnd < 0) return false;

  const bodyStart = bodyOpen.index + bodyOpen[0].length;
  const bodyXml = documentXml.slice(bodyStart, bodyEnd);
  const sectionProperties = bodyXml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/)?.[0] ?? "";
  const title = facts.systemName ? `${facts.systemName}密码应用方案` : "密码应用方案";
  const startsWithHeading = /^#{1,6}\s+/.test(compactedMarkdown.trim());
  const generatedXml = [
    startsWithHeading ? "" : buildWordParagraph(title, { heading: true }),
    ...markdownToWordParagraphs(compactedMarkdown),
    sectionProperties
  ].join("");

  const nextXml = `${documentXml.slice(0, bodyStart)}${generatedXml}${documentXml.slice(bodyEnd)}`;
  zip.file("word/document.xml", nextXml);
  return true;
}

function appendGeneratedMarkdown(zip: PizZip, markdown: string, facts: SchemeFactModel): boolean {
  const compactedMarkdown = compactText(markdown, 70000);
  if (!compactedMarkdown.trim()) return false;

  const factSummary = renderFactSummaryMarkdown(facts);
  const appendixXml = [
    buildWordParagraph("Agent 生成方案正文", { heading: true }),
    ...markdownToWordParagraphs(compactedMarkdown),
    buildWordParagraph("结构化事实摘要", { heading: true }),
    ...markdownToWordParagraphs(factSummary)
  ].join("");
  return insertBeforeDocumentSection(zip, appendixXml);
}

async function appendGeneratedDiagrams(
  zip: PizZip,
  diagrams: SchemeDiagramAsset[],
  templateJson?: WordTemplateJson
): Promise<string[]> {
  const validDiagrams = diagrams.filter((diagram) => diagram.path && diagram.label.trim());
  if (!validDiagrams.length) return [];

  const embedded: string[] = [];
  const remainingDiagrams: SchemeDiagramAsset[] = [];
  const usedFigureIds = new Set<string>();
  const hasTemplateFigures = Boolean(templateJson?.figures?.length);
  const drawingBlocks: string[] = [buildWordParagraph("方案图示", { heading: true })];
  let nextDocPrId = getNextDocPrId(zip);
  let mediaIndex = 0;

  for (const diagram of validDiagrams) {
    const figure = findTemplateFigureForDiagram(templateJson, diagram, usedFigureIds);
    if (!figure) {
      remainingDiagrams.push(diagram);
      continue;
    }

    const media = await tryBuildDiagramMedia(zip, diagram, mediaIndex, nextDocPrId);
    if (!media) {
      remainingDiagrams.push(diagram);
      continue;
    }
    if (!replaceTemplateFigureImage(zip, figure, media)) {
      throw new Error(
        `未找到模板图位锚点：${figure.id}（${figure.caption}）。这通常意味着章节正文替换时覆盖了嵌套的图片 Content Control。`
      );
    }

    nextDocPrId += 1;
    mediaIndex += 1;
    usedFigureIds.add(figure.id);
    embedded.push(diagram.label);
  }

  if (hasTemplateFigures) return embedded;

  for (const diagram of remainingDiagrams) {
    const media = await tryBuildDiagramMedia(zip, diagram, mediaIndex, nextDocPrId);
    if (!media) continue;

    nextDocPrId += 1;
    mediaIndex += 1;
    embedded.push(diagram.label);
    drawingBlocks.push(buildWordParagraph(`图：${diagram.label}`));
    drawingBlocks.push(buildWordImageParagraph(media));
  }

  if (drawingBlocks.length > 1 && !insertBeforeDocumentSection(zip, drawingBlocks.join(""))) return embedded;
  return embedded;
}

function findTemplateFigureForDiagram(
  templateJson: WordTemplateJson | undefined,
  diagram: SchemeDiagramAsset,
  usedFigureIds: Set<string>
): WordTemplateFigure | undefined {
  if (diagram.figureId) {
    const exact = (templateJson?.figures ?? []).find((figure) => figure.id === diagram.figureId);
    if (exact && !usedFigureIds.has(exact.id)) return exact;
  }

  const label = normalizeFigureLookup(diagram.label);
  if (!label) return undefined;

  return (templateJson?.figures ?? []).find((figure) => {
    if (usedFigureIds.has(figure.id)) return false;
    const caption = normalizeFigureLookup(figure.caption);
    return Boolean(caption && (caption === label || caption.includes(label) || label.includes(caption)));
  });
}

function replaceTemplateFigureImage(zip: PizZip, figure: WordTemplateFigure, media: WordImageMedia): boolean {
  const documentFile = zip.file("word/document.xml");
  const documentXml = documentFile?.asText();
  if (!documentXml) return false;

  const bodyOpen = documentXml.match(/<w:body\b[^>]*>/);
  const bodyEnd = documentXml.lastIndexOf("</w:body>");
  if (!bodyOpen || bodyOpen.index === undefined || bodyEnd < 0) return false;

  const bodyStart = bodyOpen.index + bodyOpen[0].length;
  const bodyXml = documentXml.slice(bodyStart, bodyEnd);
  const anchoredImage = findSdtElementByTags(bodyXml, getFigureImageAnchorTags(figure));
  if (anchoredImage) {
    const imageXml = buildWordImageParagraph(media);
    const replacementXml = replaceSdtContentXml(anchoredImage.xml, imageXml);
    const nextBodyXml = `${bodyXml.slice(0, anchoredImage.start)}${replacementXml}${bodyXml.slice(anchoredImage.end)}`;
    const nextXml = `${documentXml.slice(0, bodyStart)}${nextBodyXml}${documentXml.slice(bodyEnd)}`;
    zip.file("word/document.xml", nextXml);
    return true;
  }

  const blocks = collectWordBodyBlocks(bodyXml, buildStyleHeadingLevels(zip));
  const imageBlock = findSdtBlockByTags(blocks, getFigureImageAnchorTags(figure));
  const imageXml = buildWordImageParagraph(media);
  if (!imageBlock) return false;
  const insertStart = imageBlock.start;
  const insertEnd = imageBlock.end;
  const replacementXml = imageBlock.tagName === "sdt" ? replaceSdtContentXml(imageBlock.xml, imageXml) : imageXml;
  const nextBodyXml = `${bodyXml.slice(0, insertStart)}${replacementXml}${bodyXml.slice(insertEnd)}`;
  const nextXml = `${documentXml.slice(0, bodyStart)}${nextBodyXml}${documentXml.slice(bodyEnd)}`;
  zip.file("word/document.xml", nextXml);
  return true;
}

function getFigureImageAnchorTags(figure: WordTemplateFigure): string[] {
  return buildContentControlTagCandidates(figure.anchors?.image, [`ps:figure:${figure.id}:image`]);
}

function getFigureCaptionAnchorTags(figure: WordTemplateFigure): string[] {
  return buildContentControlTagCandidates(figure.anchors?.caption, [`ps:figure:${figure.id}:caption`]);
}

function getFieldBlockAnchorTags(fieldBlock: WordTemplateFieldBlock): string[] {
  return buildContentControlTagCandidates(fieldBlock.anchors?.block, [`ps:field-block:${fieldBlock.id}`]);
}

function getSectionTextBlockAnchorTags(textBlock: WordTemplateTextBlock): string[] {
  return buildContentControlTagCandidates(textBlock.anchors?.block, [`ps:section:${textBlock.section}:text:${textBlock.order}`]);
}

function normalizeFigureLookup(value: string): string {
  return normalizeHeadingLookup(value)
    .replace(/^图\d+/, "")
    .replace(/技术架构/g, "技术框架")
    .replace(/架构/g, "框架")
    .replace(/示意图$/g, "")
    .replace(/流程图$/g, "流程")
    .replace(/图$/g, "");
}

async function tryBuildDiagramMedia(
  zip: PizZip,
  diagram: SchemeDiagramAsset,
  index: number,
  docPrId: number
): Promise<WordImageMedia | undefined> {
  try {
    return await buildDiagramMedia(zip, diagram, index, docPrId);
  } catch {
    return undefined;
  }
}

async function buildDiagramMedia(
  zip: PizZip,
  diagram: SchemeDiagramAsset,
  index: number,
  docPrId: number
): Promise<WordImageMedia | undefined> {
  let extension = normalizeImageExtension(extname(diagram.path));
  let data: Buffer = Buffer.from(await readFile(diagram.path));
  if (extension === "svg") {
    data = await renderSvgToPng(data);
    extension = "png";
  }

  const contentType = getImageContentType(extension);
  if (!contentType) return undefined;

  const mediaName = `diagram-${index + 1}`;
  const mediaPath = `word/media/${mediaName}.${extension}`;
  const relationshipId = ensureImageRelationship(zip, `media/${mediaName}.${extension}`);
  ensureImageContentType(zip, extension, contentType);
  zip.file(mediaPath, data, { binary: true });

  const size = fitImageToDocument(getImageDimensions(extension, data));
  return {
    relationshipId,
    docPrId,
    name: basename(diagram.path),
    label: diagram.label,
    cx: size.cx,
    cy: size.cy
  };
}

async function renderSvgToPng(data: Buffer): Promise<Buffer> {
  const svgText = data.toString("utf-8");
  const dimensions = getSvgDimensions(svgText);
  const image = await loadImage(data);
  const width = Math.max(1, Math.ceil(dimensions?.width || image.width || 1280));
  const height = Math.max(1, Math.ceil(dimensions?.height || image.height || 720));
  const canvas = createCanvas(width, height);
  const context = canvas.getContext("2d");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  context.drawImage(image, 0, 0, width, height);
  return canvas.toBuffer("image/png");
}

interface WordImageMedia {
  relationshipId: string;
  docPrId: number;
  name: string;
  label: string;
  cx: number;
  cy: number;
}

function insertBeforeDocumentSection(zip: PizZip, xml: string): boolean {
  if (!xml.trim()) return false;
  const documentFile = zip.file("word/document.xml");
  const documentXml = documentFile?.asText();
  if (!documentXml) return false;

  const bodyEnd = documentXml.lastIndexOf("</w:body>");
  const bodyOpen = documentXml.match(/<w:body\b[^>]*>/);
  if (!bodyOpen || bodyOpen.index === undefined || bodyEnd < 0) return false;

  const bodyStart = bodyOpen.index + bodyOpen[0].length;
  const bodyXml = documentXml.slice(bodyStart, bodyEnd);
  const finalSectionProperties = bodyXml.match(
    /(?:<w:p\b[\s\S]*?<w:sectPr\b[\s\S]*?<\/w:sectPr>[\s\S]*?<\/w:p>|<w:sectPr\b[\s\S]*?<\/w:sectPr>)\s*$/
  );
  const targetIndex = bodyStart + (finalSectionProperties?.index ?? bodyXml.length);

  const nextXml = `${documentXml.slice(0, targetIndex)}${xml}${documentXml.slice(targetIndex)}`;
  zip.file("word/document.xml", nextXml);
  return true;
}

function markdownToWordParagraphs(markdown: string): string[] {
  const paragraphs: string[] = [];
  let inCodeBlock = false;
  const lines = markdown.split(/\r?\n/);

  for (let lineIndex = 0; lineIndex < lines.length; lineIndex += 1) {
    const rawLine = lines[lineIndex];
    const line = rawLine.trim();
    if (/^```/.test(line)) {
      inCodeBlock = !inCodeBlock;
      continue;
    }
    if (!line) continue;

    if (!inCodeBlock && /^#{1,6}\s+/.test(line)) {
      paragraphs.push(buildWordParagraph(line.replace(/^#{1,6}\s+/, ""), { heading: true }));
      continue;
    }

    const table = !inCodeBlock ? parseMarkdownTable(lines, lineIndex) : undefined;
    if (table) {
      paragraphs.push(buildWordTableFromMarkdown(table.rows));
      lineIndex = table.nextIndex - 1;
      continue;
    }

    const normalized = line
      .replace(/^[-*+]\s+/, "· ")
      .replace(/^>\s*/, "引用：")
      .replace(/\*\*([^*]+)\*\*/g, "$1")
      .replace(/`([^`]+)`/g, "$1");
    paragraphs.push(buildWordParagraph(normalized, { code: inCodeBlock }));
  }

  return paragraphs;
}

function buildWordParagraph(text: string, options: { heading?: boolean; code?: boolean } = {}): string {
  const properties = [
    options.heading ? "<w:pStyle w:val=\"Heading1\"/>" : "",
    options.heading ? "<w:spacing w:before=\"240\" w:after=\"120\"/>" : "",
    options.code ? "<w:spacing w:before=\"80\" w:after=\"80\"/>" : ""
  ]
    .filter(Boolean)
    .join("");
  const runProperties = [
    options.heading ? "<w:b/><w:sz w:val=\"28\"/>" : "",
    options.code ? "<w:rFonts w:ascii=\"Consolas\" w:hAnsi=\"Consolas\"/><w:sz w:val=\"20\"/>" : ""
  ]
    .filter(Boolean)
    .join("");
  const escaped = escapeXml(text);

  return [
    "<w:p>",
    properties ? `<w:pPr>${properties}</w:pPr>` : "",
    "<w:r>",
    runProperties ? `<w:rPr>${runProperties}</w:rPr>` : "",
    `<w:t xml:space="preserve">${escaped}</w:t>`,
    "</w:r>",
    "</w:p>"
  ].join("");
}

function buildWordImageParagraph(media: WordImageMedia): string {
  const label = escapeXml(media.label);
  const name = escapeXml(media.name);
  return [
    "<w:p>",
    '<w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="240"/></w:pPr>',
    "<w:r>",
    "<w:drawing>",
    '<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">',
    `<wp:extent cx="${media.cx}" cy="${media.cy}"/>`,
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
    `<wp:docPr id="${media.docPrId}" name="${label}" descr="${label}"/>`,
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>',
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">',
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">',
    `<pic:nvPicPr><pic:cNvPr id="${media.docPrId}" name="${name}" descr="${label}"/><pic:cNvPicPr/></pic:nvPicPr>`,
    `<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${media.relationshipId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`,
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${media.cx}" cy="${media.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`,
    "</pic:pic>",
    "</a:graphicData>",
    "</a:graphic>",
    "</wp:inline>",
    "</w:drawing>",
    "</w:r>",
    "</w:p>"
  ].join("");
}

function ensureImageRelationship(zip: PizZip, target: string): string {
  const relsPath = "word/_rels/document.xml.rels";
  const defaultRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const relsXml = zip.file(relsPath)?.asText() || defaultRels;
  const relationshipId = `rId${getNextRelationshipNumber(relsXml)}`;
  const relationshipXml = `<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${escapeXml(
    target
  )}"/>`;
  const nextXml = relsXml.includes("</Relationships>")
    ? relsXml.replace("</Relationships>", `${relationshipXml}</Relationships>`)
    : defaultRels.replace("</Relationships>", `${relationshipXml}</Relationships>`);
  zip.file(relsPath, nextXml);
  return relationshipId;
}

function ensureImageContentType(zip: PizZip, extension: string, contentType: string): void {
  const contentTypesPath = "[Content_Types].xml";
  const contentTypesXml = zip.file(contentTypesPath)?.asText();
  if (!contentTypesXml) return;
  const extensionPattern = new RegExp(`<Default\\s+Extension="${escapeRegExp(extension)}"\\s+ContentType=`, "i");
  if (extensionPattern.test(contentTypesXml)) return;

  const defaultXml = `<Default Extension="${escapeXml(extension)}" ContentType="${escapeXml(contentType)}"/>`;
  const nextXml = contentTypesXml.includes("</Types>")
    ? contentTypesXml.replace("</Types>", `${defaultXml}</Types>`)
    : contentTypesXml;
  zip.file(contentTypesPath, nextXml);
}

function getNextRelationshipNumber(relsXml: string): number {
  const matches = Array.from(relsXml.matchAll(/Id="rId(\d+)"/g)).map((match) => Number(match[1]));
  return Math.max(0, ...matches.filter(Number.isFinite)) + 1;
}

function getNextDocPrId(zip: PizZip): number {
  const documentXml = zip.file("word/document.xml")?.asText() || "";
  const matches = Array.from(documentXml.matchAll(/<wp:docPr[^>]*\sid="(\d+)"/g)).map((match) => Number(match[1]));
  return Math.max(0, ...matches.filter(Number.isFinite)) + 1;
}

function normalizeImageExtension(value: string): string {
  const extension = value.replace(/^\./, "").toLowerCase();
  return extension === "jpeg" ? "jpg" : extension;
}

function getImageContentType(extension: string): string {
  const contentTypes: Record<string, string> = {
    png: "image/png",
    jpg: "image/jpeg",
    svg: "image/svg+xml"
  };
  return contentTypes[extension] ?? "";
}

function getImageDimensions(extension: string, data: Buffer): { width: number; height: number } | undefined {
  if (extension === "png" && data.length >= 24 && data.toString("ascii", 1, 4) === "PNG") {
    return {
      width: data.readUInt32BE(16),
      height: data.readUInt32BE(20)
    };
  }

  if (extension === "jpg") {
    return getJpegDimensions(data);
  }

  if (extension === "svg") {
    return getSvgDimensions(data.toString("utf-8"));
  }

  return undefined;
}

function getJpegDimensions(data: Buffer): { width: number; height: number } | undefined {
  let offset = 2;
  while (offset < data.length) {
    if (offset + 4 >= data.length) return undefined;
    if (data[offset] !== 0xff) return undefined;
    const marker = data[offset + 1];
    const length = data.readUInt16BE(offset + 2);
    if (length < 2) return undefined;
    if (marker >= 0xc0 && marker <= 0xc3 && offset + 8 < data.length) {
      return {
        height: data.readUInt16BE(offset + 5),
        width: data.readUInt16BE(offset + 7)
      };
    }
    offset += 2 + length;
  }
  return undefined;
}

function getSvgDimensions(svg: string): { width: number; height: number } | undefined {
  const width = Number(svg.match(/\bwidth="([\d.]+)"/)?.[1]);
  const height = Number(svg.match(/\bheight="([\d.]+)"/)?.[1]);
  if (Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0) {
    return { width, height };
  }

  const viewBox = svg.match(/\bviewBox="[\d.\-]+\s+[\d.\-]+\s+([\d.]+)\s+([\d.]+)"/);
  const viewBoxWidth = Number(viewBox?.[1]);
  const viewBoxHeight = Number(viewBox?.[2]);
  if (Number.isFinite(viewBoxWidth) && Number.isFinite(viewBoxHeight) && viewBoxWidth > 0 && viewBoxHeight > 0) {
    return { width: viewBoxWidth, height: viewBoxHeight };
  }
  return undefined;
}

function fitImageToDocument(dimensions: { width: number; height: number } | undefined): { cx: number; cy: number } {
  const emuPerInch = 914400;
  const maxWidthInches = 6.35;
  const maxHeightInches = 4.2;
  const fallbackRatio = 16 / 9;
  const ratio = dimensions && dimensions.height > 0 ? dimensions.width / dimensions.height : fallbackRatio;
  let widthInches = maxWidthInches;
  let heightInches = widthInches / ratio;
  if (heightInches > maxHeightInches) {
    heightInches = maxHeightInches;
    widthInches = heightInches * ratio;
  }
  return {
    cx: Math.round(widthInches * emuPerInch),
    cy: Math.round(heightInches * emuPerInch)
  };
}

export function renderFactSummaryMarkdown(facts: SchemeFactModel): string {
  const lines = [
    `- 应用系统：${facts.systemName || "待补充"}`,
    `- 建设单位：${facts.organizationName || "待补充"}`,
    `- 所属省份：${facts.province || "待补充"}`,
    `- 单位地址：${facts.address || "待补充"}`,
    `- 邮政编码：${facts.zipCode || "待补充"}`,
    `- 等保级别：${facts.securityLevel || "待补充"}`,
    `- 部署模式：${facts.deploymentMode || "待补充"}`,
    `- 云平台：${facts.cloudPlatform || "未识别/待确认"}`,
    `- 物理机房：${facts.machineRooms.map((room) => room.name).filter(Boolean).join("、") || "待补充"}`,
    `- 应用子系统：${facts.subsystems.join("、") || "待补充"}`,
    `- 关键数据：${facts.keyDataTypes.join("、") || "待补充"}`,
    `- 应用用户：${facts.users.join("、") || "待补充"}`,
    `- 数据库：${facts.databases.join("、") || "待补充"}`,
    `- 密码产品：${facts.cryptoProducts.join("、") || "待补充"}`,
    `- 待补充字段：${facts.missingFields.join("、") || "无"}`
  ];
  return lines.join("\n");
}

function extractWordPackageText(zip: PizZip): string {
  return getTemplateXmlFileNames(zip)
    .map((fileName) => zip.file(fileName)?.asText() ?? "")
    .filter(Boolean)
    .map(extractTextFromXml)
    .join("\n");
}

function extractTextFromXml(xml: string): string {
  return collectTextNodeTokens(xml).map((token) => token.text).join("");
}

function pickValue(source: string, patterns: RegExp[]): string {
  for (const pattern of patterns) {
    const match = source.match(pattern);
    const value = match?.[1]?.trim();
    if (value) return cleanValue(value);
  }
  return "";
}

function cleanValue(value: string): string {
  return value.replace(/[`*_#>|]/g, "").replace(/\s+/g, " ").slice(0, 80).trim();
}

function inferList(source: string, patterns: RegExp[]): string[] {
  for (const pattern of patterns) {
    const match = source.match(pattern);
    if (!match?.[1]) continue;
    const items = splitList(match[1]).map(cleanValue).filter(Boolean);
    if (items.length) return unique(items).slice(0, 8);
  }
  return [];
}

function splitList(value: string): string[] {
  return value
    .replace(/两个子业务板块/g, "")
    .split(/[、,，；;\/\n]/)
    .map((item) => item.trim())
    .filter((item) => item && !/^和|及|等$/.test(item));
}

function inferKeywordList(source: string, keywords: string[]): string[] {
  return keywords.filter((keyword) => new RegExp(escapeRegExp(keyword), "i").test(source));
}

function inferDeploymentMode(source: string): string {
  if (/B\/S|Browser\/Server|浏览器/i.test(source)) return "B/S 架构";
  if (/C\/S|Client\/Server|客户端/i.test(source)) return "C/S 架构";
  if (/IaaS|云服务|云平台/i.test(source)) return "IaaS 云服务模式";
  if (/本地化部署|私有化部署|自建机房/i.test(source)) return "本地化部署";
  return "";
}

function buildMissingFactList(facts: SchemeFactModel): string[] {
  const entries: Array<[string, string | string[]]> = [
    ["应用系统", facts.systemName],
    ["建设单位", facts.organizationName],
    ["单位省份", facts.province],
    ["单位地址", facts.address],
    ["单位邮编", facts.zipCode],
    ["等保级别", facts.securityLevel],
    ["物理机房", facts.machineRooms.map((room) => room.name).filter(Boolean)],
    ["应用子系统", facts.subsystems],
    ["关键数据类型", facts.keyDataTypes],
    ["应用用户", facts.users],
    ["密码产品", facts.cryptoProducts]
  ];

  return entries.filter(([, value]) => (Array.isArray(value) ? value.length === 0 : !value)).map(([label]) => label);
}

function inferProvince(source: string): string {
  const provinces = [
    "北京市",
    "天津市",
    "上海市",
    "重庆市",
    "河北省",
    "山西省",
    "辽宁省",
    "吉林省",
    "黑龙江省",
    "江苏省",
    "浙江省",
    "安徽省",
    "福建省",
    "江西省",
    "山东省",
    "河南省",
    "湖北省",
    "湖南省",
    "广东省",
    "海南省",
    "四川省",
    "贵州省",
    "云南省",
    "陕西省",
    "甘肃省",
    "青海省",
    "台湾省",
    "内蒙古自治区",
    "广西壮族自治区",
    "西藏自治区",
    "宁夏回族自治区",
    "新疆维吾尔自治区",
    "香港特别行政区",
    "澳门特别行政区"
  ];
  return provinces.find((province) => source.includes(province)) ?? "";
}

function normalizeLevel(level: string): string {
  const map: Record<string, string> = {
    "1": "一级",
    "2": "二级",
    "3": "三级",
    "4": "四级",
    一: "一级",
    二: "二级",
    三: "三级",
    四: "四级"
  };
  return map[level] ?? level;
}

function unique(items: string[]): string[] {
  return Array.from(new Set(items));
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&#x([0-9a-f]+);/gi, (_, codePoint: string) => decodeXmlCodePoint(Number.parseInt(codePoint, 16)))
    .replace(/&#(\d+);/g, (_, codePoint: string) => decodeXmlCodePoint(Number.parseInt(codePoint, 10)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&");
}

function decodeXmlCodePoint(codePoint: number): string {
  if (!Number.isFinite(codePoint) || codePoint < 0) return "";
  try {
    return String.fromCodePoint(codePoint);
  } catch {
    return "";
  }
}

function formatChineseDate(date: Date): string {
  return `${date.getFullYear()} 年 ${date.getMonth() + 1} 月 ${date.getDate()} 日`;
}
