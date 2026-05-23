import { mkdir, readFile, writeFile } from "node:fs/promises";
import { basename, dirname, extname } from "node:path";
import PizZip from "pizzip";
import { compactText, getCurrentTimeText, sanitizeFileName } from "./agentTools";

export interface SchemeDocumentInput {
  prompt: string;
  memory: string;
  generatedMarkdown: string;
  fields?: SchemeTemplateFieldInput;
  templateFields?: SchemeTemplateFieldInput;
  diagrams?: SchemeDiagramAsset[];
}

export type SchemeTemplateFieldInput = Record<string, unknown> | Array<Record<string, unknown>>;

export interface SchemeDiagramAsset {
  label: string;
  kind?: string;
  path: string;
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

export async function writeSchemeDocxFromTemplate(
  templatePath: string,
  outputPath: string,
  input: SchemeDocumentInput
): Promise<SchemeDocumentResult> {
  const facts = extractSchemeFacts(input);
  const data = buildSchemeTemplateData(input);
  const content = await readFile(templatePath, "binary");
  const zip = new PizZip(content);
  const templateReplacementCount = replaceTemplatePlaceholders(zip, data);
  const renderedZip = zip;
  const appendedMarkdown = appendGeneratedMarkdown(renderedZip, input.generatedMarkdown, facts);
  const embeddedDiagrams = await appendGeneratedDiagrams(renderedZip, input.diagrams ?? []);
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
    templateReplacementCount
  };
}

export function replaceTemplatePlaceholders(zip: PizZip, data: TemplateData): number {
  const replacements = buildPlaceholderReplacements(data);
  if (!replacements.length) return 0;

  let replacementCount = 0;
  for (const fileName of getTemplateXmlFileNames(zip)) {
    const file = zip.file(fileName);
    const xml = file?.asText();
    if (!xml) continue;

    const result = replaceTextNodePlaceholders(xml, replacements);
    if (result.count > 0) {
      zip.file(fileName, result.xml);
      replacementCount += result.count;
    }
  }

  return replacementCount;
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
  let key = rawKey.trim();
  if (key.startsWith("${") && key.endsWith("}")) {
    key = key.slice(2, -1);
  } else if (key.startsWith("{") && key.endsWith("}")) {
    key = key.slice(1, -1);
  }
  return key.replace(/^\$/, "").trim();
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
  data[`$${key}`] = value;
  data[key.replace(/\s+/g, "")] = value;
  data[`$${key.replace(/\s+/g, "")}`] = value;
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

function buildPlaceholderReplacements(data: TemplateData): PlaceholderReplacement[] {
  const replacements = new Map<string, string>();
  for (const [key, value] of Object.entries(data)) {
    if (!key || key.startsWith("$")) continue;
    if (key === "方案生成正文" || key === "结构化事实摘要" || key === "当前时间") continue;
    replacements.set(`{${key}}`, value);
    replacements.set(`\${${key}}`, value);
  }

  return Array.from(replacements, ([placeholder, value]) => ({ placeholder, value })).sort(
    (left, right) => right.placeholder.length - left.placeholder.length
  );
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

async function appendGeneratedDiagrams(zip: PizZip, diagrams: SchemeDiagramAsset[]): Promise<string[]> {
  const validDiagrams = diagrams.filter((diagram) => diagram.path && diagram.label.trim());
  if (!validDiagrams.length) return [];

  const embedded: string[] = [];
  const drawingBlocks: string[] = [buildWordParagraph("方案图示", { heading: true })];
  let nextDocPrId = getNextDocPrId(zip);

  for (const [index, diagram] of validDiagrams.entries()) {
    const media = await tryBuildDiagramMedia(zip, diagram, index, nextDocPrId);
    if (!media) continue;

    nextDocPrId += 1;
    embedded.push(diagram.label);
    drawingBlocks.push(buildWordParagraph(`图：${diagram.label}`));
    drawingBlocks.push(buildWordImageParagraph(media));
  }

  if (!embedded.length) return [];
  if (!insertBeforeDocumentSection(zip, drawingBlocks.join(""))) return [];
  return embedded;
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
  const extension = normalizeImageExtension(extname(diagram.path));
  const contentType = getImageContentType(extension);
  if (!contentType) return undefined;

  const data = await readFile(diagram.path);
  const mediaName = sanitizeFileName(`diagram-${index + 1}-${diagram.label}`).replace(/\.+$/g, "") || `diagram-${index + 1}`;
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

  const insertAt = documentXml.lastIndexOf("<w:sectPr");
  const bodyEnd = documentXml.lastIndexOf("</w:body>");
  const targetIndex = insertAt >= 0 ? insertAt : bodyEnd;
  if (targetIndex < 0) return false;

  const nextXml = `${documentXml.slice(0, targetIndex)}${xml}${documentXml.slice(targetIndex)}`;
  zip.file("word/document.xml", nextXml);
  return true;
}

function markdownToWordParagraphs(markdown: string): string[] {
  const paragraphs: string[] = [];
  let inCodeBlock = false;

  for (const rawLine of markdown.split(/\r?\n/)) {
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

    if (!inCodeBlock && /^\|.*\|$/.test(line)) {
      const cells = line
        .split("|")
        .map((cell) => cell.trim())
        .filter(Boolean);
      if (cells.every((cell) => /^:?-{3,}:?$/.test(cell))) continue;
      paragraphs.push(buildWordParagraph(cells.join("    ")));
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
