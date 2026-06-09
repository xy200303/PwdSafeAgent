import { createHash } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { basename, extname, join } from "node:path";
import PizZip from "pizzip";
import { sanitizeFileName, writeUtf8File } from "./agentTools";
import type { DocumentProfile } from "./documentConfig";

export interface RegisteredDocumentTemplate {
  templateId: string;
  displayName: string;
  templatePath: string;
  templateJsonPath: string;
  profilePath: string;
  templatePathForTool: string;
  templateJsonPathForTool: string;
  profilePathForTool: string;
  sectionCount: number;
  sectionGroupCount: number;
  renderMode: "full_document" | "template_sections";
}

interface GeneratedTemplateJson {
  schemaVersion: 1;
  templateId: string;
  source: {
    docx: string;
    generatedAt: string;
    documentSha256: string;
  };
  strategy: {
    contentSource: string;
    styleSource: string;
    defaultWriteMode: string;
    notes: string[];
  };
  statistics: {
    blockCount: number;
    sectionCount: number;
    tableCount: number;
    placeholderCount: number;
  };
  placeholders: string[];
  sections: GeneratedTemplateSection[];
  tables: GeneratedTemplateTable[];
  blocks: GeneratedTemplateBlock[];
}

interface GeneratedTemplateSection {
  id: string;
  number: string;
  title: string;
  headingLevel: number;
  headingBlock: number;
  bodyRange: [number, number];
  directBodyRange: [number, number];
  childSections: string[];
  placeholders: string[];
  writingHint: string;
  anchors?: {
    body?: {
      tag: string;
      alias?: string;
    };
  };
}

interface GeneratedTemplateTable {
  id: string;
  block: number;
  section?: string;
  sectionNumber?: string;
  rowCount: number;
  columnCount: number;
  header?: string[];
}

interface GeneratedTemplateBlock {
  index: number;
  type: "p" | "tbl";
  text?: string;
  styleId?: string;
  styleName?: string;
  headingLevel?: number;
  section?: string;
  sectionNumber?: string;
  rowCount?: number;
  columnCount?: number;
  header?: string[];
  placeholders?: string[];
}

interface ParsedStyle {
  id: string;
  name: string;
  headingLevel?: number;
}

interface ParsedBlock extends GeneratedTemplateBlock {
  xml: string;
  start: number;
  end: number;
}

export async function registerDocumentTemplateFromDocx(input: {
  sourceDocxPath: string;
  outputDir: string;
  name?: string;
  profile?: string;
}): Promise<RegisteredDocumentTemplate> {
  const sourceBuffer = await readFile(input.sourceDocxPath);
  const hash = createHash("sha256").update(sourceBuffer).digest("hex");
  const baseName = sanitizeFileName(input.name || basename(input.sourceDocxPath, extname(input.sourceDocxPath)));
  const templateId = sanitizeTemplateId(input.profile || `${baseName}-${hash.slice(0, 8)}`);
  const templateDir = join(input.outputDir, "document-templates", templateId);
  const templatePath = join(templateDir, "template.docx");
  const templateJsonPath = join(templateDir, "template.json");
  const profilePath = join(templateDir, "profile.json");
  const templatePathForTool = `document-templates/${templateId}/template.docx`;
  const templateJsonPathForTool = `document-templates/${templateId}/template.json`;
  const profilePathForTool = `document-templates/${templateId}/profile.json`;

  const generated = buildGeneratedTemplateAssets(sourceBuffer, {
    templateId,
    sourceDocx: templatePathForTool,
    documentSha256: hash
  });
  const templateJson = generated.templateJson;
  const renderMode = generated.anchoredSectionCount ? "template_sections" : "full_document";
  const profile = buildProfileFromTemplateJson(templateJson, {
    profile: templateId,
    titleSuffix: baseName,
    templatePath: templatePathForTool,
    templateJsonPath: templateJsonPathForTool,
    renderMode
  });

  await mkdir(templateDir, { recursive: true });
  await writeFile(templatePath, generated.docxBuffer);
  await writeUtf8File(templateJsonPath, `${JSON.stringify(templateJson, null, 2)}\n`);
  await writeUtf8File(profilePath, `${JSON.stringify(profile, null, 2)}\n`);

  return {
    templateId,
    displayName: baseName,
    templatePath,
    templateJsonPath,
    profilePath,
    templatePathForTool,
    templateJsonPathForTool,
    profilePathForTool,
    sectionCount: templateJson.sections.length,
    sectionGroupCount: profile.sectionRules.length,
    renderMode
  };
}

function buildGeneratedTemplateAssets(
  sourceBuffer: Buffer,
  input: { templateId: string; sourceDocx: string; documentSha256: string }
): { templateJson: GeneratedTemplateJson; docxBuffer: Buffer; anchoredSectionCount: number } {
  const zip = new PizZip(sourceBuffer);
  const styles = parseStyles(readZipText(zip, "word/styles.xml", ""));
  const documentXml = readZipText(zip, "word/document.xml");
  const blocks = parseBodyBlocks(documentXml, styles);
  const sections = buildSections(blocks);
  assignSectionsToBlocks(blocks, sections);
  const tables = buildTables(blocks);
  const placeholders = uniqueStrings(blocks.flatMap((block) => block.placeholders ?? []));
  const anchored = insertSectionAnchors(documentXml, blocks, sections);
  if (anchored.documentXml !== documentXml) {
    zip.file("word/document.xml", anchored.documentXml);
  }

  const templateJson: GeneratedTemplateJson = {
    schemaVersion: 1,
    templateId: input.templateId,
    source: {
      docx: input.sourceDocx,
      generatedAt: new Date().toISOString(),
      documentSha256: input.documentSha256
    },
    strategy: {
      contentSource: "user_uploaded_word_template",
      styleSource: "docx",
      defaultWriteMode: anchored.anchoredSectionCount ? "template_sections" : "full_document",
      notes: [
        "该模板由用户上传的 Word 文档自动注册生成。",
        "profile.json 负责文档生成规则；template.json 负责记录从 Word 标题、表格和占位符解析出的模板结构。",
        "注册过程会按 Word 标题结构为可定位的章节正文自动写入隐藏 SDT Content Control 锚点，后续可使用 template_sections 精确替换章节正文。"
      ]
    },
    statistics: {
      blockCount: blocks.length,
      sectionCount: sections.length,
      tableCount: tables.length,
      placeholderCount: placeholders.length
    },
    placeholders,
    sections,
    tables,
    blocks: blocks.map(({ xml, start, end, ...block }) => block)
  };

  return {
    templateJson,
    docxBuffer: zip.generate({ type: "nodebuffer", compression: "DEFLATE" }),
    anchoredSectionCount: anchored.anchoredSectionCount
  };
}

function buildProfileFromTemplateJson(
  templateJson: GeneratedTemplateJson,
  input: {
    profile: string;
    titleSuffix: string;
    templatePath: string;
    templateJsonPath: string;
    renderMode: "full_document" | "template_sections";
  }
): DocumentProfile {
  return {
    version: 1,
    profile: input.profile,
    titleSuffix: input.titleSuffix,
    globalRules: [
      "全文必须基于用户上传 Word 模板的章节结构、标题层级和写作提示组织。",
      "每个章节组先梳理已确认事实和待补充/需确认信息，再按模板章节输出可交付正文。",
      "不得把模板提示、示例资料、行业常识或推断写成项目已经具备的事实。",
      "涉及单位名称、系统名称、设备型号、部署位置、接口、数量、金额、人员和验收结论时，必须来自用户输入、附件读取结果或项目档案已确认事实。",
      "设计、规划、拟建设类内容应使用“拟采用/建议采用/待确认”等表述；只有已确认现状才使用“已部署/已配置/已实现”等确定表述。"
    ],
    evidenceRules: [
      "仅使用用户明确提供、附件工具读到或项目档案已确认的事实。",
      "模板中的示例、占位符和空白表格只能作为待补充提示，不得转写为项目事实。"
    ],
    sectionRules: buildProfileSectionRules(templateJson.sections)
  };
}

function buildProfileSectionRules(sections: GeneratedTemplateSection[]): DocumentProfile["sectionRules"] {
  const grouped = new Map<string, GeneratedTemplateSection[]>();
  for (const section of sections) {
    const root = getTopLevelSectionNumber(section.number);
    const items = grouped.get(root) ?? [];
    items.push(section);
    grouped.set(root, items);
  }

  return Array.from(grouped.entries())
    .sort((left, right) => compareSectionNumber(left[0], right[0]))
    .map(([root, items], index) => {
      const ordered = [...items].sort((left, right) => compareSectionNumber(left.number, right.number));
      const rootSection = ordered.find((section) => section.number === root) ?? ordered[0];
      const placeholders = uniqueStrings(ordered.flatMap((section) => section.placeholders));
      return {
        id: `section_rule_${root.replace(/\./g, "_") || index + 1}`,
        match: rootSection?.title || root || `第 ${index + 1} 章节组`,
        order: Number(root) || index + 1,
        title: rootSection?.title || `第 ${root || index + 1} 章节组`,
        rules: uniqueStrings([
          "只写本章节组对应的已确认事实；资料不足处写待补充/需确认。",
          ...ordered
            .map((section) => section.writingHint)
            .filter((hint): hint is string => Boolean(hint))
            .slice(0, 4)
        ]),
        evidenceRules: [
          "仅使用用户明确提供、附件工具读到或项目档案已确认的事实。",
          "模板中的示例、占位符和空白表格只能作为待补充提示，不得转写为项目事实。"
        ],
        requiredFacts: uniqueStrings([
          ...placeholders.map((placeholder) => `${placeholder}字段取值`),
          rootSection?.title ? `${rootSection.title}相关事实` : ""
        ])
      };
    });
}

function parseStyles(xml: string): ParsedStyle[] {
  const styles: ParsedStyle[] = [];
  for (const styleXml of matchElements(xml, "w:style")) {
    const styleId = getAttr(styleXml.openTag, "w:styleId");
    const type = getAttr(styleXml.openTag, "w:type");
    if (!styleId || type !== "paragraph") continue;
    const name = decodeXml(getSingleAttrElement(styleXml.xml, "w:name", "w:val") ?? styleId).trim();
    const outlineLevel = readOutlineLevel(styleXml.xml);
    const headingLevel = outlineLevel ?? getHeadingLevelFromName(name);
    styles.push({
      id: styleId,
      name,
      ...(headingLevel ? { headingLevel } : {})
    });
  }
  return styles;
}

function parseBodyBlocks(documentXml: string, styles: ParsedStyle[]): ParsedBlock[] {
  const bodyMatch = documentXml.match(/<w:body\b[^>]*>([\s\S]*?)<\/w:body>/);
  if (!bodyMatch) throw new Error("word/document.xml does not contain w:body.");
  const bodyXml = bodyMatch[1] ?? "";
  const styleById = new Map(styles.map((style) => [style.id, style]));
  const blocks: ParsedBlock[] = [];
  let cursor = 0;

  while (cursor < bodyXml.length) {
    const paragraphStart = findNextElementStart(bodyXml, "w:p", cursor);
    const tableStart = findNextElementStart(bodyXml, "w:tbl", cursor);
    const starts = [paragraphStart, tableStart].filter((index) => index >= 0);
    if (!starts.length) break;
    const start = Math.min(...starts);
    const tagName = start === paragraphStart ? "w:p" : "w:tbl";
    const element = readBalancedElement(bodyXml, start, tagName);
    if (!element) break;
    const xml = element.xml;
    const text = extractText(xml);
    const placeholders = extractPlaceholders(text);

    if (tagName === "w:p") {
      const styleId = getSingleAttrElement(xml, "w:pStyle", "w:val") ?? "";
      const style = styleId ? styleById.get(styleId) : undefined;
      const headingLevel = readOutlineLevel(xml) ?? style?.headingLevel ?? undefined;
      blocks.push({
        index: blocks.length,
        type: "p",
        xml,
        start,
        end: element.end,
        ...(text ? { text } : {}),
        ...(styleId ? { styleId } : {}),
        ...(style?.name ? { styleName: style.name } : {}),
        ...(headingLevel ? { headingLevel } : {}),
        ...(placeholders.length ? { placeholders } : {})
      });
    } else {
      const rows = parseTableRows(xml);
      const header = rows[0]?.filter(Boolean);
      blocks.push({
        index: blocks.length,
        type: "tbl",
        xml,
        start,
        end: element.end,
        ...(text ? { text } : {}),
        rowCount: rows.length,
        columnCount: Math.max(0, ...rows.map((row) => row.length)),
        ...(header?.length ? { header } : {}),
        ...(placeholders.length ? { placeholders } : {})
      });
    }

    cursor = element.end;
  }

  return blocks;
}

function buildSections(blocks: ParsedBlock[]): GeneratedTemplateSection[] {
  const headings = blocks.filter((block) => block.type === "p" && block.headingLevel && block.text);
  const counters: number[] = [];
  const sections: GeneratedTemplateSection[] = [];
  for (const block of headings) {
    const headingLevel = block.headingLevel || 1;
    const parsedNumber = parseHeadingNumber(block.text || "");
    const number = parsedNumber || nextGeneratedSectionNumber(counters, headingLevel);
    const title = stripHeadingNumberPrefix(block.text || "") || `第 ${number} 节`;
    sections.push({
      id: uniqueSectionId(`sec_${number.replace(/\./g, "_")}`),
      number,
      title,
      headingLevel,
      headingBlock: block.index,
      bodyRange: [block.index + 1, block.index + 1] as [number, number],
      directBodyRange: [block.index + 1, block.index + 1] as [number, number],
      childSections: [],
      placeholders: [],
      writingHint: `围绕“${title}”编写与当前文档目标相关的正文；资料不足处写待补充/需确认。`
    });
  }

  if (!sections.length) {
    sections.push({
      id: "sec_1",
      number: "1",
      title: "正文",
      headingLevel: 1,
      headingBlock: 0,
      bodyRange: [0, blocks.length],
      directBodyRange: [0, blocks.length],
      childSections: [],
      placeholders: uniqueStrings(blocks.flatMap((block) => block.placeholders ?? [])),
      writingHint: "根据用户输入材料和模板正文结构编写可交付正文；资料不足处写待补充/需确认。"
    });
    return sections;
  }

  for (let index = 0; index < sections.length; index += 1) {
    const section = sections[index];
    const nextSiblingOrParent = sections
      .slice(index + 1)
      .find((candidate) => candidate.headingLevel <= section.headingLevel);
    const nextAnyChild = sections.slice(index + 1).find((candidate) => candidate.headingLevel > section.headingLevel);
    const end = nextSiblingOrParent?.headingBlock ?? blocks.length;
    const directEnd = nextAnyChild && nextAnyChild.headingBlock < end ? nextAnyChild.headingBlock : end;
    section.bodyRange = [section.headingBlock + 1, end];
    section.directBodyRange = [section.headingBlock + 1, directEnd];
    section.placeholders = uniqueStrings(blocks.slice(section.bodyRange[0], section.bodyRange[1]).flatMap((block) => block.placeholders ?? []));
    section.childSections = sections
      .filter((candidate) => candidate.number.startsWith(`${section.number}.`) && candidate.number !== section.number)
      .map((candidate) => candidate.id);
  }

  return sections;

  function uniqueSectionId(baseId: string): string {
    let id = baseId || "sec_1";
    let suffix = 2;
    while (sections.some((section) => section.id === id)) {
      id = `${baseId}_${suffix}`;
      suffix += 1;
    }
    return id;
  }
}

function assignSectionsToBlocks(blocks: ParsedBlock[], sections: GeneratedTemplateSection[]): void {
  for (const block of blocks) {
    const section = [...sections]
      .reverse()
      .find((candidate) => block.index >= candidate.headingBlock && block.index < candidate.bodyRange[1]);
    if (!section) continue;
    block.section = section.id;
    block.sectionNumber = section.number;
  }
}

function buildTables(blocks: ParsedBlock[]): GeneratedTemplateTable[] {
  return blocks
    .filter((block) => block.type === "tbl")
    .map((block, index) => ({
      id: `table_${index + 1}${block.sectionNumber ? `_${block.sectionNumber.replace(/\./g, "_")}` : ""}`,
      block: block.index,
      ...(block.section ? { section: block.section } : {}),
      ...(block.sectionNumber ? { sectionNumber: block.sectionNumber } : {}),
      rowCount: block.rowCount ?? 0,
      columnCount: block.columnCount ?? 0,
      ...(block.header?.length ? { header: block.header } : {})
    }));
}

function insertSectionAnchors(
  documentXml: string,
  blocks: ParsedBlock[],
  sections: GeneratedTemplateSection[]
): { documentXml: string; anchoredSectionCount: number } {
  const bodyOpen = documentXml.match(/<w:body\b[^>]*>/);
  const bodyEnd = documentXml.lastIndexOf("</w:body>");
  if (!bodyOpen || bodyOpen.index === undefined || bodyEnd < 0) {
    return { documentXml, anchoredSectionCount: 0 };
  }

  const bodyStart = bodyOpen.index + bodyOpen[0].length;
  let bodyXml = documentXml.slice(bodyStart, bodyEnd);
  const replacements: Array<{ start: number; end: number; xml: string }> = [];

  for (const section of sections) {
    const [startIndex, endIndex] = section.directBodyRange;
    if (startIndex >= endIndex) continue;
    const firstBlock = blocks[startIndex];
    const lastBlock = blocks[endIndex - 1];
    if (!firstBlock || !lastBlock || firstBlock.start >= lastBlock.end) continue;
    const tag = `ps:section:${section.id}:body`;
    const alias = `${section.number} ${section.title} 正文`;
    section.anchors = {
      body: {
        tag,
        alias
      }
    };
    const contentXml = bodyXml.slice(firstBlock.start, lastBlock.end);
    replacements.push({
      start: firstBlock.start,
      end: lastBlock.end,
      xml: wrapBlockContentControl(contentXml, tag, alias)
    });
  }

  for (const replacement of replacements.sort((left, right) => right.start - left.start)) {
    bodyXml = `${bodyXml.slice(0, replacement.start)}${replacement.xml}${bodyXml.slice(replacement.end)}`;
  }

  return {
    documentXml: `${documentXml.slice(0, bodyStart)}${bodyXml}${documentXml.slice(bodyEnd)}`,
    anchoredSectionCount: replacements.length
  };
}

function wrapBlockContentControl(contentXml: string, tag: string, alias: string): string {
  return [
    "<w:sdt>",
    "<w:sdtPr>",
    `<w:alias w:val="${escapeXmlAttr(alias)}"/>`,
    `<w:tag w:val="${escapeXmlAttr(tag)}"/>`,
    "</w:sdtPr>",
    "<w:sdtContent>",
    contentXml,
    "</w:sdtContent>",
    "</w:sdt>"
  ].join("");
}

function readZipText(zip: PizZip, path: string, fallback?: string): string {
  const file = zip.file(path);
  if (!file) {
    if (fallback !== undefined) return fallback;
    throw new Error(`Missing ${path} in Word template.`);
  }
  return file.asText();
}

function getHeadingLevelFromName(name: string): number | undefined {
  const match = name.match(/(?:heading|标题)\s*([1-9])/i);
  if (!match?.[1]) return undefined;
  return Number(match[1]);
}

function readOutlineLevel(xml: string): number | undefined {
  const value = getSingleAttrElement(xml, "w:outlineLvl", "w:val");
  if (value === undefined) return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed + 1 : undefined;
}

function parseHeadingNumber(text: string): string | undefined {
  return text.trim().match(/^(\d+(?:\.\d+)*)(?:[\.．、\s]+|$)/)?.[1];
}

function nextGeneratedSectionNumber(counters: number[], headingLevel: number): string {
  const level = Math.max(1, Math.min(headingLevel, 9));
  counters[level - 1] = (counters[level - 1] || 0) + 1;
  counters.length = level;
  for (let index = 0; index < level; index += 1) {
    if (!counters[index]) counters[index] = 1;
  }
  return counters.join(".");
}

function stripHeadingNumberPrefix(text: string): string {
  return text.replace(/^\s*\d+(?:\.\d+)*[\.．、\s]*/, "").trim();
}

function parseTableRows(xml: string): string[][] {
  return Array.from(xml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)).map((rowMatch) =>
    Array.from(rowMatch[0].matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)).map((cellMatch) => extractText(cellMatch[0]))
  );
}

function extractText(xml: string): string {
  return Array.from(xml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g))
    .map((match) => decodeXml(match[1] ?? ""))
    .join("")
    .replace(/\s+/g, " ")
    .trim();
}

function extractPlaceholders(text: string): string[] {
  return uniqueStrings(Array.from(text.matchAll(/[{\uFF5B]([^{}\uFF5B\uFF5D]{1,80})[}\uFF5D]/g)).map((match) => match[1]));
}

function compareSectionNumber(left: string, right: string): number {
  const leftParts = left.split(".").map((part) => Number(part) || 0);
  const rightParts = right.split(".").map((part) => Number(part) || 0);
  const length = Math.max(leftParts.length, rightParts.length);
  for (let index = 0; index < length; index += 1) {
    const diff = (leftParts[index] || 0) - (rightParts[index] || 0);
    if (diff) return diff;
  }
  return 0;
}

function getTopLevelSectionNumber(number: string): string {
  return number.split(".")[0] || number || "1";
}

function sanitizeTemplateId(value: string): string {
  return sanitizeFileName(value).replace(/\s+/g, "_").replace(/^\.+|\.+$/g, "") || "document-template";
}

function uniqueStrings(items: Array<string | undefined>): string[] {
  const seen = new Set<string>();
  const unique: string[] = [];
  for (const raw of items) {
    const item = String(raw || "").trim();
    if (!item || seen.has(item)) continue;
    seen.add(item);
    unique.push(item);
  }
  return unique;
}

function getSingleAttrElement(xml: string, elementName: string, attrName: string): string | undefined {
  const escapedElement = elementName.replace(":", "\\:");
  const element = xml.match(new RegExp(`<${escapedElement}\\b([^>]*)\\/?>`, "i"));
  return element?.[1] ? getAttr(element[1], attrName) : undefined;
}

function getAttr(attrs: string, name: string): string | undefined {
  const escapedName = name.replace(":", "\\:");
  const match = attrs.match(new RegExp(`${escapedName}=\"([^\"]*)\"`, "i"));
  return match ? decodeXml(match[1] ?? "") : undefined;
}

function matchElements(xml: string, tagName: string): Array<{ xml: string; openTag: string }> {
  const escapedTag = tagName.replace(":", "\\:");
  const regex = new RegExp(`<${escapedTag}\\b[^>]*>[\\s\\S]*?<\\/${escapedTag}>`, "gi");
  return Array.from(xml.matchAll(regex)).map((match) => {
    const elementXml = match[0];
    return {
      xml: elementXml,
      openTag: elementXml.match(new RegExp(`^<${escapedTag}\\b[^>]*>`, "i"))?.[0] ?? ""
    };
  });
}

function findNextElementStart(xml: string, tagName: string, start: number): number {
  const escapedTag = tagName.replace(":", "\\:");
  const regex = new RegExp(`<${escapedTag}\\b`, "g");
  regex.lastIndex = start;
  const match = regex.exec(xml);
  return match?.index ?? -1;
}

function readBalancedElement(xml: string, start: number, tagName: string): { xml: string; end: number } | undefined {
  const openEnd = xml.indexOf(">", start);
  if (openEnd < 0) return undefined;
  const openTag = xml.slice(start, openEnd + 1);
  if (/\/>$/.test(openTag)) return { xml: openTag, end: openEnd + 1 };

  const escapedTag = tagName.replace(":", "\\:");
  const regex = new RegExp(`<\\/?${escapedTag}\\b[^>]*>`, "g");
  regex.lastIndex = start;
  let depth = 0;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(xml))) {
    const tag = match[0];
    const isClosing = tag.startsWith("</");
    const isSelfClosing = /\/>$/.test(tag);
    if (!isClosing && !isSelfClosing) depth += 1;
    if (isClosing) depth -= 1;
    if (depth === 0) return { xml: xml.slice(start, regex.lastIndex), end: regex.lastIndex };
  }
  return undefined;
}

function decodeXml(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function escapeXmlAttr(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
