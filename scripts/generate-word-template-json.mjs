import { createHash } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, join, relative } from "node:path";
import process from "node:process";
import PizZip from "pizzip";

const ROOT = process.cwd();
const DEFAULT_DOCX = join(ROOT, "docs", "密码应用方案.docx");
const DEFAULT_OUTPUT = join(ROOT, "docs", "密码应用方案.template.json");
const SECTION_BODY_PLACEHOLDER_TEXT = "【正文占位】";
const TABLE_CELL_PLACEHOLDER_TEXT = "【待填写】";
const FIGURE_PLACEHOLDER_TEXT = "【图片占位】";

const [, , docxArg, outputArg] = process.argv;
const docxPath = docxArg ? join(ROOT, docxArg) : DEFAULT_DOCX;
const outputPath = outputArg ? join(ROOT, outputArg) : DEFAULT_OUTPUT;

let content = await readFile(docxPath);
const zip = new PizZip(content);
const originalDocumentXml = readZipText(zip, "word/document.xml");
const documentXml = normalizeTemplateFieldMarkers(
  stripManagedTemplateVisibleMarkers(stripManagedTemplateBookmarks(stripManagedTemplateSdts(originalDocumentXml)))
);
const stylesXml = readZipText(zip, "word/styles.xml");
const relsXml = readZipText(zip, "word/_rels/document.xml.rels", "");

let model = buildTemplateModel(documentXml);
let markedDocumentXml = model.markedDocumentXml;
const cleanedDocumentXml = normalizeTemplateFieldMarkers(
  stripManagedTemplateVisibleMarkers(stripManagedTemplateBookmarks(stripManagedTemplateSdts(markedDocumentXml)))
);
if (cleanedDocumentXml !== documentXml) {
  model = buildTemplateModel(cleanedDocumentXml);
  markedDocumentXml = model.markedDocumentXml;
}
const { styles, blocks, sections, tables, figures, placeholders, fieldAnchors, drawingCount } = model;

if (markedDocumentXml !== originalDocumentXml) {
  zip.file("word/document.xml", markedDocumentXml);
  pruneUnusedDocumentImageMedia(zip, markedDocumentXml);
  content = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
  await writeFile(docxPath, content);
}

const templateJson = {
  schemaVersion: 1,
  templateId: "password-application-scheme-v1",
  source: {
    docx: toProjectPath(docxPath),
    generatedAt: new Date().toISOString(),
    documentSha256: createHash("sha256").update(content).digest("hex")
  },
  strategy: {
    contentSource: "structured_agent_content",
    styleSource: "docx",
    defaultWriteMode: "invisible_anchor_sections",
    workflow: [
      "Agent 根据 sections、tables、figures 的语义信息组织结构化内容，template.json 不作为 Word 重建模板。",
      "运行时只用 anchors 中的 Word Content Control / SDT tag 做机器定位，不依赖章节标题、编号或正文关键词搜索。",
      "章节正文替换只更新目标锚的 w:sdtContent，外层不可见锚保留，便于后续继续增量替换。",
      "正文段落、表格行和图片槽位从 docs/密码应用方案.docx 的原始块克隆，样式、编号、缩进、边框和题注以 docx 为准。",
      "简单字段统一使用 docx-templates 的 {字段名} 短文本占位，不再保留 ${字段名} 和 [[PS:field:字段]] 兼容。",
      "表格和图片通过 ps:table:*、ps:figure:* 不可见锚定位；标题、编号、caption 和 recommendedLabel 只用于生成提示、匹配建议和异常校验。"
    ],
    notes: [
      "anchors 是写入 Word 的不可见 SDT tag，是章节、表格、图片的主定位依据。",
      "headingBlock、bodyRange、captionBlock 是模板结构快照，仅用于分析、校验和回归对比，不作为运行时主定位。",
      "sections.number/title/writingHint 用于 Agent 写作提示和人工审阅，不能作为可靠定位条件。",
      "fieldAnchors 只记录 {字段名} 短文本占位；复杂正文、表格、图片不得通过字段占位承载。",
      "渲染器应保留外层 w:sdt，仅替换 w:sdtContent 内部块，避免丢失不可见锚。",
      "模板源格式必须是 .docx；旧 .doc 二进制格式不支持该不可见锚和 OOXML 克隆渲染方案。"
    ]
  },
  fieldGuide: buildFieldGuide(),
  statistics: {
    blockCount: blocks.length,
    sectionCount: sections.length,
    tableCount: tables.length,
    figureCount: figures.length,
    drawingCount,
    placeholderCount: placeholders.length
  },
  placeholders,
  fieldAnchors,
  styles: styles
    .filter((style) => style.headingLevel || isUsefulStyleName(style.name))
    .map(({ id, name, headingLevel }) => ({ id, name, ...(headingLevel ? { headingLevel } : {}) })),
  sections,
  tables,
  figures,
  blocks: blocks.map(stripInternalBlockFields)
};

await mkdir(dirname(outputPath), { recursive: true });
await writeFile(outputPath, `${JSON.stringify(templateJson, null, 2)}\n`, "utf-8");

console.log(
  `Generated ${toProjectPath(outputPath)}: ${sections.length} sections, ${tables.length} tables, ${figures.length} figures, ${placeholders.length} placeholders.`
);

function readZipText(zipFile, path, fallback = undefined) {
  const file = zipFile.file(path);
  if (!file) {
    if (fallback !== undefined) return fallback;
    throw new Error(`Missing ${path} in ${toProjectPath(docxPath)}`);
  }
  return file.asText();
}

function buildTemplateModel(sourceDocumentXml) {
  const styles = parseStyles(stylesXml);
  const relationships = parseRelationships(relsXml);
  const blocks = parseBodyBlocks(sourceDocumentXml, styles, relationships);
  const sections = buildSections(blocks);
  assignBlockSections(blocks, sections);

  const tables = buildTables(blocks, sections);
  const figures = buildFigures(blocks, sections, relationships);
  addTemplateInvisibleAnchors(blocks, sections, tables, figures);
  applyTemplateGuidance(sections, tables, figures);

  return {
    styles,
    relationships,
    blocks,
    sections,
    tables,
    figures,
    placeholders: uniqueSorted(blocks.flatMap((block) => block.placeholders ?? [])),
    fieldAnchors: buildFieldAnchors(blocks),
    drawingCount: blocks.reduce((sum, block) => sum + (block.drawingCount ?? 0), 0),
    markedDocumentXml: cleanTemplateSourceText(insertTemplateInvisibleAnchors(sourceDocumentXml, blocks, sections, tables, figures))
  };
}

function parseStyles(xml) {
  const stylesById = new Map();
  for (const styleXml of matchElements(xml, "w:style")) {
    const styleId = getAttr(styleXml.openTag, "w:styleId");
    const type = getAttr(styleXml.openTag, "w:type");
    if (!styleId || type !== "paragraph") continue;

    const name = decodeXml(getSingleAttrElement(styleXml.xml, "w:name", "w:val") ?? styleId).trim();
    const headingLevel = getHeadingLevel(name);
    stylesById.set(styleId, {
      id: styleId,
      name,
      ...(headingLevel ? { headingLevel } : {})
    });
  }
  return Array.from(stylesById.values());
}

function parseRelationships(xml) {
  const relationships = new Map();
  for (const relationship of xml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const attrs = relationship[1] ?? "";
    const id = getAttr(attrs, "Id");
    const target = getAttr(attrs, "Target");
    const type = getAttr(attrs, "Type");
    if (!id) continue;
    relationships.set(id, { id, target, type });
  }
  return relationships;
}

function parseBodyBlocks(xml, styles, relationships) {
  const bodyMatch = xml.match(/<w:body\b[^>]*>([\s\S]*?)<\/w:body>/);
  if (!bodyMatch) throw new Error("word/document.xml does not contain w:body.");

  const styleById = new Map(styles.map((style) => [style.id, style]));
  const bodyXml = bodyMatch[1];
  const blocks = [];
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

    const block =
      tagName === "w:p"
        ? parseParagraphBlock(element.xml, blocks.length, styleById, relationships)
        : parseTableBlock(element.xml, blocks.length);
    block._start = start;
    block._end = element.end;
    blocks.push(block);
    cursor = element.end;
  }

  return blocks;
}

function parseParagraphBlock(xml, index, styleById, relationships) {
  const styleId = getSingleAttrElement(xml, "w:pStyle", "w:val");
  const style = styleId ? styleById.get(styleId) : undefined;
  const text = extractText(xml);
  const relIds = extractDrawingRelationshipIds(xml);
  const captionType = getCaptionType(text, style);
  const block = {
    index,
    type: "p",
    ...(styleId ? { styleId } : {}),
    ...(style?.name ? { styleName: style.name } : {}),
    ...(style?.headingLevel ? { headingLevel: style.headingLevel } : {}),
    ...(text ? { text } : {}),
    ...(extractPlaceholders(text).length ? { placeholders: extractPlaceholders(text) } : {}),
    ...(relIds.length ? { drawingRelationshipIds: relIds } : {}),
    ...(relIds.length ? { imageTargets: relIds.map((id) => relationships.get(id)?.target).filter(Boolean) } : {}),
    ...(getDrawingCount(xml) ? { drawingCount: getDrawingCount(xml) } : {}),
    ...(captionType ? { captionType } : {}),
    _xml: xml
  };
  return block;
}

function parseTableBlock(xml, index) {
  const rows = parseTableRows(xml);
  const text = rows.map((row) => row.cells.map((cell) => cell.text).join("")).join("");
  const placeholders = uniqueSorted(rows.flatMap((row) => row.cells.flatMap((cell) => cell.placeholders ?? [])));
  const columnCount = rows.reduce((max, row) => Math.max(max, row.logicalColumnCount), 0);

  return {
    index,
    type: "tbl",
    ...(text ? { text } : {}),
    ...(placeholders.length ? { placeholders } : {}),
    table: {
      rowCount: rows.length,
      columnCount,
      header: inferTableHeader(rows),
      rows: rows.map(({ logicalColumnCount, ...row }) => row)
    },
    _xml: xml
  };
}

function parseTableRows(tableXml) {
  const rows = [];
  let cursor = 0;

  while (cursor < tableXml.length) {
    const rowStart = findNextElementStart(tableXml, "w:tr", cursor);
    if (rowStart < 0) break;
    const rowElement = readBalancedElement(tableXml, rowStart, "w:tr");
    if (!rowElement) break;

    const cells = parseTableCells(rowElement.xml, rows.length);
    rows.push({
      index: rows.length,
      cells,
      logicalColumnCount: cells.reduce((max, cell) => Math.max(max, cell.columnIndex + cell.columnSpan), 0)
    });
    cursor = rowElement.end;
  }

  return rows;
}

function parseTableCells(rowXml, rowIndex) {
  const cells = [];
  let cursor = 0;
  let logicalColumn = 0;

  while (cursor < rowXml.length) {
    const cellStart = findNextElementStart(rowXml, "w:tc", cursor);
    if (cellStart < 0) break;
    const cellElement = readBalancedElement(rowXml, cellStart, "w:tc");
    if (!cellElement) break;

    const text = extractText(cellElement.xml);
    const placeholders = extractPlaceholders(text);
    const columnSpan = Number.parseInt(getSingleAttrElement(cellElement.xml, "w:gridSpan", "w:val") ?? "1", 10) || 1;
    const verticalMerge = getSingleAttrElement(cellElement.xml, "w:vMerge", "w:val") ?? (cellElement.xml.includes("<w:vMerge") ? "continue" : undefined);
    const paragraphs = matchElements(cellElement.xml, "w:p")
      .map((paragraph, paragraphIndex) => ({
        index: paragraphIndex,
        text: extractText(paragraph.xml),
        placeholders: extractPlaceholders(extractText(paragraph.xml))
      }))
      .filter((paragraph) => paragraph.text || paragraph.placeholders.length);

    cells.push({
      rowIndex,
      cellIndex: cells.length,
      columnIndex: logicalColumn,
      columnSpan,
      ...(verticalMerge ? { verticalMerge } : {}),
      text,
      ...(placeholders.length ? { placeholders } : {}),
      ...(paragraphs.length ? { paragraphs } : {})
    });

    logicalColumn += columnSpan;
    cursor = cellElement.end;
  }

  return cells;
}

function buildSections(blocks) {
  const headingBlocks = blocks.filter((block) => block.type === "p" && block.headingLevel && block.text);
  const counters = [];
  const sections = [];
  const stack = [];

  for (const heading of headingBlocks) {
    const level = heading.headingLevel;
    counters[level - 1] = (counters[level - 1] ?? 0) + 1;
    counters.length = level;

    const number = counters.join(".");
    const id = `sec_${number.replaceAll(".", "_")}`;

    while (stack.length && stack.at(-1).headingLevel >= level) {
      stack.pop();
    }

    const section = {
      id,
      number,
      title: stripHeadingNumberPrefix(heading.text),
      headingLevel: level,
      headingBlock: heading.index,
      bodyRange: [heading.index + 1, heading.index],
      childSections: [],
      placeholders: [],
      mode: "replace_body",
      _parentId: stack.at(-1)?.id
    };

    const parent = stack.at(-1);
    if (parent) parent.childSections.push(id);

    sections.push(section);
    stack.push(section);
  }

  for (let index = 0; index < sections.length; index += 1) {
    const section = sections[index];
    const nextSameOrHigher = sections
      .slice(index + 1)
      .find((candidate) => candidate.headingLevel <= section.headingLevel);
    const bodyEnd = nextSameOrHigher ? nextSameOrHigher.headingBlock - 1 : blocks.length - 1;
    section.bodyRange = [section.headingBlock + 1, Math.max(section.headingBlock, bodyEnd)];
    section.placeholders = uniqueSorted(
      blocks
        .slice(section.headingBlock + 1, bodyEnd + 1)
        .flatMap((block) => block.placeholders ?? [])
    );
    delete section._parentId;
  }

  return sections;
}

function assignBlockSections(blocks, sections) {
  const sortedSections = [...sections].sort((a, b) => a.headingBlock - b.headingBlock);
  let currentSectionIndex = -1;

  for (const block of blocks) {
    while (
      currentSectionIndex + 1 < sortedSections.length &&
      sortedSections[currentSectionIndex + 1].headingBlock <= block.index
    ) {
      currentSectionIndex += 1;
    }

    const section = sortedSections[currentSectionIndex];
    if (section) {
      block.section = section.id;
      block.sectionNumber = section.number;
    }
  }
}

function buildTables(blocks) {
  const tables = [];

  for (const block of blocks) {
    if (block.type !== "tbl") continue;

    const captionBlock = findAdjacentCaptionBlock(blocks, block.index, "table");
    const id = `table_${tables.length + 1}${block.sectionNumber ? `_${block.sectionNumber.replaceAll(".", "_")}` : "_front"}`;
    tables.push({
      id,
      type: "table",
      ...(block.section ? { section: block.section } : {}),
      ...(block.sectionNumber ? { sectionNumber: block.sectionNumber } : {}),
      block: block.index,
      ...(captionBlock ? { captionBlock: captionBlock.index, caption: captionBlock.text } : {}),
      rowCount: block.table.rowCount,
      columnCount: block.table.columnCount,
      header: block.table.header,
      placeholders: block.placeholders ?? [],
      mode: "fill_cells_or_replace_table",
      rows: block.table.rows
    });
  }

  return tables;
}

function buildFigures(blocks, _sections, relationships) {
  const figures = [];

  for (const captionBlock of blocks) {
    if (captionBlock.captionType !== "figure") continue;

    const imageBlock = findFigureImageAnchor(blocks, captionBlock.index);
    const drawingCount = imageBlock?.drawingCount ?? 0;
    const placeholderText = isFigurePlaceholderText(imageBlock?.text) ? imageBlock.text : undefined;
    const anchorKind = getFigureAnchorKind(imageBlock);
    const id = `fig_${figures.length + 1}${captionBlock.sectionNumber ? `_${captionBlock.sectionNumber.replaceAll(".", "_")}` : "_front"}`;

    figures.push({
      id,
      type: "figure",
      ...(captionBlock.section ? { section: captionBlock.section } : {}),
      ...(captionBlock.sectionNumber ? { sectionNumber: captionBlock.sectionNumber } : {}),
      ...(imageBlock ? { imageBlock: imageBlock.index, anchorKind } : { anchorKind: "caption_only" }),
      captionBlock: captionBlock.index,
      caption: captionBlock.text,
      ...(placeholderText ? { placeholderText } : {}),
      drawingCount,
      ...(imageBlock?.drawingRelationshipIds?.length ? { drawingRelationshipIds: imageBlock.drawingRelationshipIds } : {}),
      ...(imageBlock?.drawingRelationshipIds?.length
        ? {
            imageTargets: imageBlock.drawingRelationshipIds
              .map((id) => relationships.get(id)?.target)
              .filter(Boolean)
          }
        : {}),
      mode: anchorKind === "drawing" || anchorKind === "placeholder" ? "replace_image_keep_caption" : "insert_image_before_caption"
    });
  }

  return figures;
}

function addTemplateInvisibleAnchors(blocks, sections, tables, figures) {
  for (const section of sections) {
    section.directBodyRange = getSectionDirectBodyRange(section, sections, blocks);
    section.anchors = {
      body: makeSdtAnchor(`ps:section:${section.id}:body`, `${section.number} ${section.title} 正文`)
    };
  }

  for (const table of tables) {
    table.anchors = {
      table: makeSdtAnchor(`ps:table:${table.id}`, table.caption || table.id),
      ...(typeof table.captionBlock === "number"
        ? { caption: makeSdtAnchor(`ps:table:${table.id}:caption`, `${table.caption || table.id} 题注`) }
        : {})
    };
  }

  for (const figure of figures) {
    figure.anchors = {
      ...(typeof figure.imageBlock === "number"
        ? { image: makeSdtAnchor(`ps:figure:${figure.id}:image`, `${figure.caption || figure.id} 图片`) }
        : {}),
      caption: makeSdtAnchor(`ps:figure:${figure.id}:caption`, `${figure.caption || figure.id} 题注`)
    };
  }
}

function getSectionDirectBodyRange(section, sections, blocks) {
  const firstChild = (section.childSections ?? [])
    .map((id) => sections.find((candidate) => candidate.id === id))
    .filter(Boolean)
    .sort((a, b) => a.headingBlock - b.headingBlock)[0];
  const nextSameOrHigher = sections
    .filter((candidate) => candidate.headingBlock > section.headingBlock && candidate.headingLevel <= section.headingLevel)
    .sort((a, b) => a.headingBlock - b.headingBlock)[0];
  const bodyEnd = firstChild
    ? firstChild.headingBlock - 1
    : nextSameOrHigher
      ? nextSameOrHigher.headingBlock - 1
      : blocks.length - 1;
  return [section.headingBlock + 1, Math.max(section.headingBlock, bodyEnd)];
}

function insertTemplateInvisibleAnchors(documentXml, blocks, sections, tables, figures) {
  const bodyOpen = documentXml.match(/<w:body\b[^>]*>/);
  const bodyEnd = documentXml.lastIndexOf("</w:body>");
  if (!bodyOpen || bodyOpen.index === undefined || bodyEnd < 0) return documentXml;

  const bodyStart = bodyOpen.index + bodyOpen[0].length;
  const bodyXml = documentXml.slice(bodyStart, bodyEnd);
  const wrappers = [];

  const addWrapper = (anchor, start, end, scope) => {
    if (!anchor?.tag || !Number.isFinite(start) || !Number.isFinite(end)) return;
    if (start < 0 || end < start || end > bodyXml.length) return;
    wrappers.push({
      start,
      end,
      scope,
      tag: anchor.tag,
      alias: anchor.alias
    });
  };

  for (const section of sections) {
    const headingBlock = blocks[section.headingBlock];
    if (!headingBlock) continue;

    const [, directBodyEnd] = section.directBodyRange ?? getSectionDirectBodyRange(section, sections, blocks);
    const start = headingBlock._end;
    const end = blocks[directBodyEnd + 1]?._start ?? bodyXml.length;
    addWrapper(section.anchors?.body, start, end, "section");
  }

  for (const table of tables) {
    const tableBlock = blocks[table.block];
    if (tableBlock) {
      addWrapper(table.anchors?.table, tableBlock._start, tableBlock._end, "block");
    }

    const captionBlock = typeof table.captionBlock === "number" ? blocks[table.captionBlock] : undefined;
    if (captionBlock) {
      addWrapper(table.anchors?.caption, captionBlock._start, captionBlock._end, "block");
    }
  }

  for (const figure of figures) {
    const imageBlock = typeof figure.imageBlock === "number" ? blocks[figure.imageBlock] : undefined;
    if (imageBlock) {
      addWrapper(figure.anchors?.image, imageBlock._start, imageBlock._end, "block");
    }

    const captionBlock = blocks[figure.captionBlock];
    if (captionBlock) {
      addWrapper(figure.anchors?.caption, captionBlock._start, captionBlock._end, "block");
    }
  }

  const nextBodyXml = applySdtWrappers(bodyXml, wrappers);
  return `${documentXml.slice(0, bodyStart)}${nextBodyXml}${documentXml.slice(bodyEnd)}`;
}

function applySdtWrappers(bodyXml, wrappers) {
  if (!wrappers.length) return bodyXml;

  const insertions = [];
  for (const wrapper of wrappers) {
    const empty = wrapper.start === wrapper.end;
    const span = wrapper.end - wrapper.start;
    insertions.push({
      offset: wrapper.start,
      edge: "start",
      empty,
      span,
      xml: buildSdtStart(wrapper.tag, wrapper.alias) + (empty ? "<w:p/>" : "")
    });
    insertions.push({
      offset: wrapper.end,
      edge: "end",
      empty,
      span,
      xml: buildSdtEnd()
    });
  }

  const orderedInsertions = insertions.sort(compareSdtWrapperInsertions);
  let cursor = 0;
  const pieces = [];

  for (const insertion of orderedInsertions) {
    if (insertion.offset < cursor) continue;
    pieces.push(bodyXml.slice(cursor, insertion.offset), insertion.xml);
    cursor = insertion.offset;
  }
  pieces.push(bodyXml.slice(cursor));
  return pieces.join("");
}

function compareSdtWrapperInsertions(left, right) {
  if (left.offset !== right.offset) return left.offset - right.offset;
  const leftGroup = getSdtWrapperInsertionGroup(left);
  const rightGroup = getSdtWrapperInsertionGroup(right);
  if (leftGroup !== rightGroup) return leftGroup - rightGroup;
  if (left.edge === "start" && right.edge === "start") return right.span - left.span;
  if (left.edge === "end" && right.edge === "end") return left.span - right.span;
  return 0;
}

function getSdtWrapperInsertionGroup(insertion) {
  if (insertion.empty) return insertion.edge === "start" ? 2 : 3;
  return insertion.edge === "end" ? 1 : 4;
}

function stripManagedTemplateBookmarks(documentXml) {
  const managedIds = new Set();
  const withoutStarts = documentXml.replace(/<w:bookmarkStart\b[^>]*\/>/g, (tag) => {
    const name = getAttr(tag, "w:name");
    const id = getAttr(tag, "w:id");
    if (!id || !isManagedBookmarkName(name)) return tag;
    managedIds.add(id);
    return "";
  });

  if (!managedIds.size) return withoutStarts;
  return withoutStarts.replace(/<w:bookmarkEnd\b[^>]*\/>/g, (tag) => {
    const id = getAttr(tag, "w:id");
    return id && managedIds.has(id) ? "" : tag;
  });
}

function stripManagedTemplateSdts(documentXml) {
  let current = documentXml;
  for (;;) {
    const next = stripManagedTemplateSdtsOnce(current);
    if (next === current) return current;
    current = next;
  }
}

function stripManagedTemplateSdtsOnce(documentXml) {
  let cursor = 0;
  const pieces = [];

  while (cursor < documentXml.length) {
    const start = findNextElementStart(documentXml, "w:sdt", cursor);
    if (start < 0) break;

    const element = readBalancedElement(documentXml, start, "w:sdt");
    if (!element) break;

    pieces.push(documentXml.slice(cursor, start));
    pieces.push(isManagedSdt(element.xml) ? extractSdtContent(element.xml) : element.xml);
    cursor = element.end;
  }

  if (cursor === 0) return cleanTemplateDrawingParagraphs(documentXml);
  pieces.push(documentXml.slice(cursor));
  return pieces.join("");
}

function stripManagedTemplateVisibleMarkers(documentXml) {
  return documentXml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraphXml) => {
    const text = extractText(paragraphXml).replace(/\s+/g, "");
    return isManagedSectionVisibleMarker(text) ? "" : paragraphXml;
  });
}

function cleanTemplateSourceText(documentXml) {
  let cursor = 0;
  const pieces = [];

  while (cursor < documentXml.length) {
    const start = findNextElementStart(documentXml, "w:sdt", cursor);
    if (start < 0) break;

    const element = readBalancedElement(documentXml, start, "w:sdt");
    if (!element) break;

    pieces.push(documentXml.slice(cursor, start));
    pieces.push(cleanTemplateSdt(element.xml));
    cursor = element.end;
  }

  if (cursor === 0) return documentXml;
  pieces.push(documentXml.slice(cursor));
  return cleanTemplateDrawingParagraphs(pieces.join(""));
}

function cleanTemplateSdt(sdtXml) {
  const tag = getSingleAttrElement(sdtXml, "w:tag", "w:val") ?? "";
  if (/^ps:section:sec_\d+(?:_\d+)*:body$/.test(tag)) {
    return replaceSdtContent(sdtXml, cleanSectionBodyContent(extractSdtContent(sdtXml)));
  }

  if (/^ps:table:[^:]+$/.test(tag)) {
    return replaceSdtContent(sdtXml, cleanTemplateTablesInXml(extractSdtContent(sdtXml)));
  }

  if (/^ps:figure:[^:]+:image$/.test(tag)) {
    return replaceSdtContent(
      sdtXml,
      buildTemplatePlaceholderParagraphFromXml(extractSdtContent(sdtXml), FIGURE_PLACEHOLDER_TEXT, buildDefaultTemplateBodyPlaceholderParagraph(FIGURE_PLACEHOLDER_TEXT))
    );
  }

  return sdtXml;
}

function cleanSectionBodyContent(contentXml) {
  const preserved = [];
  let hasBodyPlaceholder = false;
  let cursor = 0;

  while (cursor < contentXml.length) {
    const candidates = [
      { tagName: "w:sdt", start: findNextElementStart(contentXml, "w:sdt", cursor) },
      { tagName: "w:tbl", start: findNextElementStart(contentXml, "w:tbl", cursor) },
      { tagName: "w:p", start: findNextElementStart(contentXml, "w:p", cursor) }
    ].filter((candidate) => candidate.start >= 0);

    if (!candidates.length) break;
    const candidate = candidates.sort((left, right) => left.start - right.start)[0];
    const element = readBalancedElement(contentXml, candidate.start, candidate.tagName);
    if (!element) break;

    if (candidate.tagName === "w:sdt") {
      const tag = getSingleAttrElement(element.xml, "w:tag", "w:val") ?? "";
      if (/^ps:(?:table|figure):/.test(tag)) {
        preserved.push(cleanTemplateSdt(element.xml));
      }
    } else if (candidate.tagName === "w:tbl") {
      preserved.push(cleanTemplateTableXml(element.xml));
    } else if (isStructuralTemplateParagraph(element.xml)) {
      preserved.push(element.xml);
    } else if (!hasBodyPlaceholder) {
      preserved.push(buildTemplatePlaceholderParagraph(element.xml, SECTION_BODY_PLACEHOLDER_TEXT));
      hasBodyPlaceholder = true;
    }

    cursor = element.end;
  }

  if (!hasBodyPlaceholder) {
    preserved.unshift(buildDefaultTemplateBodyPlaceholderParagraph(SECTION_BODY_PLACEHOLDER_TEXT));
  }
  return preserved.join("");
}

function isStructuralTemplateParagraph(paragraphXml) {
  const text = extractText(paragraphXml).replace(/\s+/g, "");
  return Boolean(text && (/^[表图]\d/.test(text) || isFigurePlaceholderText(text)));
}

function cleanTemplateDrawingParagraphs(documentXml) {
  return documentXml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraphXml) =>
    hasImageDrawing(paragraphXml) ? buildTemplatePlaceholderParagraph(paragraphXml, FIGURE_PLACEHOLDER_TEXT) : paragraphXml
  );
}

function hasImageDrawing(xml) {
  return /<w:(?:drawing|pict)\b/.test(xml) && /<pic:pic\b|<a:blip\b|<v:imagedata\b/.test(xml);
}

function cleanTemplateTablesInXml(xml) {
  return xml.replace(/<w:tbl\b[\s\S]*?<\/w:tbl>/g, (tableXml) => cleanTemplateTableXml(tableXml));
}

function cleanTemplateTableXml(tableXml) {
  let rowIndex = -1;
  return tableXml.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, (rowXml) => {
    rowIndex += 1;
    let cellIndex = -1;
    return rowXml.replace(/<w:tc\b[\s\S]*?<\/w:tc>/g, (cellXml) => {
      cellIndex += 1;
      return shouldKeepTemplateTableCellText(rowIndex, cellIndex, cellXml) ? cellXml : clearTemplateTableCellText(cellXml);
    });
  });
}

function shouldKeepTemplateTableCellText(rowIndex, cellIndex, cellXml) {
  const text = extractText(cellXml).trim();
  if (!text) return true;
  if (isTemplateFieldOnlyCellText(text)) return true;
  if (rowIndex === 0) return true;
  return cellIndex === 0 && isLikelyTemplateLabelCellText(text);
}

function isTemplateFieldOnlyCellText(text) {
  const compact = text.replace(/\s+/g, "");
  return /^(?:\{[^{}]+\}|\$\{[^{}]+\}|\[\[PS:field:[^\]\n]+\]\])+$/.test(compact);
}

function isLikelyTemplateLabelCellText(text) {
  const compact = text.replace(/\s+/g, "");
  if (!compact || /^\d+$/.test(compact)) return false;
  if (compact.length > 18) return false;
  if (/[，。；;：:]/.test(compact)) return false;
  if (/统一身份认证系统|示例|待补充|主机房|灾备|安全认证网关|运维安全网关/.test(compact)) return false;
  if (/密码应用措施|风险替代措施|通道\d|GM\/T|GB\/T|SM[2349]|VPN|TLCP|HMAC/.test(compact)) return false;
  return true;
}

function clearTemplateTableCellText(cellXml) {
  let paragraphIndex = -1;
  const nextCellXml = cellXml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraphXml) => {
    paragraphIndex += 1;
    return paragraphIndex === 0 ? buildTemplatePlaceholderParagraph(paragraphXml, TABLE_CELL_PLACEHOLDER_TEXT) : "";
  });
  if (paragraphIndex >= 0) return nextCellXml;

  const openTag = cellXml.match(/^<w:tc\b[^>]*>/)?.[0] ?? "<w:tc>";
  const cellProperties = cellXml.match(/<w:tcPr\b[\s\S]*?<\/w:tcPr>/)?.[0] ?? "";
  return `${openTag}${cellProperties}${buildDefaultTableCellPlaceholderParagraph(TABLE_CELL_PLACEHOLDER_TEXT)}</w:tc>`;
}

function buildTemplatePlaceholderParagraphFromXml(xml, placeholderText, fallbackXml) {
  const paragraph = xml.match(/<w:p\b[\s\S]*?<\/w:p>/)?.[0];
  return paragraph ? buildTemplatePlaceholderParagraph(paragraph, placeholderText) : fallbackXml;
}

function buildTemplatePlaceholderParagraph(paragraphXml, placeholderText) {
  let openTag = paragraphXml.match(/^<w:p\b[^>]*>/)?.[0] ?? "<w:p>";
  if (openTag.endsWith("/>")) openTag = `${openTag.slice(0, -2)}>`;
  const paragraphProperties = paragraphXml.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/)?.[0] ?? "";
  const runProperties = findPreferredTemplateRunProperties(paragraphXml);
  return `${openTag}${paragraphProperties}<w:r>${runProperties}<w:t>${escapeXml(placeholderText)}</w:t></w:r></w:p>`;
}

function findPreferredTemplateRunProperties(paragraphXml) {
  const runs = Array.from(paragraphXml.matchAll(/<w:r\b[\s\S]*?<\/w:r>/g)).map((match) => match[0]);
  const textRuns = runs.filter((runXml) => extractText(runXml).trim());
  const cleanRun = textRuns.find((runXml) => !/<w:b(?:Cs)?\b/.test(runXml) && !/<w:i(?:Cs)?\b/.test(runXml));
  const selectedRun = cleanRun ?? textRuns[0] ?? runs[0] ?? "";
  return selectedRun.match(/<w:rPr\b[\s\S]*?<\/w:rPr>/)?.[0] ?? "";
}

function buildDefaultTemplateBodyPlaceholderParagraph(placeholderText) {
  return `<w:p><w:pPr><w:pStyle w:val="30"/><w:wordWrap w:val="0"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>${escapeXml(placeholderText)}</w:t></w:r></w:p>`;
}

function buildDefaultTableCellPlaceholderParagraph(placeholderText) {
  return `<w:p><w:r><w:t>${escapeXml(placeholderText)}</w:t></w:r></w:p>`;
}

function isManagedSdt(xml) {
  const tag = getSingleAttrElement(xml, "w:tag", "w:val");
  return /^ps:/.test(tag ?? "");
}

function extractSdtContent(xml) {
  const match = xml.match(/<w:sdtContent\b[^>]*>([\s\S]*)<\/w:sdtContent>/);
  return match?.[1] ?? xml;
}

function replaceSdtContent(sdtXml, contentXml) {
  const match = sdtXml.match(/<w:sdtContent\b[^>]*>/);
  const end = sdtXml.lastIndexOf("</w:sdtContent>");
  if (!match || match.index === undefined || end < 0) return sdtXml;
  return `${sdtXml.slice(0, match.index + match[0].length)}${contentXml}${sdtXml.slice(end)}`;
}

function isManagedSectionVisibleMarker(text) {
  return /^\[\[\/?PS:(?!field:)[A-Za-z0-9_]+\]\]$/.test(text);
}

function isManagedBookmarkName(name) {
  return /^(ps_sec_|ps_table_|ps_fig_)/.test(name ?? "");
}

function makeSdtAnchor(tag, alias = tag) {
  return {
    tag,
    alias
  };
}

function normalizeTemplateFieldMarkers(documentXml) {
  return documentXml.replace(/<w:p\b[\s\S]*?<\/w:p>|<w:tbl\b[\s\S]*?<\/w:tbl>/g, (blockXml) =>
    normalizeTextNodeFieldMarkers(blockXml)
  );
}

function normalizeTextNodeFieldMarkers(xml) {
  const tokens = collectTextTokens(xml);
  if (!tokens.length) return xml;

  const fullText = tokens.map((token) => token.text).join("");
  const replacements = selectFieldMarkerReplacements(xml, fullText);
  if (!replacements.length) return xml;

  for (const replacement of replacements.sort((left, right) => right.start - left.start)) {
    applyTextReplacement(tokens, replacement);
  }
  return renderTextTokens(xml, tokens);
}

function selectFieldMarkerReplacements(xml, fullText) {
  const replacements = [];
  for (const match of fullText.matchAll(/\[\[PS:field:([^\]\n]{1,120})\]\]|\$\{([^{}\n]{1,120})\}|(?<!\$)\{([^{}\n]{1,120})\}/g)) {
    const key = normalizeFieldMarkerKey(match[1] ?? match[2] ?? match[3] ?? "");
    if (!key || isManagedSectionVisibleMarker(key)) continue;
    if (isDocxtemplaterFieldText(match[0])) continue;
    replacements.push({ start: match.index, end: match.index + match[0].length, value: makeFieldMarker(key) });
  }

  if (/<w:highlight\b/.test(xml)) {
    for (const run of matchElements(xml, "w:r")) {
      if (!hasVisibleHighlight(run.xml)) continue;
      const runText = extractText(run.xml);
      const key = normalizeHighlightedFieldKey(runText);
      if (!key) continue;

      let index = fullText.indexOf(runText);
      while (index >= 0) {
        if (!overlapsReplacement(replacements, index, index + runText.length)) {
          replacements.push({ start: index, end: index + runText.length, value: makeFieldMarker(key) });
          break;
        }
        index = fullText.indexOf(runText, index + runText.length);
      }
    }
  }

  return selectNonOverlappingTextReplacements(replacements);
}

function normalizeFieldMarkerKey(value) {
  return value.replace(/\s+/g, "").trim();
}

function normalizeHighlightedFieldKey(value) {
  const key = normalizeFieldMarkerKey(value);
  if (!key || isFieldMarkerText(key)) return "";
  if (key.startsWith("[[PS:")) return "";
  if (key.length > 40) return "";
  if (key === "待补充") return "";
  if (!isLikelyHighlightedPlaceholder(key)) return "";
  return key;
}

function hasVisibleHighlight(xml) {
  return /<w:highlight\b(?=[^>]*\bw:val="(?!none")[^"]+")/.test(xml);
}

function isLikelyHighlightedPlaceholder(key) {
  if (/^(XXX|XXX\.\.\.XXX)$/i.test(key)) return true;
  if (/^\d{1,3}$/.test(key)) return true;
  if (/^[一二三四五六七八九十百]+条$/.test(key)) return true;
  return false;
}

function isFieldMarkerText(value) {
  return /^(?:\[\[PS:field:[^\]]+\]\]|\{[^{}\n]+\})$/.test(value);
}

function isDocxtemplaterFieldText(value) {
  return /^(?<!\$)\{[^{}\n]+\}$/.test(value);
}

function makeFieldMarker(key) {
  return `{${key}}`;
}

function overlapsReplacement(replacements, start, end) {
  return replacements.some((replacement) => start < replacement.end && end > replacement.start);
}

function selectNonOverlappingTextReplacements(replacements) {
  const selected = [];
  let lastEnd = -1;
  for (const replacement of [...replacements].sort((left, right) => left.start - right.start || right.end - left.end)) {
    if (replacement.start < lastEnd) continue;
    selected.push(replacement);
    lastEnd = replacement.end;
  }
  return selected;
}

function collectTextTokens(xml) {
  const tokens = [];
  const pattern = /(<w:t\b[^>]*>)([\s\S]*?)(<\/w:t>)/g;
  let textOffset = 0;
  let match;

  while ((match = pattern.exec(xml))) {
    const text = decodeXml(match[2]);
    tokens.push({
      start: match.index,
      end: match.index + match[0].length,
      openTag: match[1],
      closeTag: match[3],
      text,
      textStart: textOffset,
      textEnd: textOffset + text.length
    });
    textOffset += text.length;
  }

  return tokens;
}

function applyTextReplacement(tokens, replacement) {
  const startToken = findTextTokenAtOffset(tokens, replacement.start);
  const endToken = findTextTokenAtOffset(tokens, replacement.end - 1);
  if (!startToken || !endToken) return;

  const startOffset = replacement.start - startToken.textStart;
  const endOffset = replacement.end - endToken.textStart;
  const startIndex = tokens.indexOf(startToken);
  const endIndex = tokens.indexOf(endToken);

  if (startIndex === endIndex) {
    startToken.text = `${startToken.text.slice(0, startOffset)}${replacement.value}${startToken.text.slice(endOffset)}`;
    return;
  }

  startToken.text = `${startToken.text.slice(0, startOffset)}${replacement.value}`;
  for (let index = startIndex + 1; index < endIndex; index += 1) {
    tokens[index].text = "";
  }
  endToken.text = endToken.text.slice(endOffset);
}

function findTextTokenAtOffset(tokens, offset) {
  return tokens.find((token) => offset >= token.textStart && offset < token.textEnd);
}

function renderTextTokens(xml, tokens) {
  let nextXml = xml;
  for (const token of [...tokens].reverse()) {
    const openTag = ensureTextNodePreservesSpaces(token.openTag, token.text);
    nextXml = `${nextXml.slice(0, token.start)}${openTag}${escapeXml(token.text)}${token.closeTag}${nextXml.slice(token.end)}`;
  }
  return nextXml;
}

function ensureTextNodePreservesSpaces(openTag, text) {
  if (!/^\s|\s$/.test(text) || /\sxml:space=/.test(openTag)) return openTag;
  return openTag.replace(/>$/, ' xml:space="preserve">');
}

function buildFieldAnchors(blocks) {
  const counts = new Map();
  const anchors = [];

  for (const block of blocks) {
    const markers = extractFieldMarkers(block.text ?? "");
    for (const marker of markers) {
      const occurrence = (counts.get(marker.key) ?? 0) + 1;
      counts.set(marker.key, occurrence);
      anchors.push({
        id: `field_${sanitizeAnchorId(marker.key)}_${occurrence}`,
        key: marker.key,
        marker: marker.marker,
        occurrence,
        block: block.index,
        ...(block.section ? { section: block.section } : {}),
        ...(block.sectionNumber ? { sectionNumber: block.sectionNumber } : {}),
        type: block.type
      });
    }
  }

  return anchors;
}

function extractFieldMarkers(text) {
  return Array.from(text.matchAll(/(?<!\$)\{([^{}\n]+)\}/g), (match) => ({
    key: match[1].trim(),
    marker: match[0]
  })).filter((marker) => marker.key);
}

function sanitizeAnchorId(value) {
  return value
    .replace(/\s+/g, "")
    .replace(/[^A-Za-z0-9_\u4e00-\u9fa5]/g, "_")
    .slice(0, 40);
}

function buildSdtStart(tag, alias) {
  return [
    "<w:sdt>",
    "<w:sdtPr>",
    `<w:alias w:val="${escapeXml(alias || tag)}"/>`,
    `<w:tag w:val="${escapeXml(tag)}"/>`,
    "</w:sdtPr>",
    "<w:sdtContent>"
  ].join("");
}

function buildSdtEnd() {
  return "</w:sdtContent></w:sdt>";
}

function escapeXml(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function applyTemplateGuidance(sections, tables, figures) {
  for (const section of sections) {
    section.writingHint = getSectionWritingHint(section);
    const relatedTables = tables
      .filter((table) => table.sectionNumber === section.number)
      .map((table) => table.id);
    if (relatedTables.length) section.relatedTables = relatedTables;

    const relatedFigures = figures
      .filter((figure) => figure.sectionNumber === section.number)
      .map((figure) => figure.id);
    if (relatedFigures.length) section.relatedFigures = relatedFigures;
  }

  for (const table of tables) {
    table.purpose = getTablePurpose(table);
    table.writeStrategy = table.placeholders?.length ? "template_fields_or_template_cells" : "replace_table_when_data_complete";
  }

  for (const figure of figures) {
    figure.purpose = getFigurePurpose(figure);
    figure.recommendedLabel = getFigureRecommendedLabel(figure);
  }
}

function buildFieldGuide() {
  return [
    { key: "应用系统", description: "信息系统或业务平台名称", example: "统一身份认证系统" },
    { key: "建设单位", description: "项目建设或运营单位", example: "示例政务服务中心" },
    { key: "单位省份", description: "所属省级行政区域", example: "广东省" },
    { key: "单位地址", description: "单位通信地址", example: "广州市天河区 XX 路 XX 号" },
    { key: "单位邮编", description: "邮政编码", example: "510000" },
    { key: "等保级别", description: "网络安全等级保护级别", example: "三级" },
    { key: "物理机房1", description: "主机房名称", example: "核心机房" },
    { key: "物理机房1地址", description: "主机房地址", example: "XX 数据中心 A 区" },
    { key: "物理机房1管理单位", description: "主机房管理责任主体", example: "示例政务服务中心" },
    { key: "物理机房2", description: "备用或灾备机房名称", example: "灾备机房" },
    { key: "物理机房2地址", description: "备用或灾备机房地址", example: "XX 数据中心 B 区" },
    { key: "物理机房2管理单位", description: "备用或灾备机房管理责任主体", example: "示例政务服务中心" },
    { key: "应用子系统1", description: "第一个主要业务子系统", example: "认证服务子系统" },
    { key: "应用子系统2", description: "第二个主要业务子系统", example: "权限管理子系统" },
    { key: "云平台", description: "依赖的云平台名称", example: "政务云专有区" },
    { key: "应用安全网关", description: "应用侧安全接入或认证网关", example: "安全认证网关" },
    { key: "运维网关1", description: "运维通道网关", example: "运维安全网关" },
    { key: "运维网关2", description: "第二套运维通道网关", example: "运维 VPN 网关" },
    { key: "安全接入网关1", description: "第一套安全接入网关", example: "SSL VPN 安全网关" },
    { key: "安全接入网关2", description: "第二套安全接入网关", example: "IPSec VPN 安全网关" },
    { key: "密码系统产品", description: "密码支撑平台产品集合", example: "服务器密码机、签名验签服务器" },
    { key: "密码安全设备", description: "密码设备集合", example: "服务器密码机、密码服务管理平台" },
    { key: "网络安全设备", description: "网络安全设备集合", example: "防火墙、堡垒机、日志审计" }
  ];
}

function getSectionWritingHint(section) {
  const specific = getSectionHintMap()[section.number];
  if (specific) return specific;

  const title = section.title;
  if (/风险分析/.test(title)) return "分析当前控制措施不足、可能攻击路径、业务和数据影响，结论要支撑后续密码应用需求。";
  if (/密码应用需求/.test(title) || /安全管理需求/.test(title)) return "从适用性、保护对象、密码技术需求、管理要求四个角度给出需求结论，不要直接写产品堆砌。";
  if (/身份鉴别/.test(title)) return "说明鉴别对象、鉴别凭据、密码算法或证书机制、调用路径和预期安全效果。";
  if (/完整性/.test(title)) return "说明保护对象、完整性计算方式、存储或传输位置、校验时机和异常处理。";
  if (/机密性/.test(title)) return "说明保护对象、加密算法、密钥来源、加密位置、解密调用和访问控制。";
  if (/密钥/.test(title)) return "覆盖密钥生成、存储、分发、使用、备份恢复、归档、销毁和职责分离。";
  if (/实施|计划|保障/.test(title)) return "围绕实施范围、阶段任务、输出物、责任主体、风险控制和验收要求展开。";
  if (/表|清单|概算/.test(title)) return "优先依据相关 tables 锚点生成结构化表格，资料不足的金额、型号和数量写待补充。";
  return `围绕“${section.title}”编写与密码应用方案相关的专业正文，资料不足处写待补充。`;
}

function getTablePurpose(table) {
  const caption = table.caption ?? "";
  const header = table.header?.join("、") ?? "";
  if (table.id === "table_3_2_1") return "系统基本信息、建设单位、地址、邮编、等保和测评情况。";
  if (table.id === "table_4_2_2_1") return "物理机房名称、地址和管理责任主体。";
  if (/服务器及存储/.test(caption)) return "服务器、存储设备、用途、数量和重要程度清单。";
  if (/网络及安全设备/.test(caption)) return "网络设备、安全设备和用途清单。";
  if (/数据库/.test(caption)) return "数据库管理系统、版本、部署位置和数量。";
  if (/业务应用情况/.test(caption)) return "业务应用或子系统名称、版本、部署位置和主要功能。";
  if (/关键数据类型/.test(caption)) return "关键数据类型、所属应用、存储位置和安全需求。";
  if (/应用用户/.test(caption)) return "业务用户、管理用户、运维用户等应用用户清单。";
  if (/需求分析/.test(caption) || table.id === "table_14_3_9") return "按 GB/T 39786 适用项汇总密码应用需求。";
  if (/密钥和密码算法/.test(caption)) return "各层面密钥、密码算法、相关设备和用途。";
  if (/密码设备遵循标准/.test(caption)) return "密码产品与遵循标准、认证要求的对应关系。";
  if (/密码设备功能/.test(caption)) return "密码设备能力说明。";
  if (/密码设备部署情况/.test(caption)) return "密码设备部署位置、部署方式和管理使用方式。";
  if (/保护对象/.test(caption)) return "密码应用保护对象和安全需求。";
  if (/所需密码产品/.test(caption)) return "实现该场景所需密码产品、标准、算法和用途。";
  if (/密钥管理/.test(caption)) return "该场景密钥生命周期管理。";
  if (/产品清单|选型原则/.test(caption)) return "密码产品清单、数量、用途和选型原则。";
  if (/合规性分析/.test(caption)) return "逐项分析适用性、保障措施、替代措施和自评结果。";
  if (/经费概算/.test(caption)) return "产品、服务、数量、单价、小计和备注，金额不明确时写待测算。";
  return header ? `围绕表头“${header}”填写结构化内容。` : "根据所在章节填写模板表格内容。";
}

function getFigurePurpose(figure) {
  const caption = figure.caption ?? "";
  if (/网络框架/.test(caption)) return "体现业务区、接入区、服务器区、安全运维区、密码资源区等逻辑区域。";
  if (/网络拓扑/.test(caption)) return "体现网络边界、设备连接、外部访问路径和安全设备部署。";
  if (/密码应用技术框架/.test(caption)) return "体现计算平台、密码支撑平台、业务应用和管理体系之间的关系。";
  if (/门禁身份鉴别/.test(caption)) return "体现门禁卡、门禁控制器、密码模块和身份鉴别过程。";
  if (/门禁记录完整性/.test(caption)) return "体现门禁记录生成、完整性计算、存储和验证流程。";
  if (/视频监控完整性/.test(caption)) return "体现视频记录完整性保护流程。";
  if (/TLCP|SSL/.test(caption)) return "体现客户端、安全网关、应用系统之间的国密安全通道。";
  if (/身份鉴别流程/.test(caption)) return "体现业务用户登录、证书校验、签名验签和认证结果返回流程。";
  if (/访问控制信息/.test(caption)) return "体现访问控制信息签名、存储、验签和篡改发现流程。";
  if (/重要数据存储保护/.test(caption)) return "体现重要数据写入前的加密或完整性保护流程。";
  if (/重要数据存储读取/.test(caption)) return "体现重要数据读取、解密和完整性验证流程。";
  if (/密码应用部署拓扑/.test(caption)) return "体现密码设备部署位置、连接关系和服务调用路径。";
  return "按题注生成对应密码应用方案图示。";
}

function getFigureRecommendedLabel(figure) {
  return (figure.caption ?? "")
    .replace(/^图\s*\d+\s*/, "")
    .replace(/^图\s*\d+-\d+\s*/, "")
    .replace(/\s+/g, "")
    .trim();
}

function getSectionHintMap() {
  return {
  "1": "本章说明项目背景和合规依据，为后续密码应用建设必要性提供依据。",
  "1.1": "说明系统建设背景、业务定位、等保或密评要求、密码应用建设必要性和目标。",
  "1.2": "概述法律法规、政策文件和标准依据，并说明其与本系统密码应用建设的关系。",
  "2": "本章描述系统边界、计算平台、业务应用、密码应用现状和管理现状。",
  "2.1": "填写系统基本信息、建设单位、等保级别、是否依赖云平台、测评和密评情况；基础字段优先通过 template_fields 写入。",
  "2.2": "概述物理环境、网络环境和计算环境，为风险分析和设计章节提供现状基础。",
  "2.2.1": "描述机房名称、地址、管理责任主体、门禁、视频监控和值守情况。",
  "2.2.2": "描述网络整体结构、边界划分、设备组成、数据交互和现有安全防护措施。",
  "2.2.2.1": "描述逻辑网络框架；需要图片时使用 diagrams label：网络框架图。",
  "2.2.2.2": "描述网络拓扑、边界设备和访问路径；需要图片时使用 diagrams label：网络拓扑图。",
  "2.2.3": "描述服务器、存储、数据库、网络设备、安全设备和云资源现状。",
  "2.3": "描述业务场景、子系统、用户角色、关键数据、外部接口和责任主体。",
  "2.4": "说明当前已部署或拟部署的密码产品、密码协议、证书、密钥和密码服务。",
  "2.5": "说明现有密码管理制度、人员、运维、审计、培训和应急管理现状。",
  "3": "本章从技术和管理层面分析风险、适用性和密码应用需求，结论要支撑第 5 章。",
  "3.9": "按 GB/T 39786 适用项汇总需求分析；建议使用 table_14_3_9。",
  "4": "本章说明安全目标、设计原则、建设依据和密码技术标准。",
  "4.1": "描述合规目标、技术目标、管理目标、密评支撑目标和业务连续性目标。",
  "4.2": "说明合规性、体系化、适用性、最小改造、成熟可靠、可运维、可扩展等原则。",
  "5": "本章是核心设计章节，需要说明保护对象、密码产品、算法协议、部署位置、调用方式、密钥管理和安全效果。",
  "5.1": "描述密码应用技术框架；需要图片时使用 diagrams label：密码应用技术架构图。",
  "5.2": "描述物理和环境、网络和通信、设备和计算层面的密码应用方案。",
  "5.3": "描述密码服务机构、算法、密码设备标准、功能、部署、接入方式、密钥管理和自身安全性。",
  "5.3.7": "覆盖密钥生成、存储、分发、导入导出、使用、备份恢复、归档和销毁。",
  "5.4": "描述业务应用层身份鉴别、访问控制信息完整性、重要数据传输和存储保护、不可否认性等。",
  "5.4.9": "按保护对象、所需密码产品、所需密码服务、工作流程、密钥管理的顺序编写。",
  "5.5": "描述密码设备部署示意和部署设计；需要图片时使用 diagrams label：密码应用部署拓扑。",
  "6": "本章描述密码应用安全管理制度、人员管理、建设运行和应急处置。",
  "6.1": "说明制度体系、密钥管理制度、密码设备管理制度、操作规程和记录留存。",
  "6.2": "说明密钥管理员、密码操作员、密码审计员、系统管理员的职责和权限分离。",
  "6.3": "说明方案评审、采购、集成、上线、变更、运维、审计、定期评估和整改要求。",
  "6.4": "描述密码相关安全事件、应急组织、处置流程、公告流程、预案激活和修订机制。",
  "7": "根据第 5 章设计逐项生成安全与合规性分析，建议使用 table_36_7。",
  "8": "描述实施内容、实施计划、保障措施和经费概算。",
  "8.1": "说明设备采购部署、系统集成改造、联调测试、试运行和风险应对。",
  "8.2": "说明项目启动、调研、设计、部署、改造、联调、试运行和验收计划。",
  "8.3": "说明组织、人员、经费、质量、进度、安全和运维保障措施。",
  "8.4": "生成经费概算表；金额不明确时写待测算，建议使用 table_37_8_4。"
  };
}

function findFigureImageAnchor(blocks, captionBlockIndex) {
  const previous = blocks[captionBlockIndex - 1];
  if (
    previous?.type === "p" &&
    previous.captionType !== "figure" &&
    previous.captionType !== "table" &&
    previous.sectionNumber === blocks[captionBlockIndex]?.sectionNumber
  ) {
    return previous;
  }

  for (let index = captionBlockIndex - 2; index >= Math.max(0, captionBlockIndex - 8); index -= 1) {
    const candidate = blocks[index];
    if (!candidate || candidate.sectionNumber !== blocks[captionBlockIndex]?.sectionNumber) break;
    if (candidate.captionType || candidate.headingLevel || candidate.type === "tbl") break;
    if (candidate.drawingCount || isFigurePlaceholderText(candidate.text)) return candidate;
  }

  return undefined;
}

function getFigureAnchorKind(block) {
  if (!block) return "caption_only";
  if (block.drawingCount) return "drawing";
  if (isFigurePlaceholderText(block.text)) return "placeholder";
  if (!block.text) return "blank_paragraph";
  return "nearby_paragraph";
}

function findAdjacentCaptionBlock(blocks, blockIndex, type) {
  const previous = blocks[blockIndex - 1];
  if (previous?.captionType === type) return previous;
  const next = blocks[blockIndex + 1];
  if (next?.captionType === type) return next;
  return undefined;
}

function stripInternalBlockFields(block) {
  const { _xml, _start, _end, table, ...rest } = block;
  if (block.type === "tbl" && table) {
    return {
      ...rest,
      rowCount: table.rowCount,
      columnCount: table.columnCount,
      header: table.header
    };
  }
  return rest;
}

function inferTableHeader(rows) {
  const headerRow = rows.find((row) => row.cells.some((cell) => cell.text.trim()));
  if (!headerRow) return [];
  return headerRow.cells.map((cell) => cell.text.trim()).filter(Boolean);
}

function matchElements(xml, tagName) {
  const elements = [];
  let cursor = 0;

  while (cursor < xml.length) {
    const start = findNextElementStart(xml, tagName, cursor);
    if (start < 0) break;
    const element = readBalancedElement(xml, start, tagName);
    if (!element) break;
    elements.push(element);
    cursor = element.end;
  }

  return elements;
}

function readBalancedElement(xml, start, tagName) {
  const openTagEnd = xml.indexOf(">", start);
  if (openTagEnd < 0) return undefined;

  const openTag = xml.slice(start, openTagEnd + 1);
  if (openTag.endsWith("/>")) {
    return { xml: openTag, openTag, start, end: openTagEnd + 1 };
  }

  const openPattern = tagName;
  const closePattern = `</${tagName}>`;
  let cursor = openTagEnd + 1;
  let depth = 1;

  while (cursor < xml.length) {
    const nextOpen = findNextElementStart(xml, openPattern, cursor);
    const nextClose = xml.indexOf(closePattern, cursor);
    if (nextClose < 0) return undefined;

    if (nextOpen >= 0 && nextOpen < nextClose) {
      const nestedOpenEnd = xml.indexOf(">", nextOpen);
      if (nestedOpenEnd < 0) return undefined;
      const nestedOpen = xml.slice(nextOpen, nestedOpenEnd + 1);
      if (!nestedOpen.endsWith("/>")) depth += 1;
      cursor = nestedOpenEnd + 1;
      continue;
    }

    depth -= 1;
    cursor = nextClose + closePattern.length;
    if (depth === 0) {
      return {
        xml: xml.slice(start, cursor),
        openTag,
        start,
        end: cursor
      };
    }
  }

  return undefined;
}

function findNextElementStart(xml, tagName, cursor) {
  const pattern = new RegExp(`<${escapeRegExp(tagName)}(?=[\\s>/])`, "g");
  pattern.lastIndex = cursor;
  const match = pattern.exec(xml);
  return match ? match.index : -1;
}

function extractText(xml) {
  const pieces = [];
  const tokenPattern = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\s*\/>|<w:br\b[^>]*\/>/g;
  for (const match of xml.matchAll(tokenPattern)) {
    if (match[1] !== undefined) {
      pieces.push(decodeXml(match[1]));
    } else if (match[0].startsWith("<w:tab")) {
      pieces.push("\t");
    } else {
      pieces.push("\n");
    }
  }
  return pieces.join("").replace(/\s+\n/g, "\n").replace(/\n\s+/g, "\n").trim();
}

function extractPlaceholders(text) {
  const placeholders = [];
  for (const match of text.matchAll(/\[\[PS:field:([^\]\n]+)\]\]/g)) {
    const placeholder = match[1]?.trim();
    if (placeholder) placeholders.push(placeholder);
  }
  for (const match of text.matchAll(/\$\{([^{}]+)\}/g)) {
    const placeholder = match[1]?.trim();
    if (placeholder) placeholders.push(placeholder);
  }
  for (const match of text.matchAll(/\{([^{}]+)\}/g)) {
    if (match.index && text[match.index - 1] === "$") continue;
    const placeholder = match[1]?.trim();
    if (placeholder) placeholders.push(placeholder);
  }
  return uniqueSorted(placeholders);
}

function extractDrawingRelationshipIds(xml) {
  const ids = [];
  for (const match of xml.matchAll(/<a:blip\b[^>]*\br:embed="([^"]+)"/g)) {
    ids.push(match[1]);
  }
  return uniqueSorted(ids);
}

function getDrawingCount(xml) {
  return (xml.match(/<w:drawing\b/g) ?? []).length + (xml.match(/<w:pict\b/g) ?? []).length;
}

function getCaptionType(text, style) {
  const compact = text.replace(/\s+/g, "");
  if (!compact) return undefined;
  if (!style?.name?.toLowerCase().includes("caption") && !/^[表图]\d/.test(compact)) return undefined;
  if (/^表\d/.test(compact)) return "table";
  if (/^图\d/.test(compact)) return "figure";
  return undefined;
}

function isFigurePlaceholderText(text) {
  if (!text) return false;
  const normalized = text.replace(/\s+/g, "");
  return normalized === "【图片占位】" || normalized === "待补充" || normalized === "${图片}" || normalized === "${图示}";
}

function pruneUnusedDocumentImageMedia(zipFile, documentXml) {
  const relPath = "word/_rels/document.xml.rels";
  const relFile = zipFile.file(relPath);
  const relsXml = relFile?.asText();
  if (!relsXml) return;

  const usedRelationshipIds = new Set(
    Array.from(documentXml.matchAll(/r:(?:embed|link)="([^"]+)"/g)).map((match) => match[1])
  );
  const removedTargets = [];
  const nextRelsXml = relsXml.replace(/<Relationship\b[^>]*\/>/g, (tag) => {
    const id = getAttr(tag, "Id");
    const target = getAttr(tag, "Target");
    const type = getAttr(tag, "Type") ?? "";
    if (id && target && type.includes("/image") && !usedRelationshipIds.has(id)) {
      removedTargets.push(target);
      return "";
    }
    return tag;
  });

  if (nextRelsXml !== relsXml) zipFile.file(relPath, nextRelsXml);

  for (const target of removedTargets) {
    const mediaPath = resolveDocumentRelationshipTarget(target);
    if (mediaPath.startsWith("word/media/")) zipFile.remove(mediaPath);
  }
}

function resolveDocumentRelationshipTarget(target) {
  if (target.startsWith("/")) return target.replace(/^\/+/, "");
  return `word/${target.replace(/^\.\//, "")}`.replace(/\/{2,}/g, "/");
}

function getSingleAttrElement(xml, tagName, attrName) {
  const pattern = new RegExp(`<${escapeRegExp(tagName)}\\b([^>]*)\\/?>(?:</${escapeRegExp(tagName)}>)?`);
  const match = xml.match(pattern);
  return match ? getAttr(match[1] ?? "", attrName) : undefined;
}

function getAttr(attrs, attrName) {
  const escaped = escapeRegExp(attrName);
  const match = attrs.match(new RegExp(`(?:^|\\s)${escaped}="([^"]*)"`));
  return match ? decodeXml(match[1]) : undefined;
}

function getHeadingLevel(styleName) {
  const match = styleName.match(/^heading\s*(\d+)$/i);
  if (!match) return undefined;
  const level = Number.parseInt(match[1], 10);
  return Number.isFinite(level) ? level : undefined;
}

function stripHeadingNumberPrefix(text) {
  return text.replace(/^\s*\d+(?:\.\d+)*\s+/, "").trim();
}

function isUsefulStyleName(name) {
  return ["caption", "正文", "方案", "标题", "表格"].some((keyword) => name.includes(keyword));
}

function decodeXml(value) {
  return value
    .replace(/&#x([0-9a-f]+);/gi, (_, code) => String.fromCodePoint(Number.parseInt(code, 16)))
    .replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number.parseInt(code, 10)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function uniqueSorted(values) {
  return Array.from(new Set(values.filter(Boolean))).sort((a, b) => a.localeCompare(b, "zh-Hans-CN"));
}

function toProjectPath(path) {
  return relative(ROOT, path).replaceAll("\\", "/");
}
