import { readFile } from "node:fs/promises";
import { join } from "node:path";
import PizZip from "pizzip";
import { describe, expect, it } from "vitest";

interface TemplateCell {
  rowIndex: number;
  columnIndex: number;
  text: string;
  placeholders?: string[];
}

interface TemplateTable {
  id: string;
  caption?: string;
  rowCount: number;
  columnCount: number;
  placeholders: string[];
  purpose: string;
  writeStrategy: string;
  anchors?: {
    table: {
      tag: string;
      alias?: string;
    };
    caption?: {
      tag: string;
      alias?: string;
    };
  };
  rows: Array<{
    index: number;
    cells: TemplateCell[];
  }>;
}

interface TemplateFigure {
  id: string;
  sectionNumber?: string;
  imageBlock?: number;
  anchorKind: string;
  captionBlock: number;
  caption: string;
  mode: string;
  placeholderText?: string;
  purpose: string;
  recommendedLabel: string;
  anchors?: {
    image?: {
      tag: string;
      alias?: string;
    };
    caption: {
      tag: string;
      alias?: string;
    };
  };
}

interface TemplateJson {
  source: {
    docx: string;
  };
  statistics: {
    blockCount: number;
    sectionCount: number;
    tableCount: number;
    figureCount: number;
    drawingCount: number;
    placeholderCount: number;
  };
  fieldAnchors: Array<{
    id: string;
    key: string;
    marker: string;
    occurrence: number;
    block: number;
    section?: string;
    sectionNumber?: string;
    type: string;
  }>;
  sections: Array<{
    id: string;
    number: string;
    title: string;
    writingHint: string;
    headingBlock: number;
    bodyRange: [number, number];
    directBodyRange: [number, number];
    anchors: {
      body: {
        tag: string;
        alias?: string;
      };
    };
    relatedTables?: string[];
    relatedFigures?: string[];
  }>;
  tables: TemplateTable[];
  figures: TemplateFigure[];
}

describe("scheme template json", () => {
  it("contains detailed invisible anchors parsed from the built-in Word template", async () => {
    const template = JSON.parse(await readFile(join(process.cwd(), "docs", "密码应用方案.template.json"), "utf-8")) as TemplateJson;

    expect(template.source.docx).toBe("docs/密码应用方案.docx");
    expect("contentTemplate" in template.source).toBe(false);
    expect(template.statistics.blockCount).toBeGreaterThan(500);
    expect(template.statistics.sectionCount).toBeGreaterThanOrEqual(150);
    expect(template.statistics.tableCount).toBeGreaterThanOrEqual(35);
    expect(template.statistics.figureCount).toBeGreaterThanOrEqual(14);
    expect(template.statistics.drawingCount).toBe(0);
    expect(template.statistics.placeholderCount).toBeGreaterThanOrEqual(15);
    expect(template.fieldAnchors.length).toBeGreaterThanOrEqual(40);
    expect(template.fieldAnchors.find((anchor) => anchor.key === "应用系统")).toMatchObject({
      marker: "{应用系统}"
    });

    const networkSection = template.sections.find((section) => section.id === "sec_2_2_2");
    expect(networkSection).toMatchObject({
      number: "2.2.2",
      title: "网络环境",
      writingHint: expect.stringContaining("网络"),
      anchors: {
        body: {
          tag: "ps:section:sec_2_2_2:body"
        }
      }
    });
    expect(networkSection?.headingBlock).toBeGreaterThan(0);
    expect(networkSection?.bodyRange[0]).toBeGreaterThan(networkSection?.headingBlock ?? 0);
    expect(networkSection?.bodyRange[1]).toBeGreaterThanOrEqual(networkSection?.bodyRange[0] ?? 0);
    expect(networkSection?.directBodyRange[0]).toBeGreaterThan(networkSection?.headingBlock ?? 0);
    expect("visibleMarkers" in (networkSection ?? {})).toBe(false);

    const basicTable = template.tables.find((table) => table.caption === "表 21系统基本情况表");
    expect(basicTable).toBeTruthy();
    expect(basicTable).toMatchObject({
      rowCount: 16,
      columnCount: 7,
      purpose: expect.stringContaining("系统基本信息"),
      writeStrategy: "template_fields_or_template_cells"
    });
    expect(basicTable?.anchors).toMatchObject({
      table: {
        tag: "ps:table:table_3_2_1"
      },
      caption: {
        tag: "ps:table:table_3_2_1:caption"
      }
    });
    expect("visibleMarkers" in (basicTable ?? {})).toBe(false);
    expect(basicTable?.rows[1]?.cells[1]).toMatchObject({
      rowIndex: 1,
      columnIndex: 1,
      text: "{建设单位}",
      placeholders: ["建设单位"]
    });

    const physicalTable = template.tables.find((table) => table.caption === "表 22 物理环境情况");
    const machineRoomAddressCell = physicalTable?.rows.flatMap((row) => row.cells).find((cell) => cell.placeholders?.includes("物理机房1地址"));
    expect(machineRoomAddressCell).toMatchObject({
      rowIndex: 1,
      columnIndex: 1,
      text: "{物理机房1地址}"
    });

    const networkFigure = template.figures.find((figure) => figure.caption === "图 21 网络框架图");
    expect(networkFigure).toMatchObject({
      sectionNumber: "2.2.2.1",
      anchorKind: "placeholder",
      mode: "replace_image_keep_caption",
      placeholderText: "【图片占位】",
      anchors: {
        image: {
          tag: "ps:figure:fig_1_2_2_2_1:image"
        },
        caption: {
          tag: "ps:figure:fig_1_2_2_2_1:caption"
        }
      },
      recommendedLabel: "网络框架图",
      purpose: expect.stringContaining("逻辑区域")
    });
    expect(networkFigure?.imageBlock).toBeGreaterThan(0);
    expect(networkFigure?.captionBlock).toBeGreaterThan(networkFigure?.imageBlock ?? 0);
    expect("visibleMarkers" in (networkFigure ?? {})).toBe(false);
    expect(template.figures.find((figure) => figure.caption === "图 57 访问控制信息数字签名流程")).toMatchObject({
      anchorKind: "placeholder",
      mode: "replace_image_keep_caption",
      placeholderText: "【图片占位】"
    });
  });

  it("keeps the built-in docx as a clean template without legacy body prose", async () => {
    const docx = new PizZip(await readFile(join(process.cwd(), "docs", "密码应用方案.docx"), "binary"));
    const documentXml = docx.file("word/document.xml")?.asText() ?? "";
    const mediaFiles = Object.values(docx.files).filter((file) => file.name.startsWith("word/media/") && !file.dir);

    expect(documentXml).toContain('w:val="ps:section:sec_1_1:body"');
    expect(documentXml).toContain('w:val="ps:table:table_3_2_1"');
    expect(documentXml).toContain("【图片占位】");
    expect(documentXml).not.toContain("密码是保障网络与信息安全的核心技术和基础支撑");
    expect(documentXml).not.toContain("网络运营者开展经营和服务活动");
    expect(documentXml).not.toContain("培训与宣贯");
    expect(documentXml).not.toContain("标准规范类");
    expect(documentXml).not.toContain("项目技术文件类");
    expect(documentXml).not.toContain("密码应用措施");
    expect(documentXml).not.toContain("统一身份认证系统业务子系统");
    expect(documentXml).not.toContain("<w:drawing");
    expect(documentXml).not.toContain("<pic:pic");
    expect(documentXml).not.toContain("<a:blip");
    expect(documentXml).toContain("【正文占位】");
    expect(documentXml).toContain("【待填写】");
    expect(mediaFiles).toHaveLength(0);
  });
});
