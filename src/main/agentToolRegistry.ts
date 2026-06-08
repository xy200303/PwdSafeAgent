import { exec } from "node:child_process";
import { existsSync, readFileSync, readdirSync, statSync } from "node:fs";
import { promisify } from "node:util";
import { basename, extname, isAbsolute, join, resolve } from "node:path";
import OpenAI from "openai";
import type { ChatCompletionMessageParam, ChatCompletionMessageToolCall, ChatCompletionTool } from "openai/resources/chat/completions";
import type { AppSettings, SchemeProgressItem, SchemeSectionStatus } from "../shared/types";
import {
  compactText,
  formatBytes,
  getCurrentTimeText,
  getImageMimeType,
  isSupportedImageFile,
  readDocumentText,
  readImageDataUrl,
  sanitizeFileName,
  writeUtf8File,
  type BuiltinToolName,
  type ImageDataUrlResult
} from "./agentTools";
import { createBundledPythonEnv, type BundledPythonRuntime } from "./bundledRuntime";
import { exportDocxToPdf } from "./documentExport";
import { generateDiagramImage, type DiagramKind } from "./imageGeneration";
import {
  createWordDocxFromTemplate,
  replaceWordSectionContent,
  replaceWordSectionsContent,
  updateWordTemplateContent,
  writeSchemeDocxFromTemplate,
  type ContentControlReplacementInput,
  type SchemeCompletenessResult,
  type SchemeDiagramAsset,
  type TemplateCellReplacementInput,
  type TemplateTableReplacementInput
} from "./schemeDocument";
import type { SchemeProgressUpdateInput } from "./schemeProgress";
import { buildSchemeStandardReferenceContext } from "./schemeReference";
import {
  BUILT_IN_STANDARD_REFERENCE_RELATIVE_PATH,
  BUILT_IN_TEMPLATE_DOCX_RELATIVE_PATH,
  BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH,
  getBuiltInTemplateDocxPath,
  getBuiltInTemplateJsonPath,
  isBuiltInTemplateJsonPath
} from "./templatePaths";
import {
  DRAFT_SECTION_PARALLELISM_MAX,
  clampDraftSectionParallelism,
  clampImageGenerationParallelism
} from "../shared/types";

const execAsync = promisify(exec);
const imageGenerationQueue: Array<{ limit: number; resolve: () => void }> = [];
let activeImageGenerations = 0;
const VISION_IMAGE_MAX_BYTES = 20 * 1024 * 1024;
const NETWORK_CHANNEL_RULE =
  "网络通道/通信信道按“访问者通过网络访问系统”的形式定义，例如“业务用户通过互联网访问{应用系统}的通信信道”；访问者可为业务用户、管理用户、运维人员或第三方系统，网络可为互联网、政务外网、内网、VPN、专线或运维网。";

export interface AgentToolExecutionContext {
  rootDir: string;
  docsDir: string;
  outputDir: string;
  sessionTitle: string;
  memory: string;
  settings: AppSettings;
  userPrompt?: string;
  schemeProgress?: SchemeProgressItem;
  signal?: AbortSignal;
  allowedReadDirs: string[];
  allowedReadFiles?: string[];
  execBashEnabled: boolean;
  bundledPythonRuntime?: BundledPythonRuntime;
}

export interface AgentToolExecutionResult {
  toolName: BuiltinToolName;
  summary: string;
  content: string;
  artifactPath?: string;
  schemeProgressUpdates?: SchemeProgressUpdateInput[];
}

export interface WebSearchResult {
  title: string;
  url: string;
  snippet: string;
}

export function buildAgentChatTools(options: { includeExecBash: boolean; includeArtifactTools?: boolean }): ChatCompletionTool[] {
  const tools: ChatCompletionTool[] = [
    {
      type: "function",
      function: {
        name: "remember_project",
        description:
          "沉淀本轮对话中已经确认的项目事实、待补充信息和章节生成进度。用于持续维护项目档案；信息收集阶段应优先使用它，而不是直接生成 Word 或图片。",
        parameters: {
          type: "object",
          properties: {
            summary: {
              type: "string",
              description: "一句话概括当前已确认的项目情况。"
            },
            facts: {
              type: "array",
              description: "已确认的项目事实键值对，例如应用系统、建设单位、单位省份、等保级别、部署模式、应用子系统、关键数据、密码产品。",
              items: {
                type: "object",
                properties: {
                  key: {
                    type: "string",
                    description: "事实名称。"
                  },
                  value: {
                    type: "string",
                    description: "事实内容。"
                  }
                },
                required: ["key", "value"],
                additionalProperties: false
              }
            },
            gaps: {
              type: "array",
              description: "仍需用户补充或确认的信息；没有缺口时传空数组。不要为了生成而虚构缺失事实。",
              items: {
                type: "string"
              }
            },
            ready_for_generation: {
              type: "boolean",
              description: "当核心事实已足够支撑生成正式方案时为 true；仍需澄清时必须为 false。"
            }
          },
          required: ["summary", "facts", "gaps", "ready_for_generation"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "time",
        description: "获取当前北京时间，用于方案编制日期、时间敏感资料判断。",
        parameters: {
          type: "object",
          properties: {},
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "web_search",
        description: "检索互联网公开信息。只在需要查询最新政策、标准、产品背景或用户明确要求联网时使用。",
        parameters: {
          type: "object",
          properties: {
            query: {
              type: "string",
              description: "搜索关键词，建议包含中文标准号、政策名称、产品名称或机构名称。"
            },
            max_results: {
              type: "integer",
              description: "返回结果数量，默认 5，最大 8。"
            }
          },
          required: ["query"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "read_file",
        description:
          `读取 docs、data/input、data/output 范围内的文本文件内容；Word/PDF 优先使用 read_word/read_pdf。读取 ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH} 时返回规范化的章节规划任务清单，而不是原始 JSON。`,
        parameters: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "文件路径。可以是绝对路径，也可以是相对项目根目录的路径。"
            }
          },
          required: ["path"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "read_word",
        description: "读取 docs、data/input、data/output 范围内的 .docx Word 文件，提取正文文本用于方案生成。",
        parameters: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "Word 文件路径。可以是绝对路径，也可以是相对项目根目录的路径。"
            }
          },
          required: ["path"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "read_pdf",
        description: "读取 docs、data/input、data/output 范围内的 .pdf 文件，提取正文文本用于方案生成。",
        parameters: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "PDF 文件路径。可以是绝对路径，也可以是相对项目根目录的路径。"
            }
          },
          required: ["path"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "read_image",
        description:
          "识别用户上传的图片或截图内容。用户指出截图里有问题、要求根据截图修改方案/界面/文档，或需要读取图片中的文字、布局、标注、错误现象时使用。不要只凭图片文件名猜测内容。",
        parameters: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "图片路径。可以是绝对路径，也可以是相对项目根目录的路径；通常来自用户附件列表。"
            },
            question: {
              type: "string",
              description:
                "希望识别或定位的问题。截图修改场景应写清要关注的对象，例如“找出方案截图中哪里不符合用户要求，并给出修改建议”。"
            },
            detail: {
              type: "string",
              enum: ["auto", "low", "high"],
              description: "视觉解析精度，默认 auto；截图文字较多或细节较小时使用 high。"
            }
          },
          required: ["path"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "write_file",
        description: "把中间分析、清单或方案片段写入 data/output，并登记为前端文件卡片。",
        parameters: {
          type: "object",
          properties: {
            name: {
              type: "string",
              description: "输出文件名，建议使用 .md、.txt 或 .json 后缀。"
            },
            content: {
              type: "string",
              description: "需要写入文件的 UTF-8 文本内容。"
            }
          },
          required: ["name", "content"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "plan_scheme_batches",
        description: [
          `根据 ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH} 生成稳定的章节起草批次计划，不写 Word、不生成正文。`,
          "正式生成整篇方案时，先调用本工具获得批次，再按返回的 first_draft_call 直接调用 draft_scheme_sections。这样避免模型每次只传 1 个章节。",
          "批次严格使用模板 JSON 中真实存在的 section id，并保持模板顺序；可通过 completed_sections 跳过已完成章节，通过 start_section 从指定章节继续。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            start_section: {
              type: "string",
              description: "可选。指定从某个模板章节开始规划，必须是 template.json 中存在且唯一匹配的 section id，推荐传 id。"
            },
            completed_sections: {
              type: "array",
              description: "可选。已完成或不需要再起草的章节 id，必须能在 template.json 中唯一匹配。",
              items: {
                type: "string"
              }
            },
            batch_size: {
              type: "integer",
              description: "可选。每批章节数；不传时使用设置中的方案章节并行数，范围 1-20。"
            }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "plan_scheme_assets",
        description: [
          `根据 ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH} 规划某些章节关联的表格单元格和图片生成任务，不写 Word、不生成图片。`,
          "用于正文写入后补齐表格和图位：输入 section_ids 后，返回动态表 template_tables_plan、固定表 template_cells_plan，以及需要调用 image_generate 的 figure_id/label/prompt。",
          "只使用模板 JSON 中真实存在的 sections/tables/figures；不要自行编造 table_id、figure_id、行列号。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            section_ids: {
              type: "array",
              description: "要规划表格和图片的章节 id/编号，必须来自 template.json。为空时规划所有包含 relatedTables/relatedFigures 的章节。",
              items: {
                type: "string"
              }
            },
            include_tables: {
              type: "boolean",
              description: "是否包含表格任务，默认 true；返回中会自动区分 dynamic 的 template_tables_plan 和 fixed 的 template_cells_plan。"
            },
            include_figures: {
              type: "boolean",
              description: "是否包含图片 image_generate/diagrams 任务，默认 true。"
            },
            max_items: {
              type: "integer",
              description: "兼容参数。最多选择多少个表格/图示任务，默认 120；表格单元格仍由 max_cells 分页返回。"
            },
            table_ids: {
              type: "array",
              description: "可选。只规划指定表格 id，必须来自 template.json。为空时按 section_ids 关联表格规划。",
              items: {
                type: "string"
              }
            },
            cell_offset: {
              type: "integer",
              description: "可选。表格单元格分页起点，默认 0。"
            },
            max_cells: {
              type: "integer",
              description: "可选。最多返回多少个可填表格单元格，默认 120，范围 1-200。"
            },
            figure_offset: {
              type: "integer",
              description: "可选。图示任务分页起点，默认 0。"
            },
            max_figures: {
              type: "integer",
              description: "可选。最多返回多少个图示任务，默认 20，范围 1-50；如果 include_figures=false，则不返回图示。"
            }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "draft_scheme_sections",
        description: [
          "并行起草多个 Word 模板章节的正文草稿，但不写入 Word。",
          "用于加速正式方案生成：先按模板 JSON 顺序选取待生成小节调用本工具并行生成正文；再把返回草稿合并到 write_word.sections，按章节顺序批量写入同一个 docx。",
          `sections[].section 必须来自 ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH} 的真实 sections 条目，优先传 id，例如 sec_2_2_2；不要自行拆分或编造模板中不存在的 7.2、sec_7_2 等虚拟章节。`,
          "每个章节会按 paragraph_tasks 生成 2-4 个连续段落；如果 plan_scheme_batches 返回了 paragraph_tasks，必须原样传入。",
          "本工具只生成正文段落和必要列表，不生成 Markdown 表格，不生成图片，不修改模板表格；动态表 template_tables、固定表 template_cells 和配图 image_generate/diagrams 应在所有正文写入后由 Agent 统一处理。",
          "草稿必须贴合模板章节写作提示、占位字段、关联表格、关联图示和已确认项目事实；入参只传 section、title、writing_hint、paragraph_tasks，资料不足处写待补充，不编造关键事实。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            sections: {
              type: "array",
              description:
                `要并行起草的模板章节，必须按 ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH} 的 sections 顺序传入。建议每批数量与设置中的章节并行数接近。`,
              items: {
                type: "object",
                properties: {
                  section: {
                    type: "string",
                    description: "必须是模板 JSON 中已存在的章节 id 或编号，推荐 id，例如 sec_2_2_2；不存在的虚拟章节会被拒绝。"
                  },
                  title: {
                    type: "string",
                    description: "章节标题，可省略，工具会从模板 JSON 尝试补全。"
                  },
                  writing_hint: {
                    type: "string",
                    description: "该章节写作提示，可省略，工具会从模板 JSON 尝试补全。"
                  },
                  paragraph_tasks: {
                    type: "array",
                    description: "可选。该章节内要按顺序展开的段落任务，通常直接使用 plan_scheme_batches 返回的 paragraph_tasks。",
                    items: {
                      type: "string"
                    }
                  }
                },
                required: ["section"],
                additionalProperties: false
              }
            },
            project_context: {
              type: "string",
              description: "本批章节需要特别参考的项目事实；不传时使用已沉淀项目档案和最近工具结果。"
            },
            max_parallel: {
              type: "integer",
              description: "临时覆盖本次并行起草数；不传时使用设置中的方案章节并行数，范围 1-20。"
            }
          },
          required: ["sections"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "create_word",
        description:
          `根据内置 ${BUILT_IN_TEMPLATE_DOCX_RELATIVE_PATH} 模板创建一个 Word 文件。默认直接复制模板，内容和格式与模板保持一致；可选 template_fields 替换 {字段名} 占位，template_tables 整表重建动态表格，content_controls 按当前 Word Content Control tag 或 template.json 块 id 精确替换模板控件内容。`,
        parameters: {
          type: "object",
          properties: {
            name: {
              type: "string",
              description: "输出 Word 文件名，建议以 .docx 结尾。"
            },
            fields: {
              type: "object",
              description: "可选结构化模板字段覆盖，例如 constructionUnit、subsystems、machineRooms、cryptoProducts。",
              additionalProperties: true
            },
            template_fields: {
              type: "array",
              description:
                "可选模板字段覆盖。内置模板使用 docx-templates 驱动的 {字段名} 占位，这里优先传字段名；未提供时不会改动模板内容，会生成与模板一致的 Word 副本。",
              items: {
                type: "object",
                properties: {
                  key: {
                    type: "string",
                    description: "模板字段名，例如 应用系统、建设单位、cloudPlatform；也可直接传 {建设单位}。"
                  },
                  value: {
                    type: "string",
                    description: "模板字段值。"
                  }
                },
                required: ["key", "value"],
                additionalProperties: false
              }
            },
            content_controls: {
              type: "array",
              description:
                "可选 Content Control tag 精准替换。用于按 Word 模板里的 ps:* 锚点或 template.json 块 id 直接替换局部正文块或模板块，保留该控件外层样式和结构。固定字段必须使用 template_fields。",
              items: {
                type: "object",
                properties: {
                  tag: {
                    type: "string",
                    description:
                      "Word Content Control 的当前 tag 或 template.json 中的块 id，例如 sec_2_2_2_text_1、field_block_front_9、ps:section:sec_1:body。不要传字段名或旧标签。"
                  },
                  value: {
                    type: "string",
                    description: "写入内容；可包含多行文本或简单 Markdown 表格。"
                  }
                },
                required: ["tag", "value"],
                additionalProperties: false
              }
            },
            template_cells: {
              type: "array",
              description:
                `可选模板表格单元格替换。按 ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH} 的 table_id 找到表格模板锚点，再按 0 基行列坐标精准替换单元格，保留单元格格式。`,
              items: {
                type: "object",
                properties: {
                  table_id: {
                    type: "string",
                    description: "模板 JSON 中的表格 ID，例如 table_4_2_2_1。"
                  },
                  caption: {
                    type: "string",
                    description: "表格题注，仅用于在模板 JSON 中选择表格；Word 内定位使用表格 anchors.table 的不可见 SDT 锚。"
                  },
                  row_index: {
                    type: "number",
                    description: "0 基行号。"
                  },
                  cell_index: {
                    type: "number",
                    description: "0 基单元格序号，优先于 column_index。"
                  },
                  column_index: {
                    type: "number",
                    description: "0 基逻辑列号，可用于带合并单元格的表格。"
                  },
                  value: {
                    type: "string",
                    description: "写入单元格的文本。"
                  }
                },
                required: ["row_index", "value"],
                additionalProperties: false
              }
            },
            template_tables: {
              type: "array",
              description:
                `可选模板整表替换。按 ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH} 的 table_id/caption 找到表格模板锚点，并用新的 Markdown 表格或 rows 重建整张表，适合行数不固定的动态表。`,
              items: {
                type: "object",
                properties: {
                  table_id: {
                    type: "string",
                    description: "模板 JSON 中的表格 ID，例如 table_31_5_4_9_4。"
                  },
                  caption: {
                    type: "string",
                    description: "表格题注，仅用于在模板 JSON 中选择表格；Word 内定位使用表格 anchors.table 的不可见 SDT 锚。"
                  },
                  markdown: {
                    type: "string",
                    description: "完整 Markdown 表格，首行为表头。"
                  },
                  rows: {
                    type: "array",
                    description: "二维字符串数组，首行为表头。",
                    items: {
                      type: "array",
                      items: {
                        type: "string"
                      }
                    }
                  }
                },
                additionalProperties: false
              }
            }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "write_word",
        description: [
          "增量写入 Word 文档，并登记为前端文件卡片。",
          `标准流程：先读取 ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH} 获取真实 sections、tables、figures 和 anchors；再调用 create_word 基于 ${BUILT_IN_TEMPLATE_DOCX_RELATIVE_PATH} 创建模板副本；随后把多个章节草稿放入 sections 批量写入。`,
          "正式生成优先使用 sections 批量写入多个已起草章节：一次打开 docx、替换多个不可见锚、一次保存，明显快于多次调用 write_word。",
          "正文可先由 draft_scheme_sections 并行起草，再把多个草稿合并到 sections 数组中按模板顺序一次性写入同一个 docx。",
          "表格和配图必须放在所有正文章节写入后统一补充：先调用 plan_scheme_assets；动态表用 template_tables 重建整表，固定骨架表用 template_cells 精确补单元格，再生图并嵌入。",
          "固定字段必须使用 template_fields；局部正文块或模板块可用 content_controls 按当前 Word Content Control tag 或 template.json 块 id 精确替换。",
          "如果 template.json 某节已经列出局部正文块 textBlocks，而需求只是补一句、改一段或细化局部说明，优先使用 write_word.content_controls；这样更稳定，也更利于保留后续图表锚点。",
          "兼容模式：不传 section 时可用 template_sections 按 Markdown 编号拆分章节，但正式交付不推荐一次性写入整篇长文。",
          "传入 section 时，section 必须精确匹配 template.json 中已存在的章节 id 或 number，推荐传 id，例如 sec_7；不要传模板中不存在的 7.2、sec_7_2 等虚拟章节。",
          "Word 内部定位优先使用该章节 anchors.body.tag/alias 对应的 Content Control tag（不可见 SDT 锚）；只替换 w:sdtContent，保留模板其他章节、页眉页脚、样式和编号。",
          "Markdown 表格会渲染为真实 Word 表格；可用 render_mode=full_document 优先重建正文，但如果同时需要保留模板图位或表格锚点且 Markdown 能匹配模板章节，会先按模板章节写入以保留锚点；append 追加到文末。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "要增量更新的 data/output 中已有 .docx 路径。配合 section 使用；不传则从内置模板创建新文件。"
            },
            name: {
              type: "string",
              description: "输出 Word 文件名，建议以 .docx 结尾。传 path 且不传 name 时默认原地更新该文件。"
            },
            section: {
              type: "string",
              description: `要替换的章节，必须存在于 ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH} 的 sections 中。推荐传 id，例如 sec_1_2_1；也可传已存在的编号，例如 7。不存在的 7.2/sec_7_2 会失败。`
            },
            section_title: {
              type: "string",
              description: "section 的别名；仅允许唯一精确匹配 template.json 中已有章节标题。推荐改传 section id。"
            },
            prompt: {
              type: "string",
              description: "用户需求或系统事实描述。"
            },
            content: {
              type: "string",
              description: "要写入的 Markdown 正文。传 section 时可只写该章节内容。"
            },
            sections: {
              type: "array",
              description:
                "批量章节写入。数组项必须使用 template.json 中真实存在的 section id 或 number，推荐 id。传入后会一次打开 Word、按不可见锚替换多节、一次保存。",
              items: {
                type: "object",
                properties: {
                  section: {
                    type: "string",
                    description: "模板 JSON 中已存在的章节 id 或 number，例如 sec_1_2_1、sec_7。"
                  },
                  content: {
                    type: "string",
                    description: "该章节要写入的 Markdown 正文，不包含章节标题。"
                  }
                },
                required: ["section", "content"],
                additionalProperties: false
              }
            },
            fields: {
              type: "object",
              description: "可选结构化模板字段覆盖，例如 constructionUnit、subsystems、machineRooms、cryptoProducts。",
              additionalProperties: true
            },
            template_fields: {
              type: "array",
              description:
                "可选模板字段覆盖。用于把用户已明确提供的信息写入 Word 模板的 docx-templates 字段占位 {字段名}，key 优先传字段名，可用 应用系统、建设单位、单位省份、单位地址、单位邮编、等保级别、物理机房1、物理机房1地址、物理机房2地址、云平台、密码系统产品等。",
              items: {
                type: "object",
                properties: {
                  key: {
                    type: "string",
                    description: "模板字段名，例如 应用系统、建设单位、cloudPlatform；也可直接传 {建设单位}。"
                  },
                  value: {
                    type: "string",
                    description: "模板字段值。"
                  }
                },
                required: ["key", "value"],
                additionalProperties: false
              }
            },
            template_cells: {
              type: "array",
              description:
                `可选模板表格单元格替换。按 ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH} 的 table_id 找到表格模板锚点，再按 0 基行列坐标精准替换单元格，适合只更新模板表格中的个别单元格。`,
              items: {
                type: "object",
                properties: {
                  table_id: {
                    type: "string",
                    description: "模板 JSON 中的表格 ID，例如 table_4_2_2_1。"
                  },
                  caption: {
                    type: "string",
                    description: "表格题注，仅用于在模板 JSON 中选择表格；Word 内定位使用表格 anchors.table 的不可见 SDT 锚。"
                  },
                  row_index: {
                    type: "number",
                    description: "0 基行号。"
                  },
                  cell_index: {
                    type: "number",
                    description: "0 基单元格序号，优先于 column_index。"
                  },
                  column_index: {
                    type: "number",
                    description: "0 基逻辑列号，可用于带合并单元格的表格。"
                  },
                  value: {
                    type: "string",
                    description: "写入单元格的文本。"
                  }
                },
                required: ["row_index", "value"],
                additionalProperties: false
              }
            },
            template_tables: {
              type: "array",
              description:
                `可选模板整表替换。按 ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH} 的 table_id/caption 找到表格模板锚点，并用新的 Markdown 表格或 rows 重建整张表，适合行数不固定的动态表。`,
              items: {
                type: "object",
                properties: {
                  table_id: {
                    type: "string",
                    description: "模板 JSON 中的表格 ID，例如 table_31_5_4_9_4。"
                  },
                  caption: {
                    type: "string",
                    description: "表格题注，仅用于在模板 JSON 中选择表格；Word 内定位使用表格 anchors.table 的不可见 SDT 锚。"
                  },
                  markdown: {
                    type: "string",
                    description: "完整 Markdown 表格，首行为表头。"
                  },
                  rows: {
                    type: "array",
                    description: "二维字符串数组，首行为表头。",
                    items: {
                      type: "array",
                      items: {
                        type: "string"
                      }
                    }
                  }
                },
                additionalProperties: false
              }
            },
            content_controls: {
              type: "array",
              description:
                "可选 Content Control tag 精准替换。用于按 Word 模板里的 ps:* 锚点或 template.json 块 id 直接替换局部正文块或模板块，保留该控件外层样式和结构。固定字段必须使用 template_fields。",
              items: {
                type: "object",
                properties: {
                  tag: {
                    type: "string",
                    description:
                      "Word Content Control 的当前 tag 或 template.json 中的块 id，例如 sec_2_2_2_text_1、field_block_front_9、ps:section:sec_1:body。不要传字段名或旧标签。"
                  },
                  value: {
                    type: "string",
                    description: "写入内容；可包含多行文本或简单 Markdown 表格。"
                  }
                },
                required: ["tag", "value"],
                additionalProperties: false
              }
            },
            diagrams: {
              type: "array",
              description:
                "已由 image_generate 生成的图示文件，用于嵌入 Word。可按需传入，不再强制要求完整图示清单。",
              items: {
                type: "object",
                properties: {
                  label: {
                    type: "string",
                    description: "图示名称，例如 网络架构图、网络拓扑图、密码应用技术架构图、典型业务密码应用流程图。"
                  },
                  kind: {
                    type: "string",
                    description: "图示类型，architecture 或 flow。"
                  },
                  figure_id: {
                    type: "string",
                    description: "可选。模板 JSON 中的 figure id，例如 fig_12_5_4_9_4；传入后按该不可见图片锚精确嵌入。"
                  },
                  path: {
                    type: "string",
                    description: "image_generate 返回的 data/output 图片路径。"
                  }
                },
                required: ["label", "path"],
                additionalProperties: false
              }
            },
            render_mode: {
              type: "string",
              enum: ["template_sections", "append", "full_document"],
              description:
                "兼容模式下不传 section/sections 时使用。默认 template_sections：按 Markdown 编号替换 Word 模板；append 表示追加正文；full_document 优先用正文重建文档主体，但若同时需要保留模板图位或表格锚点且 Markdown 可匹配模板章节，则会先按模板章节写入。正式方案优先传 sections 批量写入。"
            }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "image_generate",
        description:
          "生成密码应用技术架构图或业务流程图，并登记为前端文件卡片。需要多张配图时可以并行调用多个 image_generate；实际并发数受设置中的生图并行数限制。生成的图片应只包含图形内容和必要节点标签，不要在图内单独绘制图号、题注或标题。仅当用户明确要求生成方案交付物、架构图、流程图、拓扑图或配图时使用；寒暄、答疑、资料澄清阶段不要调用。",
        parameters: {
          type: "object",
          properties: {
            kind: {
              type: "string",
              enum: ["architecture", "flow"],
              description: "architecture 表示技术架构图，flow 表示业务流程图。"
            },
            label: {
              type: "string",
              description: "图示名称，例如 网络架构图、网络拓扑图、密码应用技术架构图、典型业务密码应用流程图。"
            },
            prompt: {
              type: "string",
              description: "图片生成说明，应包含系统名称、设备、流程和中文标签要求；不要要求在图内单独绘制图号、题注或标题。"
            }
          },
          required: ["kind", "prompt"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "send_file",
        description: "把 data/output 中已经存在的文件发送给前端显示为文件卡片。不再对 .docx 做章节完整性强制拦截。",
        parameters: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "data/output 中的文件路径。"
            }
          },
          required: ["path"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "write_pdf",
        description: "将 data/output 中已经存在的 Word 方案文档导出为 PDF，并登记为前端文件卡片。需要本机安装 LibreOffice/soffice。",
        parameters: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "data/output 中的 .docx 文件路径。"
            }
          },
          required: ["path"],
          additionalProperties: false
        }
      }
    }
  ];

  if (options.includeExecBash) {
    tools.push({
      type: "function",
      function: {
        name: "exec_bash",
        description:
          "执行诊断命令。仅用于非破坏性的本地检查，禁止删除、移动、覆盖系统文件。Windows 打包环境会优先注入内置 Python runtime。",
        parameters: {
          type: "object",
          properties: {
            command: {
              type: "string",
              description: "要执行的命令。"
            }
          },
          required: ["command"],
          additionalProperties: false
        }
      }
    });
  }

  const enabledTools =
    options.includeArtifactTools === false
      ? tools.filter((tool) => tool.type !== "function" || !isFinalArtifactToolName(tool.function.name))
      : tools;

  return enabledTools.map((tool) =>
    tool.type === "function"
      ? {
          ...tool,
          function: {
            ...tool.function,
            strict: true
          }
        }
      : tool
  );
}

function isFinalArtifactToolName(name: string): boolean {
  return (
    name === "plan_scheme_batches" ||
    name === "plan_scheme_assets" ||
    name === "draft_scheme_sections" ||
    name === "create_word" ||
    name === "write_word" ||
    name === "write_pdf" ||
    name === "image_generate" ||
    name === "send_file"
  );
}

export async function executeAgentToolCall(
  toolCall: ChatCompletionMessageToolCall,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  if (toolCall.type !== "function") {
    return {
      toolName: "read_file",
      summary: `暂不支持 custom tool：${toolCall.type}`,
      content: `Unsupported tool call type: ${toolCall.type}`
    };
  }

  const toolName = toolCall.function.name as BuiltinToolName;
  const args = parseToolArguments(toolCall.function.arguments);

  switch (toolName) {
    case "time": {
      const time = getCurrentTimeText();
      return { toolName, summary: time, content: time };
    }
    case "remember_project":
      return executeRememberProject(args);
    case "web_search":
      return executeWebSearch(args);
    case "read_file":
      return executeReadFile(args, context, "read_file");
    case "read_word":
      return executeReadFile(args, context, "read_word");
    case "read_pdf":
      return executeReadFile(args, context, "read_pdf");
    case "read_image":
      return executeReadImage(args, context);
    case "write_file":
      return executeWriteFile(args, context);
    case "plan_scheme_batches":
      return executePlanSchemeBatches(args, context);
    case "plan_scheme_assets":
      return executePlanSchemeAssets(args, context);
    case "draft_scheme_sections":
      return executeDraftSchemeSections(args, context);
    case "create_word":
      return executeCreateWord(args, context);
    case "write_word":
      return executeWriteWord(args, context);
    case "image_generate":
      return executeImageGenerate(args, context);
    case "send_file":
      return executeSendFile(args, context);
    case "write_pdf":
      return executeWritePdf(args, context);
    case "exec_bash":
      return executeBash(args, context);
    default:
      return {
        toolName: "read_file",
        summary: `未知工具：${String(toolName)}`,
        content: `Unknown tool: ${String(toolName)}`
      };
  }
}

export function parseToolArguments(raw: string): Record<string, unknown> {
  if (!raw.trim()) return {};
  try {
    const parsed = JSON.parse(raw) as unknown;
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? (parsed as Record<string, unknown>) : {};
  } catch {
    return {};
  }
}

export async function searchWeb(
  query: string,
  options: { maxResults?: number; fetchImpl?: typeof fetch } = {}
): Promise<WebSearchResult[]> {
  const fetchImpl = options.fetchImpl ?? fetch;
  const maxResults = Math.min(Math.max(options.maxResults ?? 5, 1), 8);
  const url = `https://duckduckgo.com/html/?q=${encodeURIComponent(query)}`;
  const response = await fetchImpl(url, {
    headers: {
      "user-agent": "PwdSafeAgent/0.1 (+https://local.agent)"
    }
  });

  if (!response.ok) {
    throw new Error(`web_search 请求失败：${response.status} ${response.statusText}`);
  }

  const html = await response.text();
  return parseDuckDuckGoHtml(html).slice(0, maxResults);
}

export function parseDuckDuckGoHtml(html: string): WebSearchResult[] {
  const results: WebSearchResult[] = [];
  const anchorMatches = Array.from(
    html.matchAll(/<a[^>]+class="[^"]*result__a[^"]*"[^>]+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi)
  );
  const blocks = anchorMatches.length
    ? anchorMatches.map((match, index) => {
        const next = anchorMatches[index + 1];
        const end = next?.index ?? html.length;
        return html.slice(match.index ?? 0, end);
      })
    : html.match(/<div[^>]+class="[^"]*result[^"]*"[\s\S]*?(?=<div[^>]+class="[^"]*result[^"]*"|<\/body>|$)/gi) ?? [];

  for (const block of blocks) {
    const linkMatch =
      block.match(/<a[^>]+class="[^"]*result__a[^"]*"[^>]+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/i) ??
      block.match(/<a[^>]+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/i);
    if (!linkMatch) continue;

    const snippetMatch = block.match(/<a[^>]+class="[^"]*result__snippet[^"]*"[^>]*>([\s\S]*?)<\/a>/i) ??
      block.match(/<div[^>]+class="[^"]*result__snippet[^"]*"[^>]*>([\s\S]*?)<\/div>/i);
    const title = cleanHtmlText(linkMatch[2]);
    const url = normalizeDuckDuckGoUrl(decodeHtml(linkMatch[1]));
    const snippet = cleanHtmlText(snippetMatch?.[1] ?? "");

    if (title && url && !results.some((item) => item.url === url)) {
      results.push({ title, url, snippet });
    }
  }

  return results;
}

function executeWebSearch(args: Record<string, unknown>): Promise<AgentToolExecutionResult> {
  const query = readStringArg(args, "query");
  const maxResults = readNumberArg(args, "max_results", 5);
  if (!query) {
    return Promise.resolve({
      toolName: "web_search",
      summary: "缺少搜索关键词",
      content: "web_search failed: missing query"
    });
  }

  return searchWeb(query, { maxResults }).then((results) => ({
    toolName: "web_search",
    summary: results.length ? `已检索“${query}”，返回 ${results.length} 条结果` : `未检索到“${query}”的公开结果`,
    content: formatWebSearchResults(query, results)
  }));
}

function executeRememberProject(args: Record<string, unknown>): AgentToolExecutionResult {
  const summary = readStringArg(args, "summary") || "已更新项目档案";
  const facts = readKeyValueListArg(args, "facts");
  const gaps = readStringListArg(args, "gaps");
  const readyForGeneration = readBooleanArg(args, "ready_for_generation", false);
  const lines = [
    "项目档案更新",
    `摘要：${summary}`,
    "已确认事实：",
    facts.length ? facts.map((fact) => `- ${fact.key}：${fact.value}`).join("\n") : "- 暂无新增事实",
    "待补充信息：",
    gaps.length ? gaps.map((gap) => `- ${gap}`).join("\n") : "- 暂无",
    `生成就绪：${readyForGeneration ? "是" : "否"}`
  ];

  return {
    toolName: "remember_project",
    summary: readyForGeneration ? "项目档案已更新，信息已基本就绪" : "项目档案已更新，继续收集信息",
    content: lines.join("\n")
  };
}

async function executeReadFile(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext,
  requestedToolName: "read_file" | "read_word" | "read_pdf"
): Promise<AgentToolExecutionResult> {
  const inputPath = readStringArg(args, "path");
  if (!inputPath) {
    return { toolName: requestedToolName, summary: "缺少文件路径", content: `${requestedToolName} failed: missing path` };
  }

  const filePath = resolveToolPath(inputPath, context);
  assertPathAllowed(filePath, context.allowedReadDirs, context.allowedReadFiles ?? [], "read_file");
  const extension = extname(filePath).toLowerCase();
  if (requestedToolName === "read_word" && extension !== ".docx") {
    return {
      toolName: "read_word",
      summary: "read_word 只能读取 .docx 文件",
      content: `read_word failed: unsupported extension ${extension || "(none)"}`
    };
  }
  if (requestedToolName === "read_pdf" && extension !== ".pdf") {
    return {
      toolName: "read_pdf",
      summary: "read_pdf 只能读取 .pdf 文件",
      content: `read_pdf failed: unsupported extension ${extension || "(none)"}`
    };
  }

  const result = await readDocumentText(filePath, 18000);
  if (requestedToolName === "read_file" && isBuiltInTemplateJson(filePath, context)) {
    const templateSummary = buildSchemeTemplateTaskSummary(filePath, context);
    return {
      toolName: result.toolName,
      summary: `已读取 ${result.sourceName}，返回 Agent 规划任务清单`,
      content: templateSummary
    };
  }
  return {
    toolName: result.toolName,
    summary: result.summary,
    content: result.content || result.summary
  };
}

async function executeReadImage(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const inputPath = readStringArg(args, "path");
  if (!inputPath) {
    return { toolName: "read_image", summary: "缺少图片路径", content: "read_image failed: missing path" };
  }

  const filePath = resolveToolPath(inputPath, context);
  assertPathAllowed(filePath, context.allowedReadDirs, context.allowedReadFiles ?? [], "read_image");
  if (!existsSync(filePath)) {
    throw new Error(`read_image 图片不存在：${basename(filePath)}`);
  }
  if (!isSupportedImageFile(filePath)) {
    const extension = extname(filePath).toLowerCase() || "(none)";
    return {
      toolName: "read_image",
      summary: "read_image 仅支持 png/jpg/jpeg/webp/gif 图片",
      content: `read_image failed: unsupported extension ${extension}`
    };
  }

  const sourceName = basename(filePath);
  const fileSize = statSync(filePath).size;
  const question =
    readStringArg(args, "question") ||
    readStringArg(args, "prompt") ||
    buildDefaultImageQuestion(context.userPrompt);
  const detail = normalizeImageDetailArg(readStringArg(args, "detail"));
  const apiKey = context.settings.openai.apiKeyConfigured ? process.env.OPENAI_API_KEY : undefined;
  if (!apiKey) {
    return {
      toolName: "read_image",
      summary: "识图需要配置 OPENAI_API_KEY",
      content: [
        "read_image failed: OPENAI_API_KEY 未配置",
        `图片：${sourceName}`,
        `格式：${getImageMimeType(filePath) || "unknown"}`,
        `大小：${formatBytes(fileSize)}`,
        "请在设置中配置 API Key，并确认 Chat 模型支持视觉输入。"
      ].join("\n")
    };
  }

  const image = await readImageDataUrl(filePath, VISION_IMAGE_MAX_BYTES);
  const client = new OpenAI({
    apiKey,
    baseURL: context.settings.openai.baseUrl,
    timeout: context.settings.openai.requestTimeoutMs
  });
  const response = await client.chat.completions.create(
    {
      model: context.settings.openai.chatModel,
      messages: buildImageUnderstandingMessages({ image, question, detail }, context),
      max_tokens: Math.min(Math.max(Math.floor(context.settings.openai.maxOutputTokens / 4), 900), 3500)
    },
    { signal: context.signal }
  );
  const content = response.choices[0]?.message?.content?.trim();
  if (!content) {
    throw new Error("模型未返回图片识别结果");
  }

  return {
    toolName: "read_image",
    summary: `已识别 ${image.sourceName}`,
    content: [
      `read_image completed: ${filePath}`,
      `图片：${image.sourceName}`,
      `格式：${image.mimeType}`,
      `大小：${formatBytes(image.size)}`,
      `问题：${question}`,
      "",
      content
    ].join("\n")
  };
}

function buildImageUnderstandingMessages(
  input: {
    image: ImageDataUrlResult;
    question: string;
    detail: "auto" | "low" | "high";
  },
  context: AgentToolExecutionContext
): ChatCompletionMessageParam[] {
  const taskContext = compactText([context.userPrompt, context.memory].filter(Boolean).join("\n\n"), 8000);
  const text = [
    `图片文件：${input.image.sourceName}`,
    taskContext ? `相关上下文：\n${taskContext}` : "",
    `用户希望你解决的问题：${input.question}`,
    "请用中文输出：1）图片内容概述；2）识别到的关键文字、界面元素或图形关系；3）与用户问题相关的异常/需要修改点；4）下一步可执行修改建议。",
    "如果图片中看不清或没有足够证据，请明确说明不确定之处，不要编造截图里不存在的内容。"
  ]
    .filter(Boolean)
    .join("\n\n");

  return [
    {
      role: "system",
      content:
        "你是截图和图片识别工具，擅长读取界面截图、文档截图、架构图和流程图。你只根据图片可见内容和给定上下文分析问题，并给出可执行结论。"
    },
    {
      role: "user",
      content: [
        { type: "text", text },
        {
          type: "image_url",
          image_url: {
            url: input.image.dataUrl,
            detail: input.detail
          }
        }
      ]
    }
  ];
}

function buildDefaultImageQuestion(userPrompt: string | undefined): string {
  const prompt = userPrompt?.trim();
  if (prompt) return `根据用户当前需求识别图片内容并定位需要修改的问题：${prompt}`;
  return "识别这张图片的主要内容、可见文字、异常点和可执行修改建议。";
}

function normalizeImageDetailArg(value: string): "auto" | "low" | "high" {
  if (value === "low" || value === "high" || value === "auto") return value;
  return "auto";
}

interface SchemeTemplateTaskJson {
  fieldGuide?: unknown[];
  sections?: unknown[];
  tables?: unknown[];
  figures?: unknown[];
  fieldBlocks?: unknown[];
  textBlocks?: unknown[];
}

interface SchemeTemplateTaskSection {
  id: string;
  number: string;
  title: string;
  writingHint?: string;
  placeholders?: string[];
  relatedTables?: string[];
  relatedFigures?: string[];
}

interface SchemeTemplateTaskTable {
  id: string;
  section?: string;
  sectionNumber?: string;
  caption?: string;
  purpose?: string;
  writeStrategy?: string;
  header?: string[];
  placeholders?: string[];
  rows?: SchemeTemplateTaskRow[];
}

interface SchemeTemplateTaskFigure {
  id: string;
  section?: string;
  sectionNumber?: string;
  recommendedLabel?: string;
  caption?: string;
  purpose?: string;
}

interface SchemeTemplateTaskFieldBlock {
  id: string;
  text: string;
  section?: string;
  sectionNumber?: string;
  placeholders?: string[];
  anchors?: {
    block?: {
      tag: string;
      alias?: string;
    };
  };
}

interface SchemeTemplateTaskTextBlock {
  id: string;
  section: string;
  sectionNumber?: string;
  order: number;
  text?: string;
  anchors?: {
    block?: {
      tag: string;
      alias?: string;
    };
  };
}

interface SchemeTemplateTaskRow {
  index: number;
  cells: SchemeTemplateTaskCell[];
}

interface SchemeTemplateTaskCell {
  rowIndex: number;
  cellIndex: number;
  columnIndex: number;
  text: string;
  placeholders?: string[];
}

function isBuiltInTemplateJson(filePath: string, context: AgentToolExecutionContext): boolean {
  return isBuiltInTemplateJsonPath(filePath, context.docsDir);
}

function buildSchemeTemplateTaskSummary(filePath: string, context: AgentToolExecutionContext): string {
  const parsed = JSON.parse(readFileSync(filePath, "utf-8")) as SchemeTemplateTaskJson;
  const sections = readTemplateTaskSections(parsed.sections);
  const tables = readTemplateTaskTables(parsed.tables);
  const figures = readTemplateTaskFigures(parsed.figures);
  const fieldBlocks = readTemplateTaskFieldBlocks(parsed.fieldBlocks);
  const textBlocks = readTemplateTaskTextBlocks(parsed.textBlocks);
  const tableMap = new Map(tables.map((table) => [table.id, table]));
  const figureMap = new Map(figures.map((figure) => [figure.id, figure]));
  const textBlockMap = groupTemplateTaskTextBlocksBySection(textBlocks);
  const batchSize = clampDraftSectionParallelism(context.settings.agent.draftSectionParallelism);

  const lines = [
    "Word 模板规划任务清单",
    "",
    "使用规则：",
    `- draft_scheme_sections 每批尽量传 ${batchSize} 个 section，严格按下方 sections 顺序放入 sections 数组；不要只传 1 个，除非用户明确要求局部更新。`,
    "- section 必须使用下方真实 id，例如 sec_2_2_2；不要编造不存在的章节。",
    "- 每节正文按 paragraph_plan 分段编写，一项任务对应一个自然段；段落之间要承接上文，不能各写各的。",
    "- tables/figures 只记录后续任务，正文阶段不要生成 Markdown 表格或图片。",
    `- ${NETWORK_CHANNEL_RULE}`,
    "- 正文草稿完成后，用 write_word.sections 按相同顺序批量写入 Word；随后调用 plan_scheme_assets 规划表格和图片任务。",
    "- dynamic 表格按 plan_scheme_assets 返回的 template_tables_plan 整表生成；fixed 表按 template_cells_plan 改写 value 后写入；图片按 image_generate_plan 生成，再用 diagrams.figure_id 精确嵌入。",
    "- content_controls[].tag 可直接传 fieldBlocks/textBlocks 的 id，例如 field_block_front_9、sec_2_2_2_text_1；工具会自动映射到当前模板的 ps:* tag。",
    "- 如果只想改某节中的一小段正文，优先使用该节的 textBlocks / write_word.content_controls 做局部替换，不必重写整节。",
    "- 某节若已列出“局部正文块”，补一句、改一段、细化说明时都先用局部块；只有需要整体改写多段结构时再用 write_word.sections。",
    "",
    "字段占位任务：",
    formatFieldGuide(parsed.fieldGuide)
  ];

  if (fieldBlocks.length) {
    lines.push("", "固定模板块：", formatTemplateTaskFieldBlocks(fieldBlocks));
  }

  lines.push("", "章节起草任务（按模板顺序）：");

  sections.forEach((section, index) => {
    lines.push(formatTemplateTaskSection(section, index + 1, tableMap, figureMap, textBlockMap));
  });

  return compactText(lines.join("\n"), 60000);
}

function readTemplateTaskSections(value: unknown): SchemeTemplateTaskSection[] {
  if (!Array.isArray(value)) return [];
  const sections: SchemeTemplateTaskSection[] = [];
  for (const item of value) {
    if (!item || typeof item !== "object") continue;
    const record = item as Record<string, unknown>;
    const id = readRecordString(record, "id");
    const number = readRecordString(record, "number");
    const title = readRecordString(record, "title");
    if (!id || !number || !title) continue;
    sections.push({
      id,
      number,
      title,
      writingHint: readRecordString(record, "writingHint") || undefined,
      placeholders: readRecordStringArray(record, "placeholders"),
      relatedTables: readRecordStringArray(record, "relatedTables"),
      relatedFigures: readRecordStringArray(record, "relatedFigures")
    });
  }
  return sections;
}

function readTemplateTaskTables(value: unknown): SchemeTemplateTaskTable[] {
  if (!Array.isArray(value)) return [];
  const tables: SchemeTemplateTaskTable[] = [];
  for (const item of value) {
    if (!item || typeof item !== "object") continue;
    const record = item as Record<string, unknown>;
    const id = readRecordString(record, "id");
    if (!id) continue;
    tables.push({
      id,
      section: readRecordString(record, "section") || undefined,
      sectionNumber: readRecordString(record, "sectionNumber") || undefined,
      caption: readRecordString(record, "caption") || undefined,
      purpose: readRecordString(record, "purpose") || undefined,
      writeStrategy: readRecordString(record, "writeStrategy") || undefined,
      header: readRecordStringArray(record, "header"),
      placeholders: readRecordStringArray(record, "placeholders"),
      rows: readTemplateTaskRows(record.rows)
    });
  }
  return tables;
}

function readTemplateTaskFigures(value: unknown): SchemeTemplateTaskFigure[] {
  if (!Array.isArray(value)) return [];
  const figures: SchemeTemplateTaskFigure[] = [];
  for (const item of value) {
    if (!item || typeof item !== "object") continue;
    const record = item as Record<string, unknown>;
    const id = readRecordString(record, "id");
    if (!id) continue;
    figures.push({
      id,
      section: readRecordString(record, "section") || undefined,
      sectionNumber: readRecordString(record, "sectionNumber") || undefined,
      recommendedLabel: readRecordString(record, "recommendedLabel") || undefined,
      caption: readRecordString(record, "caption") || undefined,
      purpose: readRecordString(record, "purpose") || undefined
    });
  }
  return figures;
}

function readTemplateTaskFieldBlocks(value: unknown): SchemeTemplateTaskFieldBlock[] {
  if (!Array.isArray(value)) return [];
  const fieldBlocks: SchemeTemplateTaskFieldBlock[] = [];

  for (const item of value) {
    if (!item || typeof item !== "object") continue;
    const record = item as Record<string, unknown>;
    const id = readRecordString(record, "id");
    const text = readRecordString(record, "text");
    if (!id || !text) continue;

    const anchors = readTemplateTaskFieldBlockAnchors(record.anchors);
    fieldBlocks.push({
      id,
      text,
      section: readRecordString(record, "section") || undefined,
      sectionNumber: readRecordString(record, "sectionNumber") || undefined,
      placeholders: readRecordStringArray(record, "placeholders"),
      ...(anchors ? { anchors } : {})
    });
  }

  return fieldBlocks;
}

function readTemplateTaskTextBlocks(value: unknown): SchemeTemplateTaskTextBlock[] {
  if (!Array.isArray(value)) return [];
  const textBlocks: SchemeTemplateTaskTextBlock[] = [];

  for (const item of value) {
    if (!item || typeof item !== "object") continue;
    const record = item as Record<string, unknown>;
    const id = readRecordString(record, "id");
    const section = readRecordString(record, "section");
    const order = readRecordNumber(record, "order", Number.NaN);
    if (!id || !section || !Number.isInteger(order)) continue;

    const anchors = readTemplateTaskFieldBlockAnchors(record.anchors);
    textBlocks.push({
      id,
      section,
      sectionNumber: readRecordString(record, "sectionNumber") || undefined,
      order,
      text: readRecordString(record, "text") || undefined,
      ...(anchors ? { anchors } : {})
    });
  }

  return textBlocks;
}

function readTemplateTaskFieldBlockAnchors(value: unknown): SchemeTemplateTaskFieldBlock["anchors"] | undefined {
  if (!value || typeof value !== "object") return undefined;
  const record = value as Record<string, unknown>;
  const block = readTemplateTaskAnchor(record.block);
  return block ? { block } : undefined;
}

function readTemplateTaskAnchor(value: unknown): { tag: string; alias?: string } | undefined {
  if (!value || typeof value !== "object") return undefined;
  const record = value as Record<string, unknown>;
  const tag = readRecordString(record, "tag");
  if (!tag) return undefined;
  return {
    tag,
    alias: readRecordString(record, "alias") || undefined
  };
}

function readTemplateTaskRows(value: unknown): SchemeTemplateTaskRow[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const rows: SchemeTemplateTaskRow[] = [];
  for (const item of value) {
    if (!item || typeof item !== "object") continue;
    const record = item as Record<string, unknown>;
    const index = readRecordNumber(record, "index", Number.NaN);
    const cells = readTemplateTaskCells(record.cells);
    if (!Number.isInteger(index) || !cells.length) continue;
    rows.push({ index, cells });
  }
  return rows.length ? rows : undefined;
}

function readTemplateTaskCells(value: unknown): SchemeTemplateTaskCell[] {
  if (!Array.isArray(value)) return [];
  const cells: SchemeTemplateTaskCell[] = [];
  for (const item of value) {
    if (!item || typeof item !== "object") continue;
    const record = item as Record<string, unknown>;
    const rowIndex = readRecordNumber(record, "rowIndex", Number.NaN);
    const cellIndex = readRecordNumber(record, "cellIndex", Number.NaN);
    const columnIndex = readRecordNumber(record, "columnIndex", Number.NaN);
    if (!Number.isInteger(rowIndex) || !Number.isInteger(cellIndex) || !Number.isInteger(columnIndex)) continue;
    cells.push({
      rowIndex,
      cellIndex,
      columnIndex,
      text: readRecordString(record, "text"),
      placeholders: readRecordStringArray(record, "placeholders")
    });
  }
  return cells;
}

function formatTemplateTaskFieldBlocks(fieldBlocks: SchemeTemplateTaskFieldBlock[]): string {
  const lines = fieldBlocks.slice(0, 12).map((fieldBlock) => {
    const scope = fieldBlock.section ?? fieldBlock.sectionNumber ?? "front";
    const placeholders = fieldBlock.placeholders?.length ? ` | 字段：${fieldBlock.placeholders.join("、")}` : "";
    const blockTags = [
      fieldBlock.anchors?.block?.tag,
      fieldBlock.anchors?.block?.alias
    ]
      .filter(Boolean)
      .slice(0, 3)
      .join(" / ");
    return `- ${fieldBlock.id} | ${scope} | 文本：${fieldBlock.text}${placeholders}${blockTags ? ` | tag：${blockTags}` : ""}`;
  });

  if (fieldBlocks.length > 12) {
    lines.push(`- 其余 ${fieldBlocks.length - 12} 个模板块见 template.json 的 fieldBlocks。`);
  }

  return lines.length ? lines.join("\n") : "- 无";
}

function formatFieldGuide(value: unknown): string {
  if (!Array.isArray(value)) return "- 无";
  const items = value
    .map((item) => {
      if (!item || typeof item !== "object") return "";
      const record = item as Record<string, unknown>;
      const key = readRecordString(record, "key");
      const description = readRecordString(record, "description");
      return key ? `- ${key}${description ? `：${description}` : ""}` : "";
    })
    .filter(Boolean);
  return items.length ? items.join("\n") : "- 无";
}

function formatTemplateTaskSection(
  section: SchemeTemplateTaskSection,
  index: number,
  tableMap: Map<string, SchemeTemplateTaskTable>,
  figureMap: Map<string, SchemeTemplateTaskFigure>,
  textBlockMap: Map<string, SchemeTemplateTaskTextBlock[]>
): string {
  const paragraphPlan = buildSectionParagraphTasks(section, { tableMap, figureMap });
  const textBlocks = textBlockMap.get(section.id) ?? [];
  const tasks = [
    `正文：${section.writingHint || `围绕“${section.title}”编写项目化正文，资料不足处写待补充。`}`,
    `段落：${paragraphPlan.join("；")}`,
    textBlocks.length ? `局部正文块：${formatTemplateTaskTextBlocks(textBlocks)}` : "",
    section.placeholders?.length ? `字段：${section.placeholders.join("、")}` : "",
    section.relatedTables?.length ? `表格：${section.relatedTables.map((id) => formatRelatedTable(id, tableMap.get(id))).join("；")}` : "",
    section.relatedFigures?.length ? `图示：${section.relatedFigures.map((id) => formatRelatedFigure(id, figureMap.get(id))).join("；")}` : ""
  ].filter(Boolean);
  return `${index}. ${section.id} | ${section.number} ${section.title}\n   task: ${tasks.join(" / ")}`;
}

function formatTemplateTaskTextBlocks(textBlocks: SchemeTemplateTaskTextBlock[]): string {
  const items = textBlocks.slice(0, 3).map((block) => {
    const tag = block.anchors?.block?.tag?.trim() || "";
    return tag ? `${block.id}（tag：${tag}）` : block.id;
  });
  if (textBlocks.length > 3) {
    items.push(`其余 ${textBlocks.length - 3} 个见 template.json`);
  }
  return items.join("、");
}

function groupTemplateTaskTextBlocksBySection(
  textBlocks: SchemeTemplateTaskTextBlock[]
): Map<string, SchemeTemplateTaskTextBlock[]> {
  const grouped = new Map<string, SchemeTemplateTaskTextBlock[]>();

  for (const textBlock of textBlocks) {
    const items = grouped.get(textBlock.section) ?? [];
    items.push(textBlock);
    grouped.set(textBlock.section, items);
  }

  for (const [section, items] of grouped) {
    grouped.set(
      section,
      [...items].sort((left, right) => left.order - right.order || left.id.localeCompare(right.id))
    );
  }

  return grouped;
}

function formatRelatedTable(id: string, table?: SchemeTemplateTaskTable): string {
  if (!table) return id;
  const name = table.caption || "关联表格";
  const details = [table.purpose, table.header?.length ? `列=${table.header.join("|")}` : "", table.placeholders?.length ? `字段=${table.placeholders.join("、")}` : ""].filter(Boolean);
  return `${name}${details.length ? `（${details.join("；")}）` : ""}`;
}

function formatRelatedFigure(id: string, figure?: SchemeTemplateTaskFigure): string {
  if (!figure) return id;
  const label = formatFigureDisplayName(figure);
  const details = [label, figure.purpose].filter(Boolean);
  return details.length ? details.join("（") + (details.length > 1 ? "）" : "") : label;
}

function buildSectionParagraphTasks(
  section: Pick<SchemeTemplateTaskSection, "title" | "writingHint" | "relatedTables" | "relatedFigures">,
  lookup?: {
    tableMap?: Map<string, SchemeTemplateTaskTable>;
    figureMap?: Map<string, SchemeTemplateTaskFigure>;
  }
): string[] {
  const text = `${section.title} ${section.writingHint || ""}`;
  const tasks: string[] = [];

  if (/风险|威胁|不足|问题/.test(text)) {
    tasks.push("承接现状说明本节分析对象和边界", "分析主要风险来源、影响路径和业务后果", "归纳需通过密码技术或管理措施控制的重点");
  } else if (/需求|适用|要求/.test(text)) {
    tasks.push("承接风险结论说明本节需求来源", "按保护对象提炼密码技术需求", "说明管理、运维或审计配套要求");
  } else if (/设计|实现|保护|措施|建设|部署|应用/.test(text)) {
    tasks.push("说明保护对象和建设目标", "描述密码措施、产品位置和调用路径", "说明与业务流程、运维管理或后续表格图示的衔接");
  } else if (/环境|现状|情况|概述|基本|组成|框架|拓扑/.test(text)) {
    tasks.push("说明本节范围和已确认对象", "描述组成、位置、边界、责任主体或数据流向", "指出资料缺口并引出后续风险或设计分析");
  } else {
    tasks.push("承接上一节说明本节主题和范围", "结合项目事实展开关键对象、关系和约束", "总结本节结论并自然引出下一节");
  }

  if (/网络|通信|通道|信道|拓扑|边界/.test(text)) {
    tasks.splice(1, 0, "按“访问者通过网络访问系统”的形式定义网络通道/通信信道");
  }

  if (section.relatedTables?.length) {
    const tableNames = formatRelatedTableNames(section.relatedTables, lookup?.tableMap);
    tasks.push(`为后续${tableNames}填充提供文字依据，不在正文中生成表格，也不要输出 table_id 或表格任务清单`);
  }
  if (section.relatedFigures?.length) {
    const figureNames = formatRelatedFigureNames(section.relatedFigures, lookup?.figureMap);
    tasks.push(`为后续${figureNames}生成提供场景说明，不在正文中生成图片，也不要输出 fig_id 或图示任务清单`);
  }

  return tasks.slice(0, 5);
}

function formatRelatedTableNames(ids: string[], tableMap?: Map<string, SchemeTemplateTaskTable>): string {
  const names = ids.map((id) => tableMap?.get(id)?.caption).filter(Boolean);
  return names.length ? `“${names.join("”“")}”` : "关联表格";
}

function formatRelatedFigureNames(ids: string[], figureMap?: Map<string, SchemeTemplateTaskFigure>): string {
  const names = ids.map((id) => {
    const figure = figureMap?.get(id);
    return figure ? formatFigureDisplayName(figure) : "";
  }).filter(Boolean);
  return names.length ? `“${names.join("”“")}”` : "关联图示";
}

function formatFigureDisplayName(figure: SchemeTemplateTaskFigure): string {
  return figure.recommendedLabel || normalizeFigureCaption(figure.caption) || figure.caption || "关联图示";
}

function readRecordString(record: Record<string, unknown>, key: string): string {
  const value = record[key];
  return typeof value === "string" ? value.trim() : "";
}

function readRecordNumber(record: Record<string, unknown>, key: string, fallback: number): number {
  const value = record[key];
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function readRecordStringArray(record: Record<string, unknown>, key: string): string[] | undefined {
  const value = record[key];
  if (!Array.isArray(value)) return undefined;
  const items = value.filter((item): item is string => typeof item === "string" && Boolean(item.trim()));
  return items.length ? items : undefined;
}

function executePlanSchemeBatches(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): AgentToolExecutionResult {
  const templateJsonPath = getBuiltInTemplateJsonPath(context.docsDir);
  const parsed = JSON.parse(readFileSync(templateJsonPath, "utf-8")) as SchemeTemplateTaskJson;
  const allSections = readTemplateTaskSections(parsed.sections);
  const tableMap = new Map(readTemplateTaskTables(parsed.tables).map((table) => [table.id, table]));
  const figureMap = new Map(readTemplateTaskFigures(parsed.figures).map((figure) => [figure.id, figure]));
  const startSection = readStringArg(args, "start_section");
  const resolvedStartSection = startSection ? resolveTemplateTaskSection(startSection, allSections) : undefined;
  if (startSection && !resolvedStartSection) {
    return {
      toolName: "plan_scheme_batches",
      summary: "起始章节不在模板中",
      content: `plan_scheme_batches failed: unknown start_section ${startSection}; use an existing sections[].id from ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH}`
    };
  }

  const skippedPlan = resolveSkippedPlanSections(args, allSections, context.schemeProgress);
  if (skippedPlan.unknown.length) {
    return {
      toolName: "plan_scheme_batches",
      summary: "跳过章节不在模板中",
      content: `plan_scheme_batches failed: unknown completed_sections ${skippedPlan.unknown.join("、")}; use existing sections[].id values from ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH}`
    };
  }

  const skippedSections = skippedPlan.skipped;
  const requestedBatchSize = readNumberArg(args, "batch_size", Number.NaN);
  const batchSize = Number.isFinite(requestedBatchSize)
    ? clampDraftSectionParallelism(requestedBatchSize)
    : clampDraftSectionParallelism(context.settings.agent.draftSectionParallelism);
  const startIndex = resolvedStartSection ? allSections.findIndex((section) => section.id === resolvedStartSection.id) : 0;
  const effectiveStartIndex = startIndex >= 0 ? startIndex : 0;
  const sections = allSections.slice(effectiveStartIndex).filter((section) => !skippedSections.has(section.id));
  const batches = chunkArray(sections, batchSize);
  const firstBatch = batches[0] ?? [];

  return {
    toolName: "plan_scheme_batches",
    summary: batches.length
      ? `已规划 ${batches.length} 个章节批次，首批 ${firstBatch.length} 个章节`
      : "没有需要起草的章节",
    content: formatSchemeBatchPlan({
      batches,
      batchSize,
      startSection,
      skippedCount: skippedSections.size,
      tableMap,
      figureMap
    })
  };
}

function resolveSkippedPlanSections(
  args: Record<string, unknown>,
  sections: SchemeTemplateTaskSection[],
  progress: SchemeProgressItem | undefined
): { skipped: Set<string>; unknown: string[] } {
  const skipped = new Set<string>();
  const unknown: string[] = [];
  const sectionIds = new Set(sections.map((section) => section.id));
  for (const section of readStringListArg(args, "completed_sections")) {
    const resolved = resolveTemplateTaskSection(section, sections);
    if (resolved) {
      skipped.add(resolved.id);
    } else {
      unknown.push(section);
    }
  }
  for (const section of progress?.sections ?? []) {
    if (section.status === "completed" || section.status === "drafted" || section.status === "running" || section.status === "drafting") {
      if (sectionIds.has(section.id)) skipped.add(section.id);
    }
  }
  return { skipped, unknown };
}

function resolveTemplateTaskSection(
  value: string,
  sections: SchemeTemplateTaskSection[]
): SchemeTemplateTaskSection | undefined {
  const normalized = normalizeDraftSectionLookup(value);
  if (!normalized) return undefined;
  const matches = sections.filter((section) =>
    [
      section.id,
      section.number,
      section.title,
      `${section.number}${section.title}`,
      `${section.number}.${section.title}`,
      `${section.number} ${section.title}`,
      `${section.number}、${section.title}`
    ].some((candidate) => normalizeDraftSectionLookup(candidate) === normalized)
  );
  return matches.length === 1 ? matches[0] : undefined;
}

function chunkArray<T>(items: T[], size: number): T[][] {
  const chunks: T[][] = [];
  for (let index = 0; index < items.length; index += size) {
    chunks.push(items.slice(index, index + size));
  }
  return chunks;
}

function formatSchemeBatchPlan(input: {
  batches: SchemeTemplateTaskSection[][];
  batchSize: number;
  startSection?: string;
  skippedCount: number;
  tableMap: Map<string, SchemeTemplateTaskTable>;
  figureMap: Map<string, SchemeTemplateTaskFigure>;
}): string {
  const firstBatch = input.batches[0] ?? [];
  const firstDraftCall = {
    sections: firstBatch.map((section) => ({
      section: section.id,
      title: section.title,
      writing_hint: section.writingHint,
      paragraph_tasks: buildSectionParagraphTasks(section, input)
    })),
    max_parallel: input.batchSize
  };
  const lines = [
    "plan_scheme_batches completed",
    `批次大小：${input.batchSize}`,
    input.startSection ? `起始章节：${input.startSection}` : "",
    input.skippedCount ? `已跳过章节数：${input.skippedCount}` : "",
    `批次数：${input.batches.length}`,
    "",
    "下一步：直接按 first_draft_call 调用 draft_scheme_sections。不要只传其中 1 个 section。",
    "",
    "first_draft_call:",
    JSON.stringify(firstDraftCall, null, 2),
    "",
    "all_batches:"
  ].filter(Boolean);

  input.batches.forEach((batch, index) => {
    lines.push(
      `BATCH ${index + 1} (${batch.length}): ${batch.map((section) => section.id).join(", ")}`,
      ...batch.map((section) =>
        [
          `- ${section.id} | ${section.number} ${section.title} | ${section.writingHint || "按模板章节主题编写正文。"}`,
          `  paragraph_tasks: ${buildSectionParagraphTasks(section, input).join("；")}`
        ].join("\n")
      )
    );
  });

  return compactText(lines.join("\n"), 60000);
}

function executePlanSchemeAssets(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): AgentToolExecutionResult {
  const templateJsonPath = getBuiltInTemplateJsonPath(context.docsDir);
  const parsed = JSON.parse(readFileSync(templateJsonPath, "utf-8")) as SchemeTemplateTaskJson;
  const sections = readTemplateTaskSections(parsed.sections);
  const tables = readTemplateTaskTables(parsed.tables);
  const figures = readTemplateTaskFigures(parsed.figures);
  const tableMap = new Map(tables.map((table) => [table.id, table]));
  const figureMap = new Map(figures.map((figure) => [figure.id, figure]));
  const requestedSections = readStringListArg(args, "section_ids");
  const resolvedSections = requestedSections.length
    ? requestedSections.map((section) => ({ input: section, section: resolveTemplateTaskSection(section, sections) }))
    : sections
        .filter((section) => section.relatedTables?.length || section.relatedFigures?.length)
        .map((section) => ({ input: section.id, section }));
  const unknownSections = resolvedSections.filter((item) => !item.section).map((item) => item.input);
  if (unknownSections.length) {
    return {
      toolName: "plan_scheme_assets",
      summary: "章节不在模板中",
      content: `plan_scheme_assets failed: unknown section_ids ${unknownSections.join("、")}; use existing sections[].id values from ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH}`
    };
  }

  const includeTables = readBooleanArg(args, "include_tables", true);
  const includeFigures = readBooleanArg(args, "include_figures", true);
  const maxItems = Math.min(Math.max(Math.trunc(readNumberArg(args, "max_items", 120)), 1), 120);
  const cellOffset = Math.min(Math.max(Math.trunc(readNumberArg(args, "cell_offset", 0)), 0), 100000);
  const maxCells = Math.min(Math.max(Math.trunc(readNumberArg(args, "max_cells", 120)), 1), 200);
  const figureOffset = Math.min(Math.max(Math.trunc(readNumberArg(args, "figure_offset", 0)), 0), 100000);
  const requestedMaxFigures = Math.trunc(readNumberArg(args, "max_figures", 20));
  const maxFigures = includeFigures ? Math.min(Math.max(requestedMaxFigures, 1), 50) : 0;
  const selectedSections = resolvedSections.map((item) => item.section).filter((section): section is SchemeTemplateTaskSection => Boolean(section));
  const requestedTableIds = readStringListArg(args, "table_ids");
  const unknownTableIds = requestedTableIds.filter((id) => !tableMap.has(id));
  if (unknownTableIds.length) {
    return {
      toolName: "plan_scheme_assets",
      summary: "表格不在模板中",
      content: `plan_scheme_assets failed: unknown table_ids ${unknownTableIds.join("、")}; use existing tables[].id values from ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH}`
    };
  }

  const tableIds = requestedTableIds.length
    ? requestedTableIds
    : uniqueTemplateIds(selectedSections.flatMap((section) => section.relatedTables ?? []));
  const plannedTables = includeTables
    ? tableIds.map((id) => tableMap.get(id)).filter((table): table is SchemeTemplateTaskTable => Boolean(table)).slice(0, maxItems)
    : [];
  const figureIds = uniqueTemplateIds(selectedSections.flatMap((section) => section.relatedFigures ?? []));
  const plannedFigures = includeFigures
    ? figureIds.map((id) => figureMap.get(id)).filter((figure): figure is SchemeTemplateTaskFigure => Boolean(figure)).slice(0, maxItems)
    : [];

  return {
    toolName: "plan_scheme_assets",
    summary: `已规划表格 ${plannedTables.length} 个、图示 ${plannedFigures.length} 个`,
    content: formatSchemeAssetPlan({
      sections: selectedSections,
      tables: plannedTables,
      figures: plannedFigures,
      cellOffset,
      maxCells,
      figureOffset,
      maxFigures
    })
  };
}

function uniqueTemplateIds(ids: string[]): string[] {
  return Array.from(new Set(ids.filter(Boolean)));
}

function formatSchemeAssetPlan(input: {
  sections: SchemeTemplateTaskSection[];
  tables: SchemeTemplateTaskTable[];
  figures: SchemeTemplateTaskFigure[];
  cellOffset: number;
  maxCells: number;
  figureOffset: number;
  maxFigures: number;
}): string {
  const templateTables = input.tables.filter((table) => isDynamicTemplateTable(table));
  const templateCellTables = input.tables.filter((table) => !isDynamicTemplateTable(table));
  const allTableCells = templateCellTables.flatMap((table) => buildTemplateCellPlan(table));
  const tableCells = allTableCells.slice(input.cellOffset, input.cellOffset + input.maxCells);
  const tableReplacements = templateTables.map((table) => buildTemplateTablePlan(table));
  const figurePage = input.figures.slice(input.figureOffset, input.figureOffset + input.maxFigures);
  const imageCalls = figurePage.map((figure) => buildImageGeneratePlan(figure));
  const diagramRefs = figurePage.map((figure) => ({
    figure_id: figure.id,
    label: figure.recommendedLabel || normalizeFigureCaption(figure.caption) || figure.id,
    kind: inferFigureKind(figure),
    path: "使用对应 image_generate completed 路径"
  }));
  const hasMoreCells = input.cellOffset + tableCells.length < allTableCells.length;
  const hasMoreFigures = input.figureOffset + figurePage.length < input.figures.length;
  const nextCall = hasMoreCells || hasMoreFigures
    ? {
        section_ids: input.sections.map((section) => section.id),
        include_tables: hasMoreCells,
        include_figures: hasMoreFigures,
        table_ids: input.tables.map((table) => table.id),
        cell_offset: hasMoreCells ? input.cellOffset + tableCells.length : input.cellOffset,
        max_cells: input.maxCells,
        figure_offset: hasMoreFigures ? input.figureOffset + figurePage.length : input.figureOffset,
        max_figures: input.maxFigures
      }
    : undefined;
  const lines = [
    "plan_scheme_assets completed",
    `章节：${input.sections.map((section) => `${section.id}(${section.number})`).join("、") || "全部关联章节"}`,
    `表格任务：${input.tables.length} 个（整表替换：${templateTables.length} 个，单元格填充：${templateCellTables.length} 个）；可填单元格：${allTableCells.length} 个；本批单元格：${tableCells.length} 个（offset=${input.cellOffset}, limit=${input.maxCells}）`,
    `图示任务：${input.figures.length} 个；本批图示：${figurePage.length} 个（offset=${input.figureOffset}, limit=${input.maxFigures}）`,
    `是否还有后续：${hasMoreCells || hasMoreFigures ? "是" : "否"}`,
    "",
    "使用规则：",
    "- dynamic 表格先根据 project_context/项目档案生成完整表内容，再通过 write_word.template_tables 写入；建议优先传 markdown，首行为表头。",
    "- fixed 骨架表先根据 project_context/项目档案把本批 template_cells_plan 的 value 建议改成具体值，再通过 write_word.template_cells 写入。",
    "- 如果 next_plan_scheme_assets_call 不为空，必须继续调用并写入下一批，直到“是否还有后续：否”。",
    "- template_cells 的 table_id、row_index、column_index/cell_index 必须原样保留；不要自行新增行列坐标。",
    "- 对 figures 先并行调用 image_generate；再在 write_word.diagrams 中传 figure_id、label、kind、path，按不可见图片锚精确嵌入。",
    "",
    "next_plan_scheme_assets_call:",
    JSON.stringify(nextCall, null, 2),
    "",
    "template_tables_plan:",
    JSON.stringify(tableReplacements, null, 2),
    "",
    "template_cells_plan:",
    JSON.stringify(tableCells, null, 2),
    "",
    "image_generate_plan:",
    JSON.stringify(imageCalls, null, 2),
    "",
    "write_word_diagrams_plan:",
    JSON.stringify(diagramRefs, null, 2)
  ];
  return compactText(lines.join("\n"), 120000);
}

function isDynamicTemplateTable(table: SchemeTemplateTaskTable): boolean {
  return table.writeStrategy === "replace_table_when_data_complete";
}

function buildTemplateTablePlan(table: SchemeTemplateTaskTable): Record<string, unknown> {
  return {
    table_id: table.id,
    caption: table.caption,
    purpose: table.purpose || "",
    header: table.header || [],
    recommended_write: "template_tables",
    markdown_template: buildTemplateTableMarkdownTemplate(table)
  };
}

function buildTemplateTableMarkdownTemplate(table: SchemeTemplateTaskTable): string {
  const header = table.header?.length ? table.header : inferTemplateTableHeader(table);
  if (!header.length) return "";
  const separator = header.map(() => "---");
  return [`| ${header.join(" | ")} |`, `| ${separator.join(" | ")} |`].join("\n");
}

function inferTemplateTableHeader(table: SchemeTemplateTaskTable): string[] {
  const headerRow = table.rows?.find((row) => row.index === 0) ?? table.rows?.[0];
  return headerRow?.cells
    .slice()
    .sort((left, right) => left.columnIndex - right.columnIndex)
    .map((cell) => cell.text?.trim() ?? "")
    .filter(Boolean) ?? [];
}

function buildTemplateCellPlan(table: SchemeTemplateTaskTable): Array<Record<string, unknown>> {
  return (table.rows ?? [])
    .flatMap((row) => row.cells)
    .filter((cell) => isFillableTemplateCell(cell))
    .map((cell) => ({
      table_id: table.id,
      caption: table.caption,
      row_index: cell.rowIndex,
      column_index: cell.columnIndex,
      cell_index: cell.cellIndex,
      header: table.header?.[cell.columnIndex] || "",
      current_text: cell.text || "【待填写】",
      value: suggestTemplateCellValue(table, cell)
    }));
}

function isFillableTemplateCell(cell: SchemeTemplateTaskCell): boolean {
  if (cell.text.includes("【待填写】")) return true;
  return Boolean(cell.placeholders?.length);
}

function suggestTemplateCellValue(table: SchemeTemplateTaskTable, cell: SchemeTemplateTaskCell): string {
  if (cell.placeholders?.length) return cell.placeholders.map((placeholder) => `{${placeholder}}`).join("、");
  const header = table.header?.[cell.columnIndex] || "";
  if (/^序号$/.test(header)) return `${cell.rowIndex}.`;
  if (/安全要求|用途|目的/.test(header)) return "结合本节保护对象和密码应用措施填写";
  if (/产品|设备|密码/.test(header)) return "结合已选密码产品填写";
  if (/存储|位置|部署/.test(header)) return "结合系统部署位置填写";
  return "结合项目事实填写";
}

function buildImageGeneratePlan(figure: SchemeTemplateTaskFigure): Record<string, unknown> {
  const label = figure.recommendedLabel || normalizeFigureCaption(figure.caption) || figure.id;
  const kind = inferFigureKind(figure);
  return {
    figure_id: figure.id,
    label,
    kind,
    prompt: [
      `生成《密码应用方案》图示，主题：${label}。`,
      figure.caption ? `模板题注仅供匹配参考：${figure.caption}。` : "",
      figure.sectionNumber ? `所属章节：${figure.sectionNumber}。` : "",
      "要求：白底、中文标签清晰、流程方向明确，节点包含应用系统、密码服务/密码设备、数据库或存储、密钥管理/证书/算法调用等关键元素；图内不要单独放图号、题注或标题，外部文档会提供题注；不要使用模糊装饰图。"
    ]
      .filter(Boolean)
      .join("")
  };
}

function normalizeFigureCaption(value?: string): string {
  return (value || "").replace(/^图\s*\d+(?:[-－]\d+)?\s*/, "").trim();
}

function inferFigureKind(figure: SchemeTemplateTaskFigure): "architecture" | "flow" {
  const text = `${figure.recommendedLabel || ""} ${figure.caption || ""} ${figure.purpose || ""}`;
  return /流程|过程|调用|读取|写入|签名|验签|鉴别/.test(text) ? "flow" : "architecture";
}

interface DraftSchemeSectionInput {
  section: string;
  title?: string;
  writingHint?: string;
  paragraphTasks?: string[];
}

interface ResolvedDraftSchemeSection {
  section: string;
  id?: string;
  number: string;
  title: string;
  writingHint?: string;
  paragraphTasks: string[];
  previousSection?: Pick<SchemeTemplateTaskSection, "id" | "number" | "title">;
  nextSection?: Pick<SchemeTemplateTaskSection, "id" | "number" | "title">;
  placeholders?: string[];
  relatedTables?: string[];
  relatedFigures?: string[];
  relatedTableSummaries?: string[];
  relatedFigureSummaries?: string[];
}

interface DraftSchemeSectionResult {
  section: ResolvedDraftSchemeSection;
  status: Extract<SchemeSectionStatus, "drafted" | "failed">;
  content?: string;
  error?: string;
}

async function executeDraftSchemeSections(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const requestedSections = readDraftSectionsArg(args);
  if (!requestedSections.length) {
    return {
      toolName: "draft_scheme_sections",
      summary: "缺少待起草章节",
      content: "draft_scheme_sections failed: missing sections"
    };
  }

  const configuredParallel = clampDraftSectionParallelism(context.settings.agent.draftSectionParallelism);
  const requestedParallel = readNumberArg(args, "max_parallel", Number.NaN);
  const maxParallel = Number.isFinite(requestedParallel)
    ? clampDraftSectionParallelism(requestedParallel)
    : configuredParallel;
  const parsedTemplate = JSON.parse(readFileSync(getBuiltInTemplateJsonPath(context.docsDir), "utf-8")) as SchemeTemplateTaskJson;
  const templateSections = readTemplateTaskSections(parsedTemplate.sections);
  const tableMap = new Map(readTemplateTaskTables(parsedTemplate.tables).map((table) => [table.id, table]));
  const figureMap = new Map(readTemplateTaskFigures(parsedTemplate.figures).map((figure) => [figure.id, figure]));
  const sections = requestedSections
    .slice(0, DRAFT_SECTION_PARALLELISM_MAX)
    .map((section) => resolveDraftSchemeSection(section, templateSections, { tableMap, figureMap }));
  const projectContext = readStringArg(args, "project_context");
  const results = await mapWithConcurrency(sections, maxParallel, async (section) => {
    try {
      const content = await draftSchemeSection(section, context, projectContext);
      return { section, status: "drafted", content } satisfies DraftSchemeSectionResult;
    } catch (error) {
      return {
        section,
        status: "failed",
        error: error instanceof Error ? error.message : String(error)
      } satisfies DraftSchemeSectionResult;
    }
  });

  const drafted = results.filter((result) => result.status === "drafted").length;
  const failed = results.length - drafted;
  const content = formatDraftSchemeSectionResults(results, {
    maxParallel,
    truncated: requestedSections.length > sections.length
  });

  return {
    toolName: "draft_scheme_sections",
    summary: failed
      ? `已并行起草 ${drafted}/${results.length} 个章节，失败 ${failed} 个`
      : `已并行起草 ${drafted} 个章节`,
    content,
    schemeProgressUpdates: results.map((result) => ({
      section: result.section.number,
      anchorIds: result.section.id ? [result.section.id] : undefined,
      status: result.status,
      detail: result.status === "drafted" ? `${result.section.number} 已起草，等待写入 Word` : `${result.section.number} 起草失败`
    }))
  };
}

function readDraftSectionsArg(args: Record<string, unknown>): DraftSchemeSectionInput[] {
  const value = args.sections;
  if (!Array.isArray(value)) return [];
  return value
    .map((item) => {
      if (typeof item === "string") {
        return { section: item.trim() };
      }
      if (!item || typeof item !== "object" || Array.isArray(item)) return undefined;
      const record = item as Record<string, unknown>;
      const section = typeof record.section === "string" ? record.section.trim() : "";
      if (!section) return undefined;
      return {
        section,
        title: typeof record.title === "string" ? record.title.trim() : undefined,
        writingHint: typeof record.writing_hint === "string" ? record.writing_hint.trim() : undefined,
        paragraphTasks: readRecordStringArray(record, "paragraph_tasks")
      };
    })
    .filter((item): item is DraftSchemeSectionInput => Boolean(item?.section));
}

function resolveDraftSchemeSection(
  input: DraftSchemeSectionInput,
  templateSections: SchemeTemplateTaskSection[],
  lookup: {
    tableMap: Map<string, SchemeTemplateTaskTable>;
    figureMap: Map<string, SchemeTemplateTaskFigure>;
  }
): ResolvedDraftSchemeSection {
  const normalizedInput = normalizeDraftSectionLookup(input.section);
  const matches = templateSections.filter((section) =>
    [
      section.id,
      section.number,
      section.title,
      `${section.number}${section.title}`,
      `${section.number}.${section.title}`,
      `${section.number} ${section.title}`,
      `${section.number}、${section.title}`
    ].some((candidate) => normalizeDraftSectionLookup(candidate) === normalizedInput)
  );
  if (matches.length !== 1) {
    throw new Error(
      matches.length
        ? `draft_scheme_sections failed: ambiguous template section ${input.section}; use the exact sections[].id`
        : `draft_scheme_sections failed: unknown template section ${input.section}; use an existing sections[].id from ${BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH}`
    );
  }
  const matched = matches[0];
  const matchedIndex = templateSections.findIndex((section) => section.id === matched.id);
  const previousSection = matchedIndex > 0 ? templateSections[matchedIndex - 1] : undefined;
  const nextSection = matchedIndex >= 0 && matchedIndex < templateSections.length - 1 ? templateSections[matchedIndex + 1] : undefined;

  return {
    section: input.section,
    id: matched.id,
    number: matched.number,
    title: input.title || matched.title,
    writingHint: input.writingHint || matched.writingHint,
    paragraphTasks: normalizeParagraphTasksForDraft(
      input.paragraphTasks?.length ? input.paragraphTasks : buildSectionParagraphTasks(matched, lookup),
      matched,
      lookup
    ),
    previousSection,
    nextSection,
    placeholders: matched.placeholders,
    relatedTables: matched.relatedTables,
    relatedFigures: matched.relatedFigures,
    relatedTableSummaries: buildRelatedTableSummaries(matched.relatedTables, lookup.tableMap),
    relatedFigureSummaries: buildRelatedFigureSummaries(matched.relatedFigures, lookup.figureMap)
  };
}

function normalizeParagraphTasksForDraft(
  tasks: string[],
  section: Pick<SchemeTemplateTaskSection, "relatedTables" | "relatedFigures">,
  lookup: {
    tableMap: Map<string, SchemeTemplateTaskTable>;
    figureMap: Map<string, SchemeTemplateTaskFigure>;
  }
): string[] {
  return tasks
    .map((task) => replaceTemplateIdsForDraft(task, section, lookup).trim())
    .filter(Boolean);
}

function replaceTemplateIdsForDraft(
  value: string,
  section: Pick<SchemeTemplateTaskSection, "relatedTables" | "relatedFigures">,
  lookup: {
    tableMap: Map<string, SchemeTemplateTaskTable>;
    figureMap: Map<string, SchemeTemplateTaskFigure>;
  }
): string {
  let nextValue = value;
  for (const id of section.relatedTables ?? []) {
    const tableName = lookup.tableMap.get(id)?.caption || "关联表格";
    nextValue = nextValue.replaceAll(id, tableName);
  }
  for (const id of section.relatedFigures ?? []) {
    const figure = lookup.figureMap.get(id);
    const figureName = figure ? formatFigureDisplayName(figure) : "关联图示";
    nextValue = nextValue.replaceAll(id, figureName);
  }
  return nextValue;
}

function buildRelatedTableSummaries(
  ids: string[] | undefined,
  tableMap: Map<string, SchemeTemplateTaskTable>
): string[] | undefined {
  const summaries = (ids ?? [])
    .map((id) => {
      const table = tableMap.get(id);
      if (!table) return "";
      const purpose = table.purpose ? `，用于${trimSentencePunctuation(table.purpose)}` : "";
      return `${table.caption || "关联表格"}${purpose}`;
    })
    .filter(Boolean);
  return summaries.length ? summaries : undefined;
}

function buildRelatedFigureSummaries(
  ids: string[] | undefined,
  figureMap: Map<string, SchemeTemplateTaskFigure>
): string[] | undefined {
  const summaries = (ids ?? [])
    .map((id) => {
      const figure = figureMap.get(id);
      if (!figure) return "";
      const purpose = figure.purpose ? `，用于${trimSentencePunctuation(figure.purpose)}` : "";
      return `${formatFigureDisplayName(figure)}${purpose}`;
    })
    .filter(Boolean);
  return summaries.length ? summaries : undefined;
}

function trimSentencePunctuation(value: string): string {
  return value.trim().replace(/[。；;,.，]+$/g, "");
}

async function draftSchemeSection(
  section: ResolvedDraftSchemeSection,
  context: AgentToolExecutionContext,
  projectContext: string
): Promise<string> {
  if (context.signal?.aborted) {
    throw new Error("draft_scheme_sections aborted");
  }

  const apiKey = context.settings.openai.apiKeyConfigured ? process.env.OPENAI_API_KEY : undefined;
  if (!apiKey) {
    return buildFallbackDraftSchemeSection(section, context, projectContext);
  }

  const client = new OpenAI({
    apiKey,
    baseURL: context.settings.openai.baseUrl,
    timeout: context.settings.openai.requestTimeoutMs
  });
  const response = await client.chat.completions.create(
    {
      model: context.settings.openai.chatModel,
      messages: buildDraftSchemeSectionMessages(section, context, projectContext),
      max_tokens: Math.min(Math.max(Math.floor(context.settings.openai.maxOutputTokens / 6), 900), 2200)
    },
    { signal: context.signal }
  );
  const content = response.choices[0]?.message?.content?.trim();
  if (!content) {
    throw new Error("模型未返回章节草稿");
  }
  return normalizeDraftContent(content);
}

function buildDraftSchemeSectionMessages(
  section: ResolvedDraftSchemeSection,
  context: AgentToolExecutionContext,
  projectContext: string
): ChatCompletionMessageParam[] {
  const facts = compactText([context.userPrompt, projectContext, context.memory].filter(Boolean).join("\n\n"), 12000);
  const standardReference = buildSchemeStandardReferenceContext(context.docsDir, {
    sectionNumber: section.number,
    sectionTitle: section.title,
    paragraphTasks: section.paragraphTasks,
    projectContext: facts
  });
  return [
    {
      role: "system",
      content: [
        "你是密码应用方案章节正文起草器，只负责起草一个模板章节的正文。",
        "输出要求：只输出可直接传给 write_word(content) 的 Markdown 正文，不要输出章节标题，不要输出代码块，不要解释你的思路。",
        "必须按 paragraph_tasks 顺序分段输出：每个任务写 1 个自然段，段落之间用空行分隔；不要把整节写成一整坨。",
        "第一段要自然承接 previous_section，最后一段要为 next_section 留出过渡；没有上下文时也要写清本段与本节主题的关系。",
        "本阶段只写正文段落和必要列表；不要生成 Markdown 表格，不要生成图片，不要写 Mermaid/SVG，不要编造表格单元格。",
        "如该节关联表格或图示，只写引入性正文，具体表格和配图将在最后由 Agent 用 template_tables/template_cells、image_generate 和 diagrams 统一生成。",
        "正文中禁止输出模板内部 ID、锚点或任务清单，例如 fig_*、table_*、ps:figure:*；需要提到图表时只写自然名称。",
        NETWORK_CHANNEL_RULE,
        "降低 AI 味：围绕本节事实写短而具体的句子，说明对象、位置、算法/产品/调用路径/安全效果；避免万能套话、重复政策背景和空泛排比。",
        "资料不足时明确写“待补充/需确认”，不得虚构建设单位、设备型号、产品名称、网络边界或密钥管理细节。"
      ].join("\n")
    },
    {
      role: "user",
      content: [
        `章节：${section.number} ${section.title}`,
        section.previousSection ? `上一节：${section.previousSection.number} ${section.previousSection.title}` : "",
        section.nextSection ? `下一节：${section.nextSection.number} ${section.nextSection.title}` : "",
        section.writingHint ? `写作提示：${section.writingHint}` : "",
        section.placeholders?.length ? `可用字段：${section.placeholders.join("、")}` : "",
        "paragraph_tasks:",
        ...section.paragraphTasks.map((task, index) => `${index + 1}. ${task}`),
        section.relatedTableSummaries?.length ? `关联表格：${section.relatedTableSummaries.join("；")}（最后统一填充，此处不生成表格，不要输出 table_id）` : "",
        section.relatedFigureSummaries?.length ? `关联图示：${section.relatedFigureSummaries.join("；")}（最后统一生成，此处不生成图片，不要输出 fig_id）` : "",
        "",
        "项目事实和上下文：",
        facts || "暂无明确项目事实。",
        "",
        standardReference ? standardReference : `标准参考路径：${BUILT_IN_STANDARD_REFERENCE_RELATIVE_PATH}`
      ]
        .filter(Boolean)
        .join("\n")
    }
  ];
}

function buildFallbackDraftSchemeSection(
  section: ResolvedDraftSchemeSection,
  context: AgentToolExecutionContext,
  projectContext: string
): string {
  const contextText = compactText([context.userPrompt, projectContext, context.memory].filter(Boolean).join("\n"), 900);
  const bridgePrefix = section.previousSection
    ? `承接“${section.previousSection.number} ${section.previousSection.title}”，`
    : "";
  const nextSuffix = section.nextSection
    ? `本节结论将为“${section.nextSection.number} ${section.nextSection.title}”提供输入。`
    : "";
  const taskParagraphs = section.paragraphTasks.map((task, index) => {
    if (index === 0) {
      return `${bridgePrefix}本节围绕“${section.title}”展开，重点是${task}。${section.writingHint ? `应结合模板要求补充：${section.writingHint}` : "应结合已确认项目事实补充具体对象、边界、产品、算法和调用路径。"}`;
    }
    if (index === section.paragraphTasks.length - 1) {
      return `围绕“${task}”，当前资料应继续核实系统边界、设备清单、产品型号和责任主体；资料不足处应在后续补充/确认。${nextSuffix}`;
    }
    return `围绕“${task}”，应使用已确认事实说明对象、位置、调用路径和安全效果，避免脱离本系统场景的泛化描述。`;
  });
  return [
    ...taskParagraphs,
    contextText ? `已确认上下文摘要：${contextText}` : "",
    section.relatedTableSummaries?.length ? `本节后续需结合${section.relatedTableSummaries.join("、")}补充表格，正文不直接生成表格。` : "",
    section.relatedFigureSummaries?.length ? `本节后续需结合${section.relatedFigureSummaries.join("、")}生成配图，正文只保留自然场景说明。` : ""
  ]
    .filter(Boolean)
    .join("\n\n");
}

function formatDraftSchemeSectionResults(
  results: DraftSchemeSectionResult[],
  options: { maxParallel: number; truncated: boolean }
): string {
  const lines = [
    "draft_scheme_sections completed",
    `并行度：${options.maxParallel}`,
    options.truncated
      ? `本批超过 ${DRAFT_SECTION_PARALLELISM_MAX} 个章节，已只处理前 ${DRAFT_SECTION_PARALLELISM_MAX} 个；请继续分批起草。`
      : "",
    "注意：以下为按 paragraph_tasks 分段后的正文草稿，不含 Markdown 表格和配图；写入 Word 时请按章节顺序合并到 write_word.sections 批量写入。"
  ].filter(Boolean);

  for (const result of results) {
    lines.push("", `## ${result.section.number} ${result.section.title}`, `状态：${result.status === "drafted" ? "已起草" : "失败"}`);
    if (result.status === "drafted") {
      lines.push(result.content || "");
    } else {
      lines.push(`错误：${result.error || "未知错误"}`);
    }
  }

  return lines.join("\n");
}

function normalizeDraftContent(content: string): string {
  const normalized = content
    .replace(/^```(?:markdown|md)?\s*/i, "")
    .replace(/```$/i, "")
    .trim();
  return stripRawTemplatePlanningText(normalized);
}

function stripRawTemplatePlanningText(content: string): string {
  return content
    .split(/\n{2,}/)
    .map((paragraph) =>
      paragraph
        .split(/\r?\n/)
        .filter((line) => !isRawTemplatePlanningLine(line))
        .join("\n")
        .trim()
    )
    .filter(Boolean)
    .join("\n\n")
    .replace(/\bfig_\d+(?:_\d+)+\b/gi, "对应图示")
    .replace(/\btable_\d+(?:_\d+)+\b/gi, "对应表格")
    .trim();
}

function isRawTemplatePlanningLine(line: string): boolean {
  const trimmed = line.trim();
  if (!trimmed) return false;
  return /^(?:[-*]\s*)?(?:fig|table)_\d+(?:_\d+)+\s*(?:[·:：,，、-]|将|用于|需要|应)/i.test(trimmed);
}

function normalizeDraftSectionLookup(value: string): string {
  return value
    .toLowerCase()
    .replace(/[第章节]/g, "")
    .replace(/[\s.。．、,，:：()（）\[\]【】《》"'“”‘’_-]+/g, "")
    .trim();
}

async function mapWithConcurrency<T, R>(
  items: T[],
  concurrency: number,
  mapper: (item: T, index: number) => Promise<R>
): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let nextIndex = 0;
  const workers = Array.from({ length: Math.min(concurrency, items.length) }, async () => {
    while (nextIndex < items.length) {
      const index = nextIndex;
      nextIndex += 1;
      results[index] = await mapper(items[index], index);
    }
  });
  await Promise.all(workers);
  return results;
}

async function executeCreateWord(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const rawName = readStringArg(args, "name") || `${context.sessionTitle}-密码应用方案.docx`;
  const safeName = sanitizeFileName(rawName.endsWith(".docx") ? rawName : `${rawName}.docx`);
  const outputPath = join(context.outputDir, `${Date.now().toString(36)}-${safeName}`);
  const templatePath = getBuiltInTemplateDocxPath(context.docsDir);
  const templateJsonPath = getBuiltInTemplateJsonPath(context.docsDir);
  const result = await createWordDocxFromTemplate(templatePath, outputPath, {
    fields: readObjectArg(args, "fields"),
    templateFields: readTemplateFieldsArg(args),
    templateJsonPath,
    templateTables: readTemplateTablesArg(args),
    templateCells: readTemplateCellsArg(args),
    contentControls: readContentControlsArg(args)
  });

  return {
    toolName: "create_word",
    summary:
      result.templateReplacementCount > 0
        ? `已基于模板创建 ${result.fileName}，替换占位符 ${result.templateReplacementCount} 处`
        : `已基于模板创建 ${result.fileName}`,
    content: [
      `create_word completed: ${result.outputPath}`,
      `模板占位符替换次数：${result.templateReplacementCount}`,
      `模板整表替换次数：${result.templateTableReplacementCount}`,
      `模板表格单元格替换次数：${result.templateCellReplacementCount}`,
      `Content Control 替换次数：${result.contentControlReplacementCount}`,
      `显式字段：${result.filledFields.join("、") || "无"}`
    ].join("\n"),
    artifactPath: result.outputPath
  };
}

async function executeWriteWord(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const content = readStringArg(args, "content");
  const sectionBatch = readWordSectionBatchArg(args);

  const prompt = readStringArg(args, "prompt");
  const templatePath = getBuiltInTemplateDocxPath(context.docsDir);
  const templateJsonPath = getBuiltInTemplateJsonPath(context.docsDir);
  const diagrams = readDiagramAssetsArg(args, context);
  const section = readSectionArg(args);
  const fields = readObjectArg(args, "fields");
  const templateFields = readTemplateFieldsArg(args);
  const templateTables = readTemplateTablesArg(args);
  const templateCells = readTemplateCellsArg(args);
  const contentControls = readContentControlsArg(args);
  const hasFieldOverrides = Boolean(fields && Object.keys(fields).length);
  const hasDirectTemplateUpdates = Boolean(
    hasFieldOverrides || templateFields?.length || templateTables?.length || templateCells?.length || contentControls.length || diagrams.length
  );

  if (sectionBatch.length) {
    const sourcePath = readWordSourcePath(args, context, templatePath);
    const outputPath = resolveWordOutputPath(args, context, sourcePath);
    const result = await replaceWordSectionsContent(sourcePath, outputPath, {
      sections: sectionBatch,
      fields,
      templateFields,
      templateTables,
      templateCells,
      contentControls,
      diagrams,
      templateJsonPath
    });

    return {
      toolName: "write_word",
      summary: `已批量更新 ${result.fileName} 的 ${result.sections.length} 个章节`,
      content: [
        `write_word completed: ${result.outputPath}`,
        `批量更新章节：${result.sections.map((item) => item.matchedHeading).join("、")}`,
        `模板锚点：${result.sections.map((item) => item.templateAnchorId).filter(Boolean).join("、") || "未使用"}`,
        `替换原内容块：${result.sections.reduce((total, item) => total + item.replacementCount, 0)}`,
        `模板占位符替换次数：${result.templateReplacementCount}`,
        `模板整表替换次数：${result.templateTableReplacementCount}`,
        `模板表格单元格替换次数：${result.templateCellReplacementCount}`,
        `Content Control 替换次数：${result.contentControlReplacementCount}`,
        `嵌入图示：${result.embeddedDiagrams.join("、") || "无"}`
      ].join("\n"),
      artifactPath: result.outputPath
    };
  }

  if (!content && !hasDirectTemplateUpdates) {
    return { toolName: "write_word", summary: "缺少写入正文", content: "write_word failed: missing content" };
  }

  if (section && content) {
    const sourcePath = readWordSourcePath(args, context, templatePath);
    const outputPath = resolveWordOutputPath(args, context, sourcePath);
    const result = await replaceWordSectionContent(sourcePath, outputPath, {
      section,
      content,
      fields,
      templateFields,
      templateTables,
      templateCells,
      contentControls,
      diagrams,
      templateJsonPath
    });

    return {
      toolName: "write_word",
      summary: `已更新 ${result.fileName} 的章节：${result.matchedHeading}`,
      content: [
        `write_word completed: ${result.outputPath}`,
        `更新章节：${result.matchedHeading}`,
        `模板锚点：${result.templateAnchorId || "未使用"}`,
        `替换原内容块：${result.replacementCount}`,
        `模板占位符替换次数：${result.templateReplacementCount}`,
        `模板整表替换次数：${result.templateTableReplacementCount}`,
        `模板表格单元格替换次数：${result.templateCellReplacementCount}`,
        `Content Control 替换次数：${result.contentControlReplacementCount}`,
        `嵌入图示：${result.embeddedDiagrams.join("、") || "无"}`
      ].join("\n"),
      artifactPath: result.outputPath
    };
  }

  const rawName = readStringArg(args, "name") || `${context.sessionTitle}-密码应用方案.docx`;
  const safeName = sanitizeFileName(rawName.endsWith(".docx") ? rawName : `${rawName}.docx`);
  const renderMode = readRenderModeArg(args);

  if (!content && hasDirectTemplateUpdates) {
    const sourcePath = readWordSourcePath(args, context, templatePath);
    const outputPath = resolveWordOutputPath(args, context, sourcePath);
    const result = await updateWordTemplateContent(sourcePath, outputPath, {
      fields,
      templateFields,
      templateTables,
      templateCells,
      contentControls,
      diagrams,
      templateJsonPath
    });

    return {
      toolName: "write_word",
      summary: `已更新 Word ${result.fileName} 的模板字段/控件`,
      content: [
        `write_word completed: ${result.outputPath}`,
        `模板占位符替换次数：${result.templateReplacementCount}`,
        `模板整表替换次数：${result.templateTableReplacementCount}`,
        `模板表格单元格替换次数：${result.templateCellReplacementCount}`,
        `Content Control 替换次数：${result.contentControlReplacementCount}`,
        `嵌入图示：${result.embeddedDiagrams.join("、") || "无"}`
      ].join("\n"),
      artifactPath: result.outputPath
    };
  }

  const outputPath = join(context.outputDir, `${Date.now().toString(36)}-${safeName}`);

  const result = await writeSchemeDocxFromTemplate(templatePath, outputPath, {
    prompt: prompt || context.sessionTitle,
    memory: context.memory,
    generatedMarkdown: content,
    fields,
    templateFields,
    templateTables,
    templateCells,
    contentControls,
    diagrams,
    renderMode,
    templateJsonPath
  });

  return {
    toolName: "write_word",
    summary: `已生成 Word ${result.fileName}，填充 ${result.filledFields.length} 项，嵌入图示 ${result.embeddedDiagrams.length} 项`,
    content: [
      `write_word completed: ${result.outputPath}`,
      `填充字段：${result.filledFields.join("、") || "无"}`,
      `模板锚点：${result.templateAnchorsUsed.join("、") || "未使用"}`,
      `模板整表替换次数：${result.templateTableReplacementCount}`,
      `模板表格单元格替换次数：${result.templateCellReplacementCount}`,
      `Content Control 替换次数：${result.contentControlReplacementCount}`,
      `嵌入图示：${result.embeddedDiagrams.join("、") || "无"}`,
      `写入模式：${result.renderMode}`
    ].join("\n"),
    artifactPath: result.outputPath
  };
}

async function executeWriteFile(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const name = sanitizeFileName(readStringArg(args, "name") || "agent-output.md");
  const content = readStringArg(args, "content");
  if (!content) {
    return { toolName: "write_file", summary: "缺少写入内容", content: "write_file failed: missing content" };
  }

  const outputPath = join(context.outputDir, `${Date.now().toString(36)}-${name}`);
  await writeUtf8File(outputPath, content);
  return {
    toolName: "write_file",
    summary: `已写入 ${basename(outputPath)}`,
    content: compactText(content, 12000),
    artifactPath: outputPath
  };
}

async function executeImageGenerate(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const maxParallel = clampImageGenerationParallelism(context.settings.agent.imageGenerationParallelism);
  return runWithImageGenerationSlot(maxParallel, async () => executeImageGenerateUnlocked(args, context));
}

async function executeImageGenerateUnlocked(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  if (!context.settings.openai.autoImageGeneration) {
    return {
      toolName: "image_generate",
      summary: "生图工具已在设置中关闭",
      content: "image_generate skipped: OPENAI_AUTO_IMAGE_GENERATION=false"
    };
  }

  const kind = normalizeDiagramKind(readStringArg(args, "kind"));
  const label = readStringArg(args, "label") || (kind === "architecture" ? "密码应用技术架构图" : "典型业务密码应用流程图");
  const prompt = readStringArg(args, "prompt") || "生成密码应用方案配图";
  const result = await generateDiagramImage(
    {
      apiKey: process.env.OPENAI_IMAGE_API_KEY || process.env.OPENAI_API_KEY,
      baseUrl: context.settings.openai.imageBaseUrl || context.settings.openai.baseUrl,
      imageModel: context.settings.openai.imageModel,
      imageSize: context.settings.openai.imageSize,
      imageQuality: context.settings.openai.imageQuality,
      requestTimeoutMs: context.settings.openai.imageRequestTimeoutMs
    },
    {
      kind,
      label,
      sessionTitle: context.sessionTitle,
      prompt,
      memory: context.memory,
      generatedMarkdown: "",
      outputDir: context.outputDir,
      artifactStamp: createToolArtifactStamp()
    },
    context.signal
  );

  return {
    toolName: "image_generate",
    summary: `已生成 ${label}：${result.fileName}`,
    content: [`image_generate completed: ${result.outputPath}`, `图示名称：${label}`, `图示类型：${kind}`].join("\n"),
    artifactPath: result.outputPath
  };
}

async function runWithImageGenerationSlot<T>(maxParallel: number, task: () => Promise<T>): Promise<T> {
  await acquireImageGenerationSlot(maxParallel);
  try {
    return await task();
  } finally {
    releaseImageGenerationSlot();
  }
}

async function acquireImageGenerationSlot(maxParallel: number): Promise<void> {
  const limit = clampImageGenerationParallelism(maxParallel);
  if (activeImageGenerations < limit) {
    activeImageGenerations += 1;
    return;
  }

  await new Promise<void>((resolve) => {
    imageGenerationQueue.push({ limit, resolve });
  });
}

function releaseImageGenerationSlot(): void {
  activeImageGenerations = Math.max(0, activeImageGenerations - 1);
  const nextIndex = imageGenerationQueue.findIndex((item) => activeImageGenerations < item.limit);
  if (nextIndex < 0) return;

  const [next] = imageGenerationQueue.splice(nextIndex, 1);
  activeImageGenerations += 1;
  next.resolve();
}

function createToolArtifactStamp(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

async function executeSendFile(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const inputPath = readStringArg(args, "path");
  if (!inputPath) {
    return { toolName: "send_file", summary: "缺少文件路径", content: "send_file failed: missing path" };
  }

  const filePath = resolveExistingOutputToolPath(inputPath, context, "send_file");
  if (!existsSync(filePath)) {
    throw new Error(`send_file 文件不存在：${basename(filePath)}`);
  }

  return {
    toolName: "send_file",
    summary: `已发送 ${basename(filePath)}`,
    content: `send_file completed: ${filePath}`,
    artifactPath: filePath
  };
}

async function executeWritePdf(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const inputPath = readStringArg(args, "path");
  if (!inputPath) {
    return { toolName: "write_pdf", summary: "缺少 Word 文件路径", content: "write_pdf failed: missing path" };
  }

  const filePath = resolveExistingOutputToolPath(inputPath, context, "write_pdf");
  const result = await exportDocxToPdf(filePath, context.outputDir, {
    libreOfficePath: context.settings.document.libreOfficePath,
    timeoutMs: context.settings.openai.requestTimeoutMs
  });

  return {
    toolName: "write_pdf",
    summary: result.summary,
    content: result.summary,
    artifactPath: result.status === "success" ? result.outputPath : undefined
  };
}

async function executeBash(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const command = readStringArg(args, "command");
  if (!context.execBashEnabled) {
    return {
      toolName: "exec_bash",
      summary: "exec_bash 未启用",
      content: "exec_bash is disabled. Set AGENT_EXEC_BASH_ENABLED=true only in trusted local development."
    };
  }
  if (!command) {
    return { toolName: "exec_bash", summary: "缺少命令", content: "exec_bash failed: missing command" };
  }
  if (looksDestructiveCommand(command)) {
    return {
      toolName: "exec_bash",
      summary: "拒绝执行高风险命令",
      content: `Rejected destructive command: ${command}`
    };
  }

  const env = context.bundledPythonRuntime
    ? createBundledPythonEnv(context.bundledPythonRuntime, process.env)
    : process.env;
  const { stdout, stderr } = await execAsync(command, {
    cwd: context.rootDir,
    env,
    timeout: 15000,
    maxBuffer: 1024 * 256,
    windowsHide: true,
    signal: context.signal
  });
  const output = compactText([stdout, stderr].filter(Boolean).join("\n"), 12000);
  return {
    toolName: "exec_bash",
    summary: output
      ? `命令执行完成${context.bundledPythonRuntime ? "（已注入内置 Python）" : ""}`
      : `命令执行完成，无输出${context.bundledPythonRuntime ? "（已注入内置 Python）" : ""}`,
    content: output || "(empty output)"
  };
}

function formatWebSearchResults(query: string, results: WebSearchResult[]): string {
  if (!results.length) return `No web search results for: ${query}`;
  return [
    `Web search results for: ${query}`,
    ...results.map((result, index) =>
      [`${index + 1}. ${result.title}`, `URL: ${result.url}`, result.snippet ? `摘要: ${result.snippet}` : ""]
        .filter(Boolean)
        .join("\n")
    )
  ].join("\n\n");
}

function cleanHtmlText(value: string): string {
  return decodeHtml(value.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim());
}

function decodeHtml(value: string): string {
  return value
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&#39;/g, "'")
    .replace(/&#x27;/g, "'")
    .replace(/&#x2F;/g, "/");
}

function normalizeDuckDuckGoUrl(rawUrl: string): string {
  if (rawUrl.startsWith("//")) return `https:${rawUrl}`;
  if (rawUrl.startsWith("/l/?")) {
    const match = rawUrl.match(/[?&]uddg=([^&]+)/);
    if (match?.[1]) return decodeURIComponent(match[1]);
  }
  return rawUrl;
}

function resolveToolPath(inputPath: string, context: AgentToolExecutionContext): string {
  if (isAbsolute(inputPath)) return resolve(inputPath);

  const normalizedInput = inputPath.replace(/\\/g, "/");
  if (normalizedInput === "docs" || normalizedInput.startsWith("docs/")) {
    return resolve(context.docsDir, normalizedInput.replace(/^docs\/?/, ""));
  }

  return resolve(context.rootDir, inputPath);
}

function resolveOutputToolPath(inputPath: string, context: AgentToolExecutionContext): string {
  if (isAbsolute(inputPath)) return resolve(inputPath);

  const rootRelativePath = resolve(context.rootDir, inputPath);
  if (isPathInside(rootRelativePath, [context.outputDir])) {
    return rootRelativePath;
  }

  return resolve(context.outputDir, inputPath);
}

function resolveExistingOutputToolPath(inputPath: string, context: AgentToolExecutionContext, toolName: string): string {
  const exactPath = resolveOutputToolPath(inputPath, context);
  assertPathInside(exactPath, [context.outputDir], toolName);
  if (existsSync(exactPath)) return exactPath;

  const generatedPath = findGeneratedOutputFileByRequestedName(inputPath, context.outputDir);
  if (generatedPath) {
    assertPathInside(generatedPath, [context.outputDir], toolName);
    return generatedPath;
  }

  return exactPath;
}

function findGeneratedOutputFileByRequestedName(inputPath: string, outputDir: string): string | undefined {
  const requestedName = sanitizeFileName(basename(inputPath.replace(/\\/g, "/")));
  if (!requestedName || !existsSync(outputDir)) return undefined;

  const matches = readdirSync(outputDir, { withFileTypes: true })
    .filter((entry) => entry.isFile() && (entry.name === requestedName || entry.name.endsWith(`-${requestedName}`)))
    .map((entry) => {
      const filePath = join(outputDir, entry.name);
      return {
        filePath,
        mtimeMs: statSync(filePath).mtimeMs
      };
    })
    .sort((left, right) => right.mtimeMs - left.mtimeMs);

  return matches[0]?.filePath;
}

function assertPathInside(filePath: string, allowedDirs: string[], toolName: string): void {
  if (isPathInside(filePath, allowedDirs)) return;
  throw new Error(`${toolName} 只能访问允许目录内的文件`);
}

function isPathInside(filePath: string, allowedDirs: string[]): boolean {
  const normalizedPath = resolve(filePath).toLowerCase();
  const normalizedDirs = allowedDirs.map((dir) => resolve(dir).toLowerCase());
  return normalizedDirs.some((dir) => normalizedPath === dir || normalizedPath.startsWith(`${dir}\\`) || normalizedPath.startsWith(`${dir}/`));
}

function assertPathAllowed(
  filePath: string,
  allowedDirs: string[],
  allowedFiles: string[],
  toolName: string
): void {
  const normalizedPath = resolve(filePath).toLowerCase();
  const normalizedFiles = allowedFiles.map((file) => resolve(file).toLowerCase());
  if (normalizedFiles.includes(normalizedPath)) return;
  assertPathInside(filePath, allowedDirs, toolName);
}

function readStringArg(args: Record<string, unknown>, key: string): string {
  const value = args[key];
  return typeof value === "string" ? value.trim() : "";
}

function readStringListArg(args: Record<string, unknown>, key: string): string[] {
  const value = args[key];
  if (!Array.isArray(value)) return [];
  return value.map((item) => (typeof item === "string" ? item.trim() : "")).filter(Boolean);
}

function readKeyValueListArg(args: Record<string, unknown>, key: string): Array<{ key: string; value: string }> {
  const value = args[key];
  if (!Array.isArray(value)) return [];
  return value
    .map((item) => {
      if (!item || typeof item !== "object" || Array.isArray(item)) return undefined;
      const record = item as Record<string, unknown>;
      const itemKey = typeof record.key === "string" ? record.key.trim() : "";
      const itemValue = typeof record.value === "string" ? record.value.trim() : "";
      return itemKey && itemValue ? { key: itemKey, value: itemValue } : undefined;
    })
    .filter((item): item is { key: string; value: string } => Boolean(item));
}

function readBooleanArg(args: Record<string, unknown>, key: string, fallback: boolean): boolean {
  const value = args[key];
  if (typeof value === "boolean") return value;
  if (typeof value === "string") return /^(1|true|yes|是|ready)$/i.test(value.trim());
  return fallback;
}

function readObjectArg(args: Record<string, unknown>, key: string): Record<string, unknown> | undefined {
  const value = args[key];
  return value && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : undefined;
}

function readTemplateFieldsArg(args: Record<string, unknown>): Array<Record<string, unknown>> | undefined {
  const value = args.template_fields ?? args.templateFields;
  if (!Array.isArray(value)) return undefined;
  return value.filter((item): item is Record<string, unknown> => Boolean(item) && typeof item === "object" && !Array.isArray(item));
}

function readTemplateTablesArg(args: Record<string, unknown>): TemplateTableReplacementInput[] | undefined {
  const value = args.template_tables ?? args.templateTables;
  if (!Array.isArray(value)) return undefined;
  const tables = value
    .map((item) => normalizeTemplateTableReplacement(item))
    .filter((item): item is TemplateTableReplacementInput => Boolean(item));
  return tables.length ? tables : undefined;
}

function readTemplateCellsArg(args: Record<string, unknown>): TemplateCellReplacementInput[] | undefined {
  const value = args.template_cells ?? args.templateCells;
  if (!Array.isArray(value)) return undefined;
  const cells = value
    .map((item) => normalizeTemplateCellReplacement(item))
    .filter((item): item is TemplateCellReplacementInput => Boolean(item));
  return cells.length ? cells : undefined;
}

function readContentControlsArg(args: Record<string, unknown>): ContentControlReplacementInput[] {
  const value = args.content_controls ?? args.contentControls;
  if (!Array.isArray(value)) return [];
  return value
    .map((item) => normalizeContentControlReplacement(item))
    .filter((item): item is ContentControlReplacementInput => Boolean(item));
}

function normalizeContentControlReplacement(item: unknown): ContentControlReplacementInput | undefined {
  if (!item || typeof item !== "object" || Array.isArray(item)) return undefined;
  const record = item as Record<string, unknown>;
  const tag = typeof record.tag === "string" ? record.tag.trim() : "";
  const value = typeof record.value === "string" ? record.value.trim() : "";
  if (!tag || !value) return undefined;
  return {
    tag,
    value
  };
}

function normalizeTemplateTableReplacement(item: unknown): TemplateTableReplacementInput | undefined {
  if (!item || typeof item !== "object" || Array.isArray(item)) return undefined;
  const record = item as Record<string, unknown>;
  const tableId = typeof (record.table_id ?? record.tableId) === "string" ? String(record.table_id ?? record.tableId).trim() : "";
  const caption = typeof record.caption === "string" ? record.caption.trim() : "";
  const markdown = typeof record.markdown === "string" ? record.markdown.trim() : "";
  const rows = normalizeTemplateTableRows(record.rows);

  if (!tableId && !caption) return undefined;
  if (!markdown && !rows?.length) return undefined;

  return {
    ...(tableId ? { tableId } : {}),
    ...(caption ? { caption } : {}),
    ...(markdown ? { markdown } : {}),
    ...(rows?.length ? { rows } : {})
  };
}

function normalizeTemplateCellReplacement(item: unknown): TemplateCellReplacementInput | undefined {
  if (!item || typeof item !== "object" || Array.isArray(item)) return undefined;
  const record = item as Record<string, unknown>;
  const rowIndex = readTemplateCellNumber(record.row_index ?? record.rowIndex);
  const cellIndex = readTemplateCellNumber(record.cell_index ?? record.cellIndex);
  const columnIndex = readTemplateCellNumber(record.column_index ?? record.columnIndex);
  const value = typeof record.value === "string" ? record.value.trim() : "";
  const tableId = typeof (record.table_id ?? record.tableId) === "string" ? String(record.table_id ?? record.tableId).trim() : "";
  const caption = typeof record.caption === "string" ? record.caption.trim() : "";

  if (!value || !Number.isInteger(rowIndex) || rowIndex < 0) return undefined;
  if (!tableId && !caption) return undefined;
  if ((!Number.isInteger(cellIndex) || cellIndex < 0) && (!Number.isInteger(columnIndex) || columnIndex < 0)) return undefined;

  return {
    ...(tableId ? { tableId } : {}),
    ...(caption ? { caption } : {}),
    rowIndex,
    ...(Number.isInteger(cellIndex) && cellIndex >= 0 ? { cellIndex } : {}),
    ...(Number.isInteger(columnIndex) && columnIndex >= 0 ? { columnIndex } : {}),
    value
  };
}

function readTemplateCellNumber(value: unknown): number {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string" && value.trim()) return Number(value.trim());
  return Number.NaN;
}

function normalizeTemplateTableRows(value: unknown): string[][] | undefined {
  if (!Array.isArray(value) || !value.length) return undefined;
  const rows = value
    .map((row) =>
      Array.isArray(row)
        ? row.map((cell) => {
            if (typeof cell === "string") return cell.trim();
            if (typeof cell === "number" || typeof cell === "boolean") return String(cell);
            return "";
          })
        : []
    )
    .filter((row) => row.length);
  return rows.length ? rows : undefined;
}

function readSectionArg(args: Record<string, unknown>): string {
  return (
    readStringArg(args, "section") ||
    readStringArg(args, "section_title") ||
    readStringArg(args, "sectionTitle") ||
    readStringArg(args, "target") ||
    readStringArg(args, "heading")
  );
}

function readWordSectionBatchArg(args: Record<string, unknown>): Array<{ section: string; content: string }> {
  const value = args.sections;
  if (!Array.isArray(value)) return [];
  return value
    .map((item) => {
      if (!item || typeof item !== "object" || Array.isArray(item)) return undefined;
      const record = item as Record<string, unknown>;
      const section = typeof record.section === "string" ? record.section.trim() : "";
      const content =
        typeof record.content === "string"
          ? record.content.trim()
          : typeof record.markdown === "string"
            ? record.markdown.trim()
            : "";
      return section && content ? { section, content } : undefined;
    })
    .filter((item): item is { section: string; content: string } => Boolean(item));
}

function readWordSourcePath(args: Record<string, unknown>, context: AgentToolExecutionContext, templatePath: string): string {
  const rawPath =
    readStringArg(args, "path") || readStringArg(args, "docx_path") || readStringArg(args, "docxPath") || readStringArg(args, "source");
  if (!rawPath) return templatePath;

  const sourcePath = resolveOutputToolPath(rawPath, context);
  assertPathInside(sourcePath, [context.outputDir], "write_word.path");
  if (!existsSync(sourcePath)) {
    throw new Error(`write_word 源 Word 不存在：${basename(sourcePath)}`);
  }
  if (extname(sourcePath).toLowerCase() !== ".docx") {
    throw new Error("write_word.path 只能更新 .docx 文件");
  }
  return sourcePath;
}

function resolveWordOutputPath(args: Record<string, unknown>, context: AgentToolExecutionContext, sourcePath: string): string {
  const rawName = readStringArg(args, "name");
  if (!rawName && isPathInside(sourcePath, [context.outputDir])) return sourcePath;

  const fallbackName = basename(sourcePath) || `${context.sessionTitle}-密码应用方案.docx`;
  const safeName = sanitizeFileName((rawName || fallbackName).endsWith(".docx") ? rawName || fallbackName : `${rawName || fallbackName}.docx`);
  return join(context.outputDir, `${Date.now().toString(36)}-${safeName}`);
}

function readRenderModeArg(args: Record<string, unknown>): "template_sections" | "append" | "full_document" {
  const value = readStringArg(args, "render_mode") || readStringArg(args, "renderMode");
  if (value === "append" || value === "full_document" || value === "template_sections") return value;
  return "template_sections";
}

function readDiagramAssetsArg(args: Record<string, unknown>, context: AgentToolExecutionContext): SchemeDiagramAsset[] {
  const value = args.diagrams;
  const explicitDiagrams = Array.isArray(value)
    ? value
        .map((item) => normalizeDiagramAsset(item, context))
        .filter((item): item is SchemeDiagramAsset => Boolean(item))
    : [];
  if (explicitDiagrams.length) return explicitDiagrams;
  return extractDiagramAssetsFromMemory(context.memory, context);
}

function normalizeDiagramAsset(item: unknown, context: AgentToolExecutionContext): SchemeDiagramAsset | undefined {
  if (!item || typeof item !== "object" || Array.isArray(item)) return undefined;
  const record = item as Record<string, unknown>;
  const label = typeof record.label === "string" ? record.label.trim() : "";
  const rawPath = typeof record.path === "string" ? record.path.trim() : "";
  if (!label || !rawPath) return undefined;

  const path = resolveOutputToolPath(rawPath, context);
  assertPathInside(path, [context.outputDir], "write_word.diagrams");
  if (!existsSync(path)) return undefined;
  return {
    label,
    kind: typeof record.kind === "string" ? record.kind.trim() : undefined,
    figureId: typeof (record.figure_id ?? record.figureId) === "string" ? String(record.figure_id ?? record.figureId).trim() : undefined,
    path
  };
}

function extractDiagramAssetsFromMemory(memory: string, context: AgentToolExecutionContext): SchemeDiagramAsset[] {
  const assets: SchemeDiagramAsset[] = [];
  const pattern = /image_generate completed:\s*(.+?)\s*\n图示名称：(.+?)(?:\n图示类型：(.+?))?(?=\n## |\nimage_generate completed:|$)/gs;
  for (const match of memory.matchAll(pattern)) {
    const rawPath = match[1]?.trim();
    const label = match[2]?.trim();
    if (!rawPath || !label) continue;
    const asset = normalizeDiagramAsset({ path: rawPath, label, kind: match[3]?.trim() }, context);
    if (asset && !assets.some((item) => item.path === asset.path)) assets.push(asset);
  }
  return assets;
}

function formatSchemeValidationResult(result: SchemeCompletenessResult): string {
  return [
    result.summary,
    result.missingSections.length ? `缺少章节：${result.missingSections.join("、")}` : "",
    result.missingDiagrams.length ? `缺少图示：${result.missingDiagrams.join("、")}` : "",
    result.unresolvedMarkers.length ? `未完成标记：${result.unresolvedMarkers.join("、")}` : "",
    `图示数量：${result.diagramCount}`
  ]
    .filter(Boolean)
    .join("\n");
}

function readNumberArg(args: Record<string, unknown>, key: string, fallback: number): number {
  const value = args[key];
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string" && value.trim() && Number.isFinite(Number(value))) return Number(value);
  return fallback;
}

function normalizeDiagramKind(value: string): DiagramKind {
  return value === "flow" ? "flow" : "architecture";
}

function looksDestructiveCommand(command: string): boolean {
  return /\b(rm\s+-rf|del\s+\/|rmdir\s+\/s|remove-item\b|format\b|shutdown\b|restart-computer\b|git\s+reset\s+--hard)\b/i.test(command);
}
