import { exec } from "node:child_process";
import { existsSync, readFileSync, readdirSync, statSync } from "node:fs";
import { promisify } from "node:util";
import { basename, dirname, extname, isAbsolute, join, relative, resolve } from "node:path";
import OpenAI from "openai";
import type { ChatCompletionMessageParam, ChatCompletionMessageToolCall, ChatCompletionTool } from "openai/resources/chat/completions";
import type { AppSettings, SchemeProgressItem } from "../shared/types";
import {
  MAX_VISION_IMAGE_BYTES,
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
import {
  buildDocumentConfigFromProfile,
  findDocumentConfigSectionGroupForSection,
  formatDocumentConfigSummary,
  getDocumentConfigPath,
  getDocumentProfilePath,
  loadDocumentConfig,
  loadDocumentProfile,
  loadDocumentProfileFromFile,
  parseDocumentProfileJson,
  type DocumentConfig,
  type DocumentConfigSectionGroupPlanInput,
  type DocumentConfigSectionGroup,
  type DocumentGenerationPlanning,
  type DocumentProfile,
  type DocumentConfigSectionLike,
  type DocumentWordTemplateConfig
} from "./documentConfig";
import { generateDiagramImage, type DiagramKind } from "./imageGeneration";
import {
  parseWordTemplateAnchorsFromDocx,
  writeSchemeDocxFromTemplate,
  type ContentControlReplacementInput,
  type SchemeDiagramAsset,
  type TemplateCellReplacementInput,
  type TemplateTableReplacementInput,
  type WordTemplateJson
} from "./schemeDocument";
import type { SchemeProgressUpdateInput } from "./schemeProgress";
import {
  BUILT_IN_TEMPLATE_DOCX_RELATIVE_PATH,
  BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH,
  getBuiltInTemplateDocxPath,
  getBuiltInTemplateJsonPath
} from "./templatePaths";
import {
  DRAFT_SECTION_PARALLELISM_MAX,
  clampDraftSectionParallelism,
  clampImageGenerationParallelism
} from "../shared/types";

const execAsync = promisify(exec);
const imageGenerationQueue: Array<{ limit: number; resolve: () => void }> = [];
let activeImageGenerations = 0;
const NETWORK_CHANNEL_RULE =
  "网络通道/通信信道按“访问者通过网络访问系统”的形式定义，例如“业务用户通过互联网访问{应用系统}的通信信道”；访问者可为业务用户、管理用户、运维人员或第三方系统，网络可为互联网、政务外网、内网、VPN、专线或运维网。";
const DOCUMENT_SECTION_DRAFT_DIRNAME = "document-sections";
const DOCUMENT_SECTION_MANIFEST_FILENAME = "manifest.json";
const DEFAULT_DOCUMENT_CONTENT_RULES = [
  "按 document-sections/manifest.json 中的章节顺序组织正文，不打乱章节边界。",
  "每个章节只保留可交付正文，删除章节元数据、工具状态、内部 ID 和过程说明。",
  "只把已确认事实写成确定表述；资料不足处保留待补充/需确认，不编造项目现状。",
  "章节之间要有自然承接，避免重复背景、空泛套话和明显 AI 腔。",
  "最终 Markdown 应作为 Word 渲染前的正文总稿，标题层级清晰、段落连续。"
];
const DEFAULT_DOCUMENT_TEMPLATE_ID = "default";

export interface AgentToolExecutionContext {
  rootDir: string;
  docsDir: string;
  outputDir: string;
  globalOutputDir?: string;
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
                  },
                  source: {
                    type: "string",
                    description: "可选。事实来源，例如用户原话、附件文件名、read_word/read_pdf/read_image 工具结果或已确认的项目档案。"
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
          "读取 docs、data/input、当前会话 data/output 和全局 document-templates 范围内的文本文件内容；Word/PDF 优先使用 read_word/read_pdf。",
        parameters: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "文件路径。可以是绝对路径、相对项目根目录的路径，或 document-templates/<id>/profile.json。"
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
        description: "读取 docs、data/input、当前会话 data/output 和全局 document-templates 范围内的 .docx Word 文件，提取正文文本用于方案生成。",
        parameters: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "Word 文件路径。可以是绝对路径、相对项目根目录的路径，或 document-templates/<id>/template.docx。"
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
        name: "update_document_profile",
        description: [
          "更新全局模板库中某个自定义模板的 profile.json，用于让 AI 按用户偏好微调生成规则。",
          "只能写入 document-templates/<id>/profile.json，不允许修改默认内置模板、会话产物或任意路径。",
          "写入前会校验 JSON 是否为有效 DocumentProfile；profile 应主要调整 globalRules/sectionRules/evidenceRules 等生成规则，模板结构由 template.json/docx 提供。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            profile_path: {
              type: "string",
              description: "要更新的 profile 路径，必须是 document-templates/<id>/profile.json。"
            },
            template_id: {
              type: "string",
              description: "可选。模板 ID；如果未传 profile_path，则写入 document-templates/<template_id>/profile.json。"
            },
            content: {
              type: "string",
              description: "完整 profile.json 内容，必须是合法 JSON。"
            },
            change_summary: {
              type: "string",
              description: "可选。简要说明本次调整了哪些生成规则。"
            }
          },
          required: ["content"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "build_document_config",
        description: [
          "根据当前会话的输入材料、项目档案和模板结构生成项目级 document-config.json。",
          "调用前先完成整体生成规划，并把规划传入 generation_plan/section_group_plans；配置文件会保存本次规划、章节组拆分、行文规则、证据规则和必需事实。",
          "适合在资料收集接近完成、但还没有正式起草全文前调用。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            title: {
              type: "string",
              description: "文档标题；不传则使用当前会话标题。"
            },
            profile: {
              type: "string",
              description:
                "文档配置 profile，默认 generic_document。若 docs/document-profiles/<profile>.json 存在，则按该 profile 的章节规则和行文规则生成配置。"
            },
            profile_path: {
              type: "string",
              description:
                "可选。直接读取某个 profile JSON，支持 docs/...、当前会话 data/output/... 或 document-templates/<id>/profile.json；前端当前选中的模板路径会通过系统上下文提供。"
            },
            project_context: {
              type: "string",
              description: "可选。额外的项目上下文摘要，用来帮助形成配置文件。"
            },
            generation_plan: {
              type: "string",
              description: "可选但推荐。AI 在正式起草前形成的整体生成规划，只写策略、结构和风险控制，不写正文。"
            },
            planning_assumptions: {
              type: "array",
              description: "可选。生成规划中需要明确保留的假设；不得把假设写成项目事实。",
              items: {
                type: "string"
              }
            },
            planning_risks: {
              type: "array",
              description: "可选。生成过程中需要重点规避的风险，例如事实不足、模板章节歧义、证据不足。",
              items: {
                type: "string"
              }
            },
            section_group_plans: {
              type: "array",
              description: "可选但推荐。按章节组保存 AI 起草前规划，可用 section_group_id、order、match 或 title 匹配 document-config.sectionGroups。",
              items: {
                type: "object",
                properties: {
                  section_group_id: {
                    type: "string",
                    description: "可选。目标章节组 ID，例如 section_group_1。"
                  },
                  match: {
                    type: "string",
                    description: "可选。用于匹配章节组标题、章节号或章节标题。"
                  },
                  title: {
                    type: "string",
                    description: "可选。目标章节组标题或规划标题。"
                  },
                  order: {
                    type: "integer",
                    description: "可选。目标章节组顺序。"
                  },
                  objective: {
                    type: "string",
                    description: "本章节组写作目标，只写规划，不写正文。"
                  },
                  outline: {
                    type: "array",
                    description: "本章节组建议提纲或段落推进顺序。",
                    items: {
                      type: "string"
                    }
                  },
                  key_points: {
                    type: "array",
                    description: "本章节组必须覆盖的关键点。",
                    items: {
                      type: "string"
                    }
                  },
                  evidence_needs: {
                    type: "array",
                    description: "本章节组需要证据支撑的事实清单。",
                    items: {
                      type: "string"
                    }
                  },
                  open_questions: {
                    type: "array",
                    description: "本章节组尚待用户或附件确认的问题。",
                    items: {
                      type: "string"
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
        name: "list_document_sections",
        description: [
          "读取当前 document-config.json 绑定的 Word 模板，生成或刷新 data/output/document-sections/manifest.json，并列出可生成的章节清单。",
          "章节清单以 template.json/docx 解析出的 section id 为主键；章节标题只用于展示和搜索，不作为最终 Word 定位主键。",
          "用于 section-first 流程：build_document_config 后先调用本工具确认章节、草稿路径、表格和图片关联，再逐节起草或微调。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            query: {
              type: "string",
              description: "可选。按章节 id、编号或标题过滤。"
            },
            status: {
              type: "string",
              enum: ["all", "pending", "drafted", "edited", "reviewed"],
              description: "可选。按章节草稿状态过滤，默认 all。"
            },
            include_assets: {
              type: "boolean",
              description: "是否在列表中展示关联表格、图示和资源路径，默认 true。"
            },
            max_sections: {
              type: "integer",
              description: "最多返回章节数，默认 80，范围 1-200。"
            }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "get_document_section",
        description: [
          "按 section id、编号或标题查询当前模板中的一个章节，返回该章节的模板属性、关联表格/图示、草稿路径和可选草稿正文。",
          "当用户要求修改某一章/某一节、补表格、补图片、查看章节规则时使用本工具，避免靠标题猜路径或猜 Word 锚点。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            section: {
              type: "string",
              description: "章节 id、编号或标题；推荐使用 list_document_sections 返回的 section id。"
            },
            include_draft: {
              type: "boolean",
              description: "是否返回 draft.md 正文，默认 true。"
            },
            include_assets: {
              type: "boolean",
              description: "是否返回 tables.json/figures.json 内容和 assets 目录信息，默认 true。"
            }
          },
          required: ["section"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "draft_document_sections",
        description: [
          "按当前 document-sections/manifest.json 中的 section id 并行起草章节正文，并把每个章节保存为 document-sections/<section>/draft.md。",
          "这是 section-first 流程的正文起草工具：先 build_document_config，再 list_document_sections，然后调用本工具分批生成章节草稿。",
          "每个章节只使用自己的模板属性、profile 章节规则、已确认事实和项目上下文；资料不足处必须保留待补充/需确认，不编造项目现状。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            section_ids: {
              type: "array",
              description: "可选。只起草指定章节，必须来自 list_document_sections 返回的 sections[].id；不传则优先起草 pending 章节。",
              items: {
                type: "string"
              }
            },
            project_context: {
              type: "string",
              description: "可选。本批章节需要特别参考的项目事实或输入材料摘要。"
            },
            max_parallel: {
              type: "integer",
              description: "临时覆盖本次并行起草数；不传时使用设置中的方案章节并行数，范围 1-20。"
            }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "update_document_section_draft",
        description: [
          "原地更新 data/output/document-sections/<section>/draft.md，用于用户或 AI 按章节微调正文。",
          "可同时把该章节的表格原始内容写入 tables.json、图示原始内容写入 figures.json；最终 Word 写入仍按 section id/table id/figure id/template tag 精准定位。",
          "只能写入当前会话 output 的 document-sections 目录，不修改全局模板库。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            section: {
              type: "string",
              description: "章节 id、编号或标题；推荐使用 list_document_sections 返回的 section id。"
            },
            content: {
              type: "string",
              description: "章节 Markdown 正文。可以带章节标题；工具会规范保存为该章节 draft.md。"
            },
            status: {
              type: "string",
              enum: ["drafted", "edited", "reviewed"],
              description: "保存后的状态，默认 edited。"
            },
            change_summary: {
              type: "string",
              description: "可选。说明本次更新原因或用户要求。"
            },
            tables_json: {
              type: "string",
              description: "可选。该章节表格的结构化 JSON 字符串，会校验为合法 JSON 后写入 tables.json。"
            },
            figures_json: {
              type: "string",
              description: "可选。该章节图示/图片的结构化 JSON 字符串，会校验为合法 JSON 后写入 figures.json。"
            }
          },
          required: ["section", "content"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "audit_document_sections",
        description: [
          "按章节审查 document-sections/<section>/draft.md 中可能缺少证据支撑的确定性项目事实，并把报告写入对应章节目录。",
          "建议在 assemble_document_sections 前调用；这样可以逐节发现正文缺失、无证据现状、待补充信息被写成事实等问题。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            section_ids: {
              type: "array",
              description: "可选。只审查指定章节 id/编号/标题；不传则审查所有已有 draft.md 的章节。",
              items: {
                type: "string"
              }
            },
            max_findings: {
              type: "integer",
              description: "可选。每个章节最多输出多少条风险项，默认 40，范围 1-120。"
            },
            include_clean_sections: {
              type: "boolean",
              description: "是否在汇总报告中列出无风险章节，默认 true。"
            }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "revise_document_sections_evidence",
        description: [
          "按章节严格修订 document-sections/<section>/draft.md 中缺少证据支撑的确定性表述，并原地保存修订后的章节草稿。",
          "本工具复用 audit_document_sections 的证据规则；只改普通正文行，表格行和全局风险项会跳过并提示人工确认。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            section_ids: {
              type: "array",
              description: "可选。只修订指定章节 id/编号/标题；不传则修订所有已有 draft.md 且存在风险的章节。",
              items: {
                type: "string"
              }
            },
            max_rewrites: {
              type: "integer",
              description: "可选。每个章节最多修订多少行，默认 40，范围 1-120。"
            }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "polish_document_sections",
        description: [
          "按章节优化 document-sections/<section>/draft.md 的语言质量和专业表达，并原地保存。",
          "本工具只做表达、结构、术语一致性和冗余压缩，不新增项目事实，不把待补充/需确认改成确定事实。",
          "建议在 audit_document_sections/revise_document_sections_evidence 后、assemble_document_sections 前调用。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            section_ids: {
              type: "array",
              description: "可选。只润色指定章节 id/编号/标题；不传则润色所有已有 draft.md 的章节。",
              items: {
                type: "string"
              }
            },
            project_context: {
              type: "string",
              description: "可选。本批章节需要保持一致的项目事实或术语摘要。"
            },
            style_rules: {
              type: "array",
              description: "可选。额外语言优化规则，例如“减少套话”“保留待补充/需确认”“统一称谓”。",
              items: {
                type: "string"
              }
            },
            max_parallel: {
              type: "integer",
              description: "临时覆盖本次并行润色数；不传时使用设置中的方案章节并行数，范围 1-20。"
            }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "assemble_document_sections",
        description: [
          "把 data/output/document-sections 中的章节 draft.md 按当前模板 section 顺序合并为终稿 Markdown。",
          "这是 section-first 流程的确定性合并工具：不靠最终 Markdown 再反向猜章节，而是直接读取 manifest 中每个 section 的草稿路径；缺失章节会保留待补充/需确认占位。",
          "合并后的 Markdown 可继续 audit_document_evidence、plan_document_assets 和 write_document_word。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            section_ids: {
              type: "array",
              description: "可选。只合并指定章节 id/编号/标题；不传则按模板顺序合并全部章节。",
              items: {
                type: "string"
              }
            },
            name: {
              type: "string",
              description: "可选。终稿 Markdown 文件名，默认使用文档标题。"
            },
            document_rules: {
              type: "array",
              description: "可选。额外终稿清理规则，会与内置规则和 document-config.globalRules 合并展示。",
              items: {
                type: "string"
              }
            },
            include_draft: {
              type: "boolean",
              description: "是否同时输出章节合并总草稿，默认 true。"
            }
          },
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "audit_document_evidence",
        description: [
          "审查终稿 Markdown 中可能缺少证据支撑的确定性项目事实，帮助降低模型幻觉。",
          "在 assemble_document_sections 之后、write_document_word 之前调用；本工具不改正文，只生成审查报告。",
          "它会读取 document-config.json 的已确认事实和待补充信息，标记疑似把未知内容写成现状的句子。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            markdown_path: {
              type: "string",
              description: "要审查的终稿 Markdown 路径，必须位于 data/output 内。"
            },
            name: {
              type: "string",
              description: "可选。审查报告 Markdown 文件名。"
            },
            max_findings: {
              type: "integer",
              description: "可选。最多输出多少条风险项，默认 80，范围 1-200。"
            }
          },
          required: ["markdown_path"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "revise_document_evidence",
        description: [
          "根据 document-config.json 的证据规则，对终稿 Markdown 中缺少证据支撑的确定性表述做严格修订。",
          "本工具会复用 audit_document_evidence 的审查逻辑，把命中的普通正文行改为“待补充/需确认”表述，并输出新的 Markdown。",
          "适合在审查报告存在高风险项时调用；修订后应再次调用 audit_document_evidence 复核，再写 Word。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            markdown_path: {
              type: "string",
              description: "要修订的终稿 Markdown 路径，必须位于 data/output 内。"
            },
            name: {
              type: "string",
              description: "可选。修订后 Markdown 文件名。"
            },
            max_rewrites: {
              type: "integer",
              description: "可选。最多修订多少行，默认 80，范围 1-200。"
            }
          },
          required: ["markdown_path"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "plan_document_assets",
        description: [
          "根据当前 document-config.json 的 Word 模板规划表格和图示任务，不写正文、不写 Word、不生成图片。",
          "用于 section-first 文档流程：assemble_document_sections 之后、write_document_word 之前调用；默认模板和自定义模板都使用同一套逻辑。",
          "返回 dynamic 表格的 template_tables_plan、固定表的 template_cells_plan，以及需要调用 image_generate 的 figure_id/label/prompt。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            section_ids: {
              type: "array",
              description: "要规划表格和图片的章节 id/编号，必须来自当前模板清单。为空时规划所有包含 relatedTables/relatedFigures 的章节。",
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
              description: "最多选择多少个表格/图示任务，默认 120；表格单元格仍由 max_cells 分页返回。"
            },
            table_ids: {
              type: "array",
              description: "可选。只规划指定表格 id，必须来自当前模板清单。为空时按 section_ids 关联表格规划。",
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
        name: "write_document_word",
        description: [
          "把 assemble_document_sections 产出的终稿 Markdown 按 document-config.json 的 wordTemplate 或指定 Word 模板生成目标 .docx。",
          "这是 section-first 文档流程的 Word 落地工具：先 build_document_config，再 list_document_sections/draft_document_sections，随后 assemble_document_sections，再 plan_document_assets/image_generate 补齐图表，最后调用本工具。",
          "未指定 template_path 时优先使用 document-config.json 的 wordTemplate.templatePath；仍未配置时使用内置 Word 模板。"
        ].join("\n"),
        parameters: {
          type: "object",
          properties: {
            markdown_path: {
              type: "string",
              description: "assemble_document_sections 返回的终稿 Markdown 路径，必须位于 data/output 内。"
            },
            name: {
              type: "string",
              description: "可选。输出 Word 文件名，建议以 .docx 结尾。"
            },
            template_path: {
              type: "string",
              description: "可选。Word 模板路径，支持 docs/... 或 data/output/...；不传则使用 document-config.json 的 wordTemplate.templatePath 或内置模板。"
            },
            template_json_path: {
              type: "string",
              description: "可选。模板结构 JSON 路径，支持 docs/... 或 data/output/...；不传则使用 document-config.json 的 wordTemplate.templateJsonPath，未配置或文件不存在时自动从 Word 模板 docx 解析锚点。"
            },
            render_mode: {
              type: "string",
              enum: ["append", "full_document", "template_sections"],
              description: "可选。Markdown 写入模式；不传则使用 document-config.json 的 wordTemplate.renderMode，仍未配置时默认 full_document。"
            },
            fields: {
              type: "object",
              description: "可选结构化模板字段覆盖。",
              additionalProperties: true
            },
            template_fields: {
              type: "array",
              description: "可选模板字段覆盖，用于替换 Word 模板中的固定字段。",
              items: {
                type: "object",
                properties: {
                  key: {
                    type: "string",
                    description: "模板字段名。"
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
            template_tables: {
              type: "array",
              description:
                "可选模板整表替换。来自 plan_document_assets 的 template_tables_plan；适合行数不固定的动态表，按 table_id/caption 找模板锚点。",
              items: {
                type: "object",
                properties: {
                  table_id: {
                    type: "string",
                    description: "模板中的表格 ID，例如 table_31_5_4_9_4。"
                  },
                  caption: {
                    type: "string",
                    description: "表格题注，仅用于辅助选择表格。"
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
            template_cells: {
              type: "array",
              description:
                "可选模板单元格替换。来自 plan_document_assets 的 template_cells_plan；固定骨架表按 table_id、row_index、column_index/cell_index 精确填值。",
              items: {
                type: "object",
                properties: {
                  table_id: {
                    type: "string",
                    description: "模板中的表格 ID。"
                  },
                  caption: {
                    type: "string",
                    description: "可选表格题注。"
                  },
                  row_index: {
                    type: "integer",
                    description: "模板表格行序号，必须保留 plan_document_assets 返回值。"
                  },
                  column_index: {
                    type: "integer",
                    description: "模板表格列序号，必须保留 plan_document_assets 返回值。"
                  },
                  cell_index: {
                    type: "integer",
                    description: "可选。Word 单元格序号，必须保留 plan_document_assets 返回值。"
                  },
                  value: {
                    type: "string",
                    description: "写入单元格的值；资料不足时写待补充/需确认。"
                  }
                },
                required: ["row_index", "value"],
                additionalProperties: false
              }
            },
            content_controls: {
              type: "array",
              description:
                "可选 Content Control tag 精准替换。用于按 Word 模板 ps:* 锚点或 template.json 块 id 替换局部正文块，固定字段仍使用 template_fields。",
              items: {
                type: "object",
                properties: {
                  tag: {
                    type: "string",
                    description: "Word Content Control tag 或 template.json 块 id，例如 sec_2_2_2_text_1、ps:section:sec_1:body。"
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
                "已由 image_generate 生成的图示文件，用于嵌入 Word。未传时工具会尝试从当前输出目录自动识别已生成图示。",
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
                    description: "可选。模板中的 figure id，例如 fig_12_5_4_9_4；传入后按该图片锚精确嵌入。"
                  },
                  path: {
                    type: "string",
                    description: "image_generate 返回的 data/output 图片路径。"
                  }
                },
                required: ["label", "path"],
                additionalProperties: false
              }
            }
          },
          required: ["markdown_path"],
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
    name === "update_document_profile" ||
    name === "build_document_config" ||
    name === "list_document_sections" ||
    name === "get_document_section" ||
    name === "draft_document_sections" ||
    name === "update_document_section_draft" ||
    name === "audit_document_sections" ||
    name === "revise_document_sections_evidence" ||
    name === "polish_document_sections" ||
    name === "assemble_document_sections" ||
    name === "audit_document_evidence" ||
    name === "revise_document_evidence" ||
    name === "plan_document_assets" ||
    name === "write_document_word" ||
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
    case "update_document_profile":
      return executeUpdateDocumentProfile(args, context);
    case "build_document_config":
      return executeBuildDocumentConfig(args, context);
    case "list_document_sections":
      return executeListDocumentSections(args, context);
    case "get_document_section":
      return executeGetDocumentSection(args, context);
    case "draft_document_sections":
      return executeDraftDocumentSections(args, context);
    case "update_document_section_draft":
      return executeUpdateDocumentSectionDraft(args, context);
    case "audit_document_sections":
      return executeAuditDocumentSections(args, context);
    case "revise_document_sections_evidence":
      return executeReviseDocumentSectionsEvidence(args, context);
    case "polish_document_sections":
      return executePolishDocumentSections(args, context);
    case "assemble_document_sections":
      return executeAssembleDocumentSections(args, context);
    case "audit_document_evidence":
      return executeAuditDocumentEvidence(args, context);
    case "revise_document_evidence":
      return executeReviseDocumentEvidence(args, context);
    case "plan_document_assets":
      return executePlanDocumentAssets(args, context);
    case "write_document_word":
      return executeWriteDocumentWord(args, context);
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
    "事实使用规则：只有“已确认事实”可在方案正文中写成确定表述；“待补充信息”必须继续保留为待补充/需确认，不能改写成已经存在的项目事实。",
    "已确认事实：",
    facts.length ? facts.map((fact) => `- ${fact.key}：${fact.value}${fact.source ? `（来源：${fact.source}）` : ""}`).join("\n") : "- 暂无新增事实",
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
  const apiKey = resolveVisionApiKey(context.settings);
  if (!apiKey) {
    return {
      toolName: "read_image",
      summary: "识图需要配置 OPENAI_VISION_API_KEY 或 OPENAI_API_KEY",
      content: [
        "read_image failed: OPENAI_VISION_API_KEY / OPENAI_API_KEY 未配置",
        `图片：${sourceName}`,
        `格式：${getImageMimeType(filePath) || "unknown"}`,
        `大小：${formatBytes(fileSize)}`,
        "请在设置中配置识图 API Key，或配置通用 API Key 作为回退。"
      ].join("\n")
    };
  }

  const image = await readImageDataUrl(filePath, MAX_VISION_IMAGE_BYTES);
  const visionModel = resolveVisionModel(context.settings);
  const client = new OpenAI({
    apiKey,
    baseURL: resolveVisionBaseUrl(context.settings),
    timeout: context.settings.openai.requestTimeoutMs
  });
  const response = await client.chat.completions.create(
    {
      model: visionModel,
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
      `识图模型：${visionModel}`,
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

function resolveVisionApiKey(settings: AppSettings): string | undefined {
  if (settings.openai.visionApiKeyConfigured && process.env.OPENAI_VISION_API_KEY) {
    return process.env.OPENAI_VISION_API_KEY;
  }
  return settings.openai.apiKeyConfigured ? process.env.OPENAI_API_KEY : undefined;
}

function resolveVisionBaseUrl(settings: AppSettings): string {
  return settings.openai.visionBaseUrl.trim() || settings.openai.baseUrl;
}

function resolveVisionModel(settings: AppSettings): string {
  return settings.openai.visionModel.trim() || settings.openai.chatModel;
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
  anchors?: {
    body?: {
      tag: string;
      alias?: string;
    };
  };
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

async function loadSchemeTemplateTaskJson(context: AgentToolExecutionContext): Promise<SchemeTemplateTaskJson> {
  const templateJsonPath = getBuiltInTemplateJsonPath(context.docsDir);
  const parsedJson = readSchemeTemplateTaskJsonIfPresent(templateJsonPath);
  if (parsedJson) return parsedJson;

  try {
    const parsedDocx = await parseWordTemplateAnchorsFromDocx(getBuiltInTemplateDocxPath(context.docsDir));
    return buildSchemeTemplateTaskJsonFromWordTemplate(parsedDocx);
  } catch {
    return {};
  }
}

function readSchemeTemplateTaskJsonIfPresent(filePath: string): SchemeTemplateTaskJson | undefined {
  if (!existsSync(filePath)) return undefined;
  try {
    return JSON.parse(readFileSync(filePath, "utf-8")) as SchemeTemplateTaskJson;
  } catch {
    return undefined;
  }
}

function buildSchemeTemplateTaskJsonFromWordTemplate(templateJson: WordTemplateJson | undefined): SchemeTemplateTaskJson {
  const baseSections: SchemeTemplateTaskSection[] = (templateJson?.sections ?? []).map((section) => ({
    id: section.id,
    number: section.number,
    title: section.title,
    writingHint: `围绕“${section.number} ${section.title}”编写本节正文，只写已确认项目事实，缺失信息标记待补充/需确认。`,
    placeholders: []
  }));
  const tables = (templateJson?.tables ?? []).map((table) => buildDynamicSchemeTemplateTable(table, baseSections));
  const figures = (templateJson?.figures ?? []).map((figure) => {
    const sectionNumber = figure.sectionNumber || inferSectionNumberFromTemplateAssetId(figure.id);
    const section = resolveTemplateTaskSectionByNumber(baseSections, sectionNumber);
    return {
      id: figure.id,
      section: section?.id,
      sectionNumber,
      recommendedLabel: normalizeFigureCaption(figure.caption) || figure.id,
      caption: figure.caption,
      purpose: figure.caption ? `按模板图位补充“${figure.caption}”相关图示。` : undefined
    } satisfies SchemeTemplateTaskFigure;
  });
  const relatedTables = groupTaskIdsBySection(tables);
  const relatedFigures = groupTaskIdsBySection(figures);
  const sections = baseSections.map((section) => ({
    ...section,
    relatedTables: relatedTables.get(section.id) ?? [],
    relatedFigures: relatedFigures.get(section.id) ?? []
  }));
  const fieldBlocks = (templateJson?.fieldBlocks ?? []).map((block) => ({
    id: block.id,
    text: block.text,
    section: block.section,
    sectionNumber: block.sectionNumber,
    placeholders: extractTemplateTaskPlaceholders(block.text),
    anchors: block.anchors
  }));
  const textBlocks = (templateJson?.textBlocks ?? []).map((block) => ({
    id: block.id,
    section: block.section,
    sectionNumber: block.sectionNumber,
    order: block.order,
    text: block.text,
    anchors: block.anchors
  }));

  return { sections, tables, figures, fieldBlocks, textBlocks };
}

function buildDynamicSchemeTemplateTable(
  table: NonNullable<WordTemplateJson["tables"]>[number],
  sections: SchemeTemplateTaskSection[]
): SchemeTemplateTaskTable {
  const sectionNumber = inferSectionNumberFromTemplateAssetId(table.id);
  const section = resolveTemplateTaskSectionByNumber(sections, sectionNumber);
  const rows = (table.rows ?? []).map((row) => ({
    index: row.index,
    cells: row.cells.map((cell) => ({
      rowIndex: row.index,
      cellIndex: cell.cellIndex,
      columnIndex: cell.columnIndex,
      text: cell.text ?? "",
      placeholders: extractTemplateTaskPlaceholders(cell.text)
    }))
  }));
  const header = inferDynamicSchemeTemplateTableHeader(rows);
  return {
    id: table.id,
    section: section?.id,
    sectionNumber,
    caption: table.caption,
    purpose: table.caption ? `按模板表格“${table.caption}”补充项目事实。` : undefined,
    writeStrategy: inferDynamicSchemeTemplateTableWriteStrategy(table.caption, rows),
    header,
    placeholders: extractTemplateTaskPlaceholders([table.caption, ...rows.flatMap((row) => row.cells.map((cell) => cell.text))].join("\n")),
    rows
  };
}

function inferDynamicSchemeTemplateTableHeader(rows: SchemeTemplateTaskRow[]): string[] {
  const firstRow = rows.find((row) => row.index === 0) ?? rows[0];
  return firstRow?.cells
    .slice()
    .sort((left, right) => left.columnIndex - right.columnIndex)
    .map((cell) => cell.text.trim())
    .filter(Boolean) ?? [];
}

function inferDynamicSchemeTemplateTableWriteStrategy(caption: string | undefined, rows: SchemeTemplateTaskRow[]): string | undefined {
  const dataRows = rows.filter((row) => row.index > 0);
  const placeholderCells = dataRows.flatMap((row) => row.cells).filter((cell) => isFillableTemplateCell(cell));
  const firstDataCellCount = dataRows[0]?.cells.length ?? 0;
  if (dataRows.length <= 1 && placeholderCells.length >= Math.max(2, Math.ceil(firstDataCellCount / 2))) {
    return "replace_table_when_data_complete";
  }
  if (rows.length <= 2 && /清单|列表|对象|措施|映射|汇总/.test(caption ?? "")) {
    return "replace_table_when_data_complete";
  }
  return undefined;
}

function groupTaskIdsBySection<T extends { id: string; section?: string }>(items: T[]): Map<string, string[]> {
  const grouped = new Map<string, string[]>();
  for (const item of items) {
    if (!item.section) continue;
    grouped.set(item.section, [...(grouped.get(item.section) ?? []), item.id]);
  }
  return grouped;
}

function resolveTemplateTaskSectionByNumber(
  sections: SchemeTemplateTaskSection[],
  sectionNumber: string | undefined
): SchemeTemplateTaskSection | undefined {
  let current = sectionNumber?.trim() ?? "";
  while (current) {
    const match = sections.find((section) => section.number === current);
    if (match) return match;
    const parent = current.replace(/\.\d+$/, "");
    if (parent === current) return undefined;
    current = parent;
  }
  return undefined;
}

function inferSectionNumberFromTemplateAssetId(id: string): string | undefined {
  return id.match(/^(?:table|fig)_\d+_(\d+(?:_\d+)*)$/i)?.[1]?.replace(/_/g, ".");
}

function extractTemplateTaskPlaceholders(text: string | undefined): string[] {
  if (!text) return [];
  return Array.from(new Set(Array.from(text.matchAll(/\{([^{}]{1,80})\}/g)).map((match) => match[1].trim()).filter(Boolean)));
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
      relatedFigures: readRecordStringArray(record, "relatedFigures"),
      anchors: readTemplateTaskSectionAnchors(record.anchors)
    });
  }
  return sections;
}

function readTemplateTaskSectionAnchors(value: unknown): SchemeTemplateTaskSection["anchors"] | undefined {
  if (!value || typeof value !== "object") return undefined;
  const record = value as Record<string, unknown>;
  const body = readTemplateTaskAnchor(record.body);
  return body ? { body } : undefined;
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

function buildDocumentSourceSummary(
  input: { userPrompt?: string; projectContext?: string; memory?: string },
  maxChars: number
): string {
  return compactText(
    [
      cleanDocumentContextText(input.userPrompt),
      cleanDocumentContextText(input.projectContext),
      cleanDocumentContextText(input.memory)
    ].filter(Boolean).join("\n\n"),
    maxChars
  );
}

function cleanDocumentContextText(value: string | undefined): string {
  const normalized = (value || "").replace(/\r\n/g, "\n").trim();
  if (!normalized) return "";
  const firstToolOutputIndex = normalized.search(/^##\s*工具\s+\S+/m);
  const withoutToolOutputs = firstToolOutputIndex >= 0 ? normalized.slice(0, firstToolOutputIndex) : normalized;
  return withoutToolOutputs.replace(/\n{3,}/g, "\n\n").trim();
}

function buildDocumentConfigSectionLike(section: SchemeTemplateTaskSection): DocumentConfigSectionLike {
  return {
    id: section.id,
    number: section.number,
    title: section.title,
    writingHint: section.writingHint,
    placeholders: section.placeholders,
    relatedTables: section.relatedTables,
    relatedFigures: section.relatedFigures
  };
}

function buildDocumentConfigSummary(config: DocumentConfig): string {
  return formatDocumentConfigSummary(config);
}

function formatDefaultDocumentTitle(sessionTitle: string | undefined, suffix: string): string {
  const title = sessionTitle?.trim() || suffix;
  return title.includes(suffix) || /方案|报告|文档|说明书|手册/.test(title) ? title : `${title}${suffix}`;
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

async function executePlanDocumentAssets(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const loaded = await loadDocumentTemplateAssetTaskJson(context);
  return executeTemplateAssetPlan(args, context, loaded.parsed, {
    toolName: "plan_document_assets",
    sourceLabel: loaded.sourceLabel,
    nextCallName: "next_plan_document_assets_call",
    writeToolName: "write_document_word",
    diagramsPlanName: "write_document_word_diagrams_plan"
  });
}

async function loadDocumentTemplateAssetTaskJson(
  context: AgentToolExecutionContext
): Promise<{ parsed: SchemeTemplateTaskJson; sourceLabel: string }> {
  const documentConfig = loadDocumentConfig(getDocumentConfigPath(context.outputDir));
  const wordTemplate = documentConfig?.wordTemplate;
  if (wordTemplate?.templateJsonPath) {
    const templateJsonPath = resolveOptionalWordTemplateJsonToolPath(
      wordTemplate.templateJsonPath,
      context,
      "plan_document_assets.template_json_path"
    );
    if (templateJsonPath) {
      const parsed = readSchemeTemplateTaskJsonIfPresent(templateJsonPath);
      if (hasTemplateAssetTasks(parsed)) {
        return {
          parsed: parsed!,
          sourceLabel: formatToolSourcePath(templateJsonPath, context)
        };
      }
    }
  }

  const templateDocxPath = resolveWordTemplateToolPath(
    wordTemplate?.templatePath,
    context,
    "plan_document_assets.template_path"
  );
  try {
    const parsedDocx = await parseWordTemplateAnchorsFromDocx(templateDocxPath);
    const parsed = buildSchemeTemplateTaskJsonFromWordTemplate(parsedDocx);
    if (hasTemplateAssetTasks(parsed)) {
      return {
        parsed,
        sourceLabel: formatToolSourcePath(templateDocxPath, context)
      };
    }
  } catch {
    // Fall through to the built-in template task JSON below.
  }

  const parsed = await loadSchemeTemplateTaskJson(context);
  return {
    parsed,
    sourceLabel: BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH
  };
}

function hasTemplateAssetTasks(parsed: SchemeTemplateTaskJson | undefined): boolean {
  if (!parsed) return false;
  return Boolean(
    readTemplateTaskSections(parsed.sections).length ||
      readTemplateTaskTables(parsed.tables).length ||
      readTemplateTaskFigures(parsed.figures).length
  );
}

async function executeTemplateAssetPlan(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext,
  parsed: SchemeTemplateTaskJson,
  options: {
    toolName: "plan_document_assets";
    sourceLabel: string;
    nextCallName: string;
    writeToolName: "write_document_word";
    diagramsPlanName: string;
  }
): Promise<AgentToolExecutionResult> {
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
      toolName: options.toolName,
      summary: "章节不在模板中",
      content: `${options.toolName} failed: unknown section_ids ${unknownSections.join("、")}; use existing sections[].id values from ${options.sourceLabel}`
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
      toolName: options.toolName,
      summary: "表格不在模板中",
      content: `${options.toolName} failed: unknown table_ids ${unknownTableIds.join("、")}; use existing tables[].id values from ${options.sourceLabel}`
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
    toolName: options.toolName,
    summary: `已规划表格 ${plannedTables.length} 个、图示 ${plannedFigures.length} 个`,
    content: formatSchemeAssetPlan({
      sections: selectedSections,
      tables: plannedTables,
      figures: plannedFigures,
      cellOffset,
      maxCells,
      figureOffset,
      maxFigures,
      sourceLabel: options.sourceLabel,
      toolName: options.toolName,
      nextCallName: options.nextCallName,
      writeToolName: options.writeToolName,
      diagramsPlanName: options.diagramsPlanName
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
  sourceLabel?: string;
  toolName?: "plan_document_assets";
  nextCallName?: string;
  writeToolName?: "write_document_word";
  diagramsPlanName?: string;
}): string {
  const toolName = input.toolName || "plan_document_assets";
  const nextCallName = input.nextCallName || "next_plan_document_assets_call";
  const writeToolName = input.writeToolName || "write_document_word";
  const diagramsPlanName = input.diagramsPlanName || "write_document_word_diagrams_plan";
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
    `${toolName} completed`,
    input.sourceLabel ? `模板来源：${input.sourceLabel}` : "",
    `章节：${input.sections.map((section) => `${section.id}(${section.number})`).join("、") || "全部关联章节"}`,
    `表格任务：${input.tables.length} 个（整表替换：${templateTables.length} 个，单元格填充：${templateCellTables.length} 个）；可填单元格：${allTableCells.length} 个；本批单元格：${tableCells.length} 个（offset=${input.cellOffset}, limit=${input.maxCells}）`,
    `图示任务：${input.figures.length} 个；本批图示：${figurePage.length} 个（offset=${input.figureOffset}, limit=${input.maxFigures}）`,
    `是否还有后续：${hasMoreCells || hasMoreFigures ? "是" : "否"}`,
    "",
    "使用规则：",
    `- dynamic 表格先根据 project_context/项目档案生成完整表内容，再通过 ${writeToolName}.template_tables 写入；建议优先传 markdown，首行为表头。`,
    `- fixed 骨架表先根据 project_context/项目档案把本批 template_cells_plan 的 value 建议改成具体值，再通过 ${writeToolName}.template_cells 写入。`,
    `- 如果 ${nextCallName} 不为空，必须继续调用并写入下一批，直到“是否还有后续：否”。`,
    "- template_cells 的 table_id、row_index、column_index/cell_index 必须原样保留；不要自行新增行列坐标。",
    `- 对 figures 先并行调用 image_generate；再在 ${writeToolName}.diagrams 中传 figure_id、label、kind、path，按不可见图片锚精确嵌入。`,
    "",
    `${nextCallName}:`,
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
    `${diagramsPlanName}:`,
    JSON.stringify(diagramRefs, null, 2)
  ].filter((line) => line !== "");
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

type DocumentSectionDraftStatus = "pending" | "drafted" | "edited" | "reviewed";

interface DocumentSectionManifestItem {
  id: string;
  order: number;
  number: string;
  title: string;
  headingLevel: number;
  bodyTag?: string;
  bodyAlias?: string;
  writingHint?: string;
  paragraphTasks: string[];
  placeholders: string[];
  relatedTables: string[];
  relatedFigures: string[];
  relatedTableSummaries: string[];
  relatedFigureSummaries: string[];
  sectionGroupId?: string;
  sectionGroupTitle?: string;
  writingRules: string[];
  evidenceRules: string[];
  requiredFacts: string[];
  draftPath: string;
  metaPath: string;
  tablesPath: string;
  figuresPath: string;
  evidenceReportPath: string;
  evidenceReviewPath: string;
  assetsDir: string;
  status: DocumentSectionDraftStatus;
  updatedAt?: string;
}

interface DocumentSectionManifest {
  version: 1;
  title: string;
  profile?: string;
  generatedAt: string;
  templateSource: string;
  configPath?: string;
  sections: DocumentSectionManifestItem[];
}

interface DocumentSectionMeta {
  sectionId?: string;
  status?: DocumentSectionDraftStatus;
  updatedAt?: string;
  changeSummary?: string;
}

async function executeListDocumentSections(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const { manifest } = await ensureDocumentSectionManifest(context);
  const query = readStringArg(args, "query");
  const status = normalizeDocumentSectionStatus(readStringArg(args, "status")) || "all";
  const includeAssets = readBooleanArg(args, "include_assets", true);
  const maxSections = Math.min(Math.max(Math.trunc(readNumberArg(args, "max_sections", 80)), 1), 200);
  const filtered = manifest.sections
    .filter((section) => matchesDocumentSectionListQuery(section, query))
    .filter((section) => status === "all" || section.status === status)
    .slice(0, maxSections);
  const omitted = manifest.sections.length - filtered.length;
  const manifestPath = getDocumentSectionManifestPath(context);

  return {
    toolName: "list_document_sections",
    summary: `已读取 ${filtered.length}/${manifest.sections.length} 个模板章节`,
    content: formatDocumentSectionList(manifest, filtered, {
      query,
      status,
      includeAssets,
      omitted,
      manifestPath,
      context
    }),
    artifactPath: manifestPath
  };
}

async function executeGetDocumentSection(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const sectionInput = readStringArg(args, "section");
  const includeDraft = readBooleanArg(args, "include_draft", true);
  const includeAssets = readBooleanArg(args, "include_assets", true);
  const loaded = await ensureDocumentSectionManifest(context);
  const matches = findDocumentSectionManifestMatches(sectionInput, loaded.manifest.sections);

  if (!matches.length) {
    return {
      toolName: "get_document_section",
      summary: "章节不在模板中",
      content: `get_document_section failed: unknown section ${sectionInput}; call list_document_sections and use sections[].id`
    };
  }
  if (matches.length > 1) {
    return {
      toolName: "get_document_section",
      summary: "章节匹配不唯一",
      content: [
        `get_document_section failed: ambiguous section ${sectionInput}; use exact sections[].id`,
        "",
        "matches:",
        ...matches.slice(0, 20).map((section) => `- ${section.id} | ${section.number} ${section.title}`)
      ].join("\n")
    };
  }

  const section = matches[0];
  const draftPath = join(context.outputDir, section.draftPath);
  const draft = includeDraft && existsSync(draftPath) ? readDocumentSectionDraftBody(draftPath) : undefined;
  const tableMap = new Map(loaded.tables.map((table) => [table.id, table]));
  const figureMap = new Map(loaded.figures.map((figure) => [figure.id, figure]));

  return {
    toolName: "get_document_section",
    summary: `${section.number} ${section.title}`,
    content: formatDocumentSectionDetail(section, {
      context,
      draft,
      includeAssets,
      tableMap,
      figureMap
    }),
    artifactPath: existsSync(draftPath) ? draftPath : getDocumentSectionManifestPath(context)
  };
}

interface DraftDocumentSectionResult {
  section: DocumentSectionManifestItem;
  status: "drafted" | "failed";
  content?: string;
  filePath?: string;
  error?: string;
}

async function executeDraftDocumentSections(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const projectContext = readStringArg(args, "project_context");
  const { manifest } = await ensureDocumentSectionManifest(context);
  const requestedSectionIds = readStringListArg(args, "section_ids").length
    ? readStringListArg(args, "section_ids")
    : readStringListArg(args, "sectionIds");
  const resolved = resolveDocumentSectionManifestItems(requestedSectionIds, manifest.sections);
  if (resolved.unknown.length) {
    return {
      toolName: "draft_document_sections",
      summary: "章节不在模板中",
      content: `draft_document_sections failed: unknown section_ids ${resolved.unknown.join("、")}; use sections[].id from list_document_sections`
    };
  }
  if (resolved.ambiguous.length) {
    return {
      toolName: "draft_document_sections",
      summary: "章节匹配不唯一",
      content: `draft_document_sections failed: ambiguous section_ids ${resolved.ambiguous.join("、")}; use exact sections[].id`
    };
  }

  const defaultSections = manifest.sections.filter((section) => section.status === "pending");
  const plannedSections = requestedSectionIds.length ? resolved.sections : defaultSections;
  if (!plannedSections.length) {
    return {
      toolName: "draft_document_sections",
      summary: "没有待起草章节",
      content: "draft_document_sections skipped: document-sections/manifest.json 中没有 pending 章节；如需重写已起草章节，请显式传 section_ids。"
    };
  }

  const requestedParallel = readNumberArg(args, "max_parallel", Number.NaN);
  const maxParallel = Number.isFinite(requestedParallel)
    ? clampDraftSectionParallelism(requestedParallel)
    : clampDraftSectionParallelism(context.settings.agent.draftSectionParallelism);
  const sections = plannedSections.slice(0, DRAFT_SECTION_PARALLELISM_MAX);
  const documentConfig = loadDocumentConfig(getDocumentConfigPath(context.outputDir));
  const results: DraftDocumentSectionResult[] = await mapWithConcurrency(sections, maxParallel, async (section): Promise<DraftDocumentSectionResult> => {
    try {
      const content = await draftDocumentSection(section, documentConfig, context, projectContext);
      return { section, status: "drafted", content } satisfies DraftDocumentSectionResult;
    } catch (error) {
      return {
        section,
        status: "failed",
        error: error instanceof Error ? error.message : String(error)
      } satisfies DraftDocumentSectionResult;
    }
  });
  const batchArtifactPath = await persistDraftDocumentSectionResults(results, manifest, context);
  await ensureDocumentSectionManifest(context);

  const drafted = results.filter((result) => result.status === "drafted").length;
  const failed = results.length - drafted;

  return {
    toolName: "draft_document_sections",
    summary: failed ? `已并行起草 ${drafted}/${results.length} 个章节，失败 ${failed} 个` : `已并行起草 ${drafted} 个章节`,
    content: formatDraftDocumentSectionResults(results, {
      maxParallel,
      truncated: plannedSections.length > sections.length,
      batchArtifactPath,
      context
    }),
    artifactPath: batchArtifactPath,
    schemeProgressUpdates: results.map((result) => ({
      section: result.section.id,
      anchorIds: result.section.bodyTag ? [result.section.bodyTag] : [result.section.id],
      status: result.status === "drafted" ? "drafted" : "failed",
      detail:
        result.status === "drafted"
          ? `${result.section.number} ${result.section.title} 章节草稿已生成`
          : `${result.section.number} ${result.section.title} 章节起草失败：${result.error || "未知错误"}`,
      artifactName: result.filePath ? basename(result.filePath) : undefined
    }))
  };
}

async function draftDocumentSection(
  section: DocumentSectionManifestItem,
  documentConfig: DocumentConfig | undefined,
  context: AgentToolExecutionContext,
  projectContext: string
): Promise<string> {
  if (context.signal?.aborted) {
    throw new Error("draft_document_sections aborted");
  }

  const apiKey = context.settings.openai.apiKeyConfigured ? process.env.OPENAI_API_KEY : undefined;
  if (!apiKey) {
    return buildFallbackDocumentSectionDraft(section, documentConfig, context, projectContext);
  }

  const client = new OpenAI({
    apiKey,
    baseURL: context.settings.openai.baseUrl,
    timeout: context.settings.openai.requestTimeoutMs
  });
  const response = await client.chat.completions.create(
    {
      model: context.settings.openai.chatModel,
      messages: buildDocumentSectionDraftMessages(section, documentConfig, context, projectContext),
      max_tokens: Math.min(Math.max(Math.floor(context.settings.openai.maxOutputTokens / 6), 900), 2600)
    },
    { signal: context.signal }
  );
  const content = response.choices[0]?.message?.content?.trim();
  if (!content) {
    throw new Error("模型未返回章节草稿");
  }
  return normalizeDraftContent(content);
}

function buildDocumentSectionDraftMessages(
  section: DocumentSectionManifestItem,
  documentConfig: DocumentConfig | undefined,
  context: AgentToolExecutionContext,
  projectContext: string
): ChatCompletionMessageParam[] {
  const contextText = compactText(
    [
      cleanDocumentContextText(context.userPrompt),
      cleanDocumentContextText(projectContext),
      cleanDocumentContextText(context.memory),
      documentConfig?.sourceSummary
    ].filter(Boolean).join("\n\n"),
    14000
  );
  return [
    {
      role: "system",
      content: [
        "你是文档章节正文起草器，只负责起草一个模板 section。",
        "输出要求：只输出该章节 Markdown 正文，不要输出代码块，不要解释你的思路，不要输出工具状态、section id、Word tag 或模板内部任务清单。",
        "可以输出自然段、必要列表和可审校的 Markdown 表格；但最终 Word 图表仍由 plan_document_assets/update_document_section_draft 的结构化内容和 write_document_word 统一写入。",
        "必须遵守 profile/globalRules、章节 writingRules、evidenceRules 和 requiredFacts；资料不足处写待补充/需确认。",
        "事实证据规则：只有用户原始需求、project_context、项目档案已确认事实、已读取附件/工具结果可以写成确定项目事实。",
        "严禁把模板提示、行文规则、标准要求、示例资料、行业常识或推断写成项目已经具备的现状。",
        "不得使用“已部署、已建设、已接入、已采用、已配置、已完成”等确定措辞，除非项目事实和上下文里明确给出。",
        "降低 AI 味：围绕本节对象写短而具体的句子，说明对象、边界、约束、缺口和可验证依据，避免空泛套话和重复背景。"
      ].join("\n")
    },
    {
      role: "user",
      content: [
        `文档标题：${documentConfig?.title || context.sessionTitle || "文档"}`,
        documentConfig ? `文档 profile：${documentConfig.profile}` : "",
        `章节：${section.number} ${section.title}`,
        section.writingHint ? `写作提示：${section.writingHint}` : "",
        section.sectionGroupTitle ? `所属章节组：${section.sectionGroupId || ""} ${section.sectionGroupTitle}` : "",
        "",
        "整体生成规划：",
        ...formatDocumentGenerationPlanningLines(documentConfig?.planning),
        "",
        "paragraph_tasks:",
        ...section.paragraphTasks.map((task, index) => `${index + 1}. ${task}`),
        "",
        "globalRules:",
        ...(documentConfig?.globalRules.length ? documentConfig.globalRules.map((rule) => `- ${rule}`) : ["- 无"]),
        "",
        "writingRules:",
        ...(section.writingRules.length ? section.writingRules.map((rule) => `- ${rule}`) : ["- 无"]),
        "",
        "evidenceRules:",
        ...(section.evidenceRules.length ? section.evidenceRules.map((rule) => `- ${rule}`) : ["- 无"]),
        "",
        "requiredFacts:",
        ...(section.requiredFacts.length ? section.requiredFacts.map((fact) => `- ${fact}`) : ["- 无"]),
        "",
        "关联表格：",
        ...(section.relatedTableSummaries.length ? section.relatedTableSummaries.map((item) => `- ${item}`) : ["- 无"]),
        "",
        "关联图示：",
        ...(section.relatedFigureSummaries.length ? section.relatedFigureSummaries.map((item) => `- ${item}`) : ["- 无"]),
        "",
        "已确认事实：",
        ...(documentConfig ? formatDocumentConfigFactLines(documentConfig) : ["- 无"]),
        "",
        "待补充/需确认：",
        ...(documentConfig ? formatDocumentConfigGapLines(documentConfig) : ["- 无"]),
        "",
        "项目上下文和输入材料摘要：",
        contextText || "暂无明确项目事实。"
      ]
        .filter(Boolean)
        .join("\n")
    }
  ];
}

function buildFallbackDocumentSectionDraft(
  section: DocumentSectionManifestItem,
  documentConfig: DocumentConfig | undefined,
  context: AgentToolExecutionContext,
  projectContext: string
): string {
  const contextText = buildDocumentSourceSummary(
    {
      userPrompt: context.userPrompt,
      projectContext,
      memory: context.memory
    },
    900
  );
  const factText = documentConfig?.facts.length
    ? `已确认事实包括：${documentConfig.facts.map((fact) => `${fact.key}为${fact.value}`).join("；")}。`
    : "当前尚未形成足够的已确认项目事实，本节不得把模板提示或常识写成项目现状。";
  const gapItems = uniqueTextItems([
    ...(documentConfig?.gaps ?? []),
    ...section.requiredFacts.map((fact) => `${fact}待补充/需确认`)
  ]);
  const paragraphs = section.paragraphTasks.length
    ? section.paragraphTasks.map((task, index) => {
        if (index === 0) {
          return `本节围绕“${section.title}”展开，重点是${task}。${factText}`;
        }
        if (index === section.paragraphTasks.length - 1) {
          return `围绕“${task}”，当前资料应继续核实系统边界、设备清单、产品型号和责任主体；资料不足处应在后续补充/确认。`;
        }
        return `围绕“${task}”，应使用已确认事实说明对象、位置、调用路径和安全效果；没有证据的现状、产品、算法、部署位置和责任主体一律标注待补充/需确认。`;
      })
    : [`本节围绕“${section.title}”形成项目化正文。${factText}`];

  return [
    ...paragraphs,
    gapItems.length ? `本节当前至少需要补充：${gapItems.slice(0, 8).join("；")}。` : "",
    section.relatedTableSummaries.length ? `本节后续需结合${section.relatedTableSummaries.join("、")}补充表格，正文不依赖表格坐标。` : "",
    section.relatedFigureSummaries.length ? `本节后续需结合${section.relatedFigureSummaries.join("、")}生成配图，正文只保留自然场景说明。` : "",
    contextText ? `已确认上下文摘要：${contextText}` : ""
  ]
    .filter(Boolean)
    .join("\n\n");
}

async function persistDraftDocumentSectionResults(
  results: DraftDocumentSectionResult[],
  manifest: DocumentSectionManifest,
  context: AgentToolExecutionContext
): Promise<string | undefined> {
  const draftedResults = results.filter(
    (result): result is DraftDocumentSectionResult & { content: string } => result.status === "drafted" && Boolean(result.content?.trim())
  );
  if (!draftedResults.length) return undefined;

  const updatedAt = new Date().toISOString();
  for (const result of draftedResults) {
    const draftPath = join(context.outputDir, result.section.draftPath);
    const metaPath = join(context.outputDir, result.section.metaPath);
    await writeUtf8File(draftPath, formatDocumentSectionDraftMarkdown(result.section, result.content));
    await writeUtf8File(
      metaPath,
      `${JSON.stringify(
        {
          sectionId: result.section.id,
          number: result.section.number,
          title: result.section.title,
          status: "drafted",
          updatedAt,
          changeSummary: "draft_document_sections 自动起草"
        },
        null,
        2
      )}\n`
    );
    result.filePath = draftPath;
  }

  const batchName = sanitizeFileName(`${manifest.title || context.sessionTitle || "文档"}-章节草稿批次.md`);
  const batchPath = createOutputPath(context, batchName);
  await writeUtf8File(batchPath, formatDraftDocumentSectionBatchMarkdown(draftedResults, manifest, context));
  return batchPath;
}

function formatDraftDocumentSectionBatchMarkdown(
  results: Array<DraftDocumentSectionResult & { content: string }>,
  manifest: DocumentSectionManifest,
  context: AgentToolExecutionContext
): string {
  const lines = [
    `# ${manifest.title}章节草稿批次`,
    "",
    `本文件汇总本批已生成的章节 Markdown。单章节文件位于 data/output/${DOCUMENT_SECTION_DRAFT_DIRNAME}，可逐节审校和返工。`,
    "",
    "章节文件：",
    ...results.map((result) => `- ${result.section.id} ${result.section.number} ${result.section.title}: ${formatOutputRelativePath(result.filePath || "", context)}`),
    "",
    "全部章节完成后，先调用 audit_document_sections 做章节级证据审查，再调用 assemble_document_sections 生成总草稿和终稿 Markdown。",
    "",
    "next_assemble_document_sections_call:",
    JSON.stringify({ name: `${manifest.title}-章节终稿.md` }, null, 2),
    ""
  ];

  for (const result of results) {
    lines.push(`## ${result.section.number} ${result.section.title}`, "", result.content.trim(), "");
  }

  return lines.join("\n").trim();
}

function formatDraftDocumentSectionResults(
  results: DraftDocumentSectionResult[],
  options: { maxParallel: number; truncated: boolean; batchArtifactPath?: string; context: AgentToolExecutionContext }
): string {
  const lines = [
    "draft_document_sections completed",
    `并行度：${options.maxParallel}`,
    options.truncated
      ? `本批超过 ${DRAFT_SECTION_PARALLELISM_MAX} 个章节，已只处理前 ${DRAFT_SECTION_PARALLELISM_MAX} 个；不传 section_ids 再次调用时会继续处理 pending 章节。`
      : "",
    `manifest：${formatOutputRelativePath(getDocumentSectionManifestPath(options.context), options.context)}`,
    options.batchArtifactPath ? `批次文件：${formatOutputRelativePath(options.batchArtifactPath, options.context)}` : "",
    `单章节 Markdown 已写入 ${DOCUMENT_SECTION_DRAFT_DIRNAME}/；全部章节完成后先调用 audit_document_sections，再调用 assemble_document_sections 生成终稿 Markdown。`
  ].filter(Boolean);

  for (const result of results) {
    lines.push("", `## ${result.section.id} ${result.section.number} ${result.section.title}`, `状态：${result.status === "drafted" ? "已起草" : "失败"}`);
    if (result.status === "drafted") {
      if (result.filePath) lines.push(`章节文件：${formatOutputRelativePath(result.filePath, options.context)}`);
      lines.push(result.content || "");
    } else {
      lines.push(`错误：${result.error || "未知错误"}`);
    }
  }

  return lines.join("\n");
}

async function executeUpdateDocumentSectionDraft(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const sectionInput = readStringArg(args, "section");
  const content = readStringArg(args, "content");
  if (!content) {
    return {
      toolName: "update_document_section_draft",
      summary: "缺少章节正文",
      content: "update_document_section_draft failed: missing content"
    };
  }

  const { manifest } = await ensureDocumentSectionManifest(context);
  const matches = findDocumentSectionManifestMatches(sectionInput, manifest.sections);
  if (!matches.length) {
    return {
      toolName: "update_document_section_draft",
      summary: "章节不在模板中",
      content: `update_document_section_draft failed: unknown section ${sectionInput}; call list_document_sections and use sections[].id`
    };
  }
  if (matches.length > 1) {
    return {
      toolName: "update_document_section_draft",
      summary: "章节匹配不唯一",
      content: [
        `update_document_section_draft failed: ambiguous section ${sectionInput}; use exact sections[].id`,
        "",
        ...matches.slice(0, 20).map((section) => `- ${section.id} | ${section.number} ${section.title}`)
      ].join("\n")
    };
  }

  const section = matches[0];
  const requestedStatus = normalizeDocumentSectionStatus(readStringArg(args, "status"));
  const status: DocumentSectionDraftStatus = requestedStatus && requestedStatus !== "all" ? requestedStatus : "edited";
  if (status === "pending") {
    throw new Error("update_document_section_draft.status 不能为 pending");
  }
  const changeSummary = readStringArg(args, "change_summary") || readStringArg(args, "changeSummary");
  const updatedAt = new Date().toISOString();
  const draftPath = join(context.outputDir, section.draftPath);
  const metaPath = join(context.outputDir, section.metaPath);
  const tablesPath = join(context.outputDir, section.tablesPath);
  const figuresPath = join(context.outputDir, section.figuresPath);
  const tablesJson = readStringArg(args, "tables_json") || readStringArg(args, "tablesJson");
  const figuresJson = readStringArg(args, "figures_json") || readStringArg(args, "figuresJson");

  await writeUtf8File(draftPath, formatDocumentSectionDraftMarkdown(section, content));
  await writeUtf8File(
    metaPath,
    `${JSON.stringify(
      {
        sectionId: section.id,
        number: section.number,
        title: section.title,
        status,
        updatedAt,
        changeSummary
      },
      null,
      2
    )}\n`
  );
  if (tablesJson) {
    await writeUtf8File(tablesPath, `${JSON.stringify(parseJsonStringArg(tablesJson, "tables_json"), null, 2)}\n`);
  }
  if (figuresJson) {
    await writeUtf8File(figuresPath, `${JSON.stringify(parseJsonStringArg(figuresJson, "figures_json"), null, 2)}\n`);
  }
  await ensureDocumentSectionManifest(context);

  return {
    toolName: "update_document_section_draft",
    summary: `已更新章节草稿：${section.number} ${section.title}`,
    content: [
      "update_document_section_draft completed",
      `章节：${section.id} | ${section.number} ${section.title}`,
      `状态：${status}`,
      `草稿：${formatOutputRelativePath(draftPath, context)}`,
      `元数据：${formatOutputRelativePath(metaPath, context)}`,
      tablesJson ? `表格原始内容：${formatOutputRelativePath(tablesPath, context)}` : "",
      figuresJson ? `图示原始内容：${formatOutputRelativePath(figuresPath, context)}` : "",
      changeSummary ? `调整说明：${changeSummary}` : "",
      "",
      "next_assemble_document_sections_call:",
      JSON.stringify({ section_ids: [section.id] }, null, 2)
    ].filter(Boolean).join("\n"),
    artifactPath: draftPath,
    schemeProgressUpdates: [
      {
        section: section.id,
        anchorIds: section.bodyTag ? [section.bodyTag] : [section.id],
        status: status === "reviewed" ? "completed" : "drafted",
        detail: `${section.number} ${section.title} 草稿已更新`,
        artifactName: basename(draftPath)
      }
    ]
  };
}

interface DocumentSectionEvidenceResult {
  section: DocumentSectionManifestItem;
  draftPath: string;
  reportPath?: string;
  reviewPath?: string;
  findings: DocumentEvidenceFinding[];
  truncated: boolean;
  missingDraft: boolean;
  rewriteCount?: number;
  skippedCount?: number;
}

async function executeAuditDocumentSections(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const { manifest } = await ensureDocumentSectionManifest(context);
  const resolved = resolveDocumentSectionQualityTargets(args, manifest, context, "audit_document_sections");
  if (resolved.error) {
    return {
      toolName: "audit_document_sections",
      summary: resolved.summary,
      content: resolved.error
    };
  }
  if (!resolved.sections.length) {
    return {
      toolName: "audit_document_sections",
      summary: "没有可审查章节",
      content: "audit_document_sections skipped: document-sections 中没有已有 draft.md 的章节；请先调用 draft_document_sections 或 update_document_section_draft。"
    };
  }

  const documentConfig = loadDocumentConfig(getDocumentConfigPath(context.outputDir));
  const maxFindings = Math.min(Math.max(Math.trunc(readNumberArg(args, "max_findings", 40)), 1), 120);
  const includeCleanSections = readBooleanArg(args, "include_clean_sections", true);
  const results: DocumentSectionEvidenceResult[] = [];

  for (const section of resolved.sections) {
    const draftPath = join(context.outputDir, section.draftPath);
    if (!existsSync(draftPath)) {
      results.push({ section, draftPath, findings: [], truncated: false, missingDraft: true });
      continue;
    }
    const markdown = readFileSync(draftPath, "utf-8");
    const allFindings = auditMarkdownEvidence(markdown, documentConfig);
    const findings = allFindings.slice(0, maxFindings);
    const reportPath = join(context.outputDir, section.evidenceReportPath);
    const reviewPath = join(context.outputDir, section.evidenceReviewPath);
    await writeDocumentSectionEvidenceArtifacts({
      section,
      draftPath,
      reportPath,
      reviewPath,
      findings,
      truncated: allFindings.length > findings.length,
      documentConfig,
      context
    });
    if (!findings.length) {
      await writeDocumentSectionMetaUpdate(section, context, "reviewed", "audit_document_sections 未发现明显证据风险");
    }
    results.push({
      section,
      draftPath,
      reportPath,
      reviewPath,
      findings,
      truncated: allFindings.length > findings.length,
      missingDraft: false
    });
  }
  await ensureDocumentSectionManifest(context);

  const batchReportPath = createOutputPath(context, sanitizeFileName(`${manifest.title}-章节证据审查汇总.md`));
  await writeUtf8File(
    batchReportPath,
    formatDocumentSectionEvidenceBatchReport({
      title: manifest.title,
      results,
      includeCleanSections,
      context,
      mode: "audit"
    })
  );
  const findingCount = results.reduce((sum, result) => sum + result.findings.length, 0);
  const highRiskCount = results.reduce((sum, result) => sum + result.findings.filter((finding) => finding.level === "high").length, 0);
  const riskySectionIds = results.filter((result) => result.findings.length).map((result) => result.section.id);

  return {
    toolName: "audit_document_sections",
    summary: findingCount ? `发现 ${findingCount} 条章节证据风险，其中高风险 ${highRiskCount} 条` : "章节草稿未发现明显证据风险",
    content: [
      "audit_document_sections completed",
      `章节数：${results.length}`,
      `风险项：${findingCount}`,
      `高风险：${highRiskCount}`,
      `汇总报告：${formatOutputRelativePath(batchReportPath, context)}`,
      "",
      "章节报告：",
      ...results
        .filter((result) => includeCleanSections || result.findings.length || result.missingDraft)
        .map((result) =>
          `- ${result.section.id} | ${result.section.number} ${result.section.title} | 风险 ${result.findings.length}${
            result.reportPath ? ` | ${formatOutputRelativePath(result.reportPath, context)}` : result.missingDraft ? " | 草稿缺失" : ""
          }`
        ),
      riskySectionIds.length ? "" : "",
      riskySectionIds.length ? "next_revise_document_sections_evidence_call:" : "",
      riskySectionIds.length ? JSON.stringify({ section_ids: riskySectionIds }, null, 2) : "",
      "",
      "next_assemble_document_sections_call:",
      JSON.stringify({}, null, 2)
    ].filter((line) => line !== "").join("\n"),
    artifactPath: batchReportPath,
    schemeProgressUpdates: results.map((result) => ({
      section: result.section.id,
      anchorIds: result.section.bodyTag ? [result.section.bodyTag] : [result.section.id],
      status: result.missingDraft ? "failed" : result.findings.length ? "drafted" : "completed",
      detail: result.missingDraft
        ? `${result.section.number} ${result.section.title} 草稿缺失`
        : result.findings.length
          ? `${result.section.number} ${result.section.title} 存在 ${result.findings.length} 条证据风险`
          : `${result.section.number} ${result.section.title} 证据审查通过`,
      artifactName: result.reportPath ? basename(result.reportPath) : undefined
    }))
  };
}

async function executeReviseDocumentSectionsEvidence(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const { manifest } = await ensureDocumentSectionManifest(context);
  const resolved = resolveDocumentSectionQualityTargets(args, manifest, context, "revise_document_sections_evidence");
  if (resolved.error) {
    return {
      toolName: "revise_document_sections_evidence",
      summary: resolved.summary,
      content: resolved.error
    };
  }
  if (!resolved.sections.length) {
    return {
      toolName: "revise_document_sections_evidence",
      summary: "没有可修订章节",
      content: "revise_document_sections_evidence skipped: document-sections 中没有已有 draft.md 的章节；请先调用 draft_document_sections 或 update_document_section_draft。"
    };
  }

  const documentConfig = loadDocumentConfig(getDocumentConfigPath(context.outputDir));
  const maxRewrites = Math.min(Math.max(Math.trunc(readNumberArg(args, "max_rewrites", 40)), 1), 120);
  const results: DocumentSectionEvidenceResult[] = [];

  for (const section of resolved.sections) {
    const draftPath = join(context.outputDir, section.draftPath);
    if (!existsSync(draftPath)) {
      results.push({ section, draftPath, findings: [], truncated: false, missingDraft: true, rewriteCount: 0, skippedCount: 0 });
      continue;
    }
    const markdown = readFileSync(draftPath, "utf-8");
    const findings = auditMarkdownEvidence(markdown, documentConfig);
    const revision = reviseMarkdownEvidenceLines(markdown, findings, maxRewrites);
    if (revision.rewriteCount) {
      await writeUtf8File(draftPath, revision.markdown);
      await writeDocumentSectionMetaUpdate(section, context, "edited", "revise_document_sections_evidence 自动降级无证据确定表述");
    }
    const revisedMarkdown = existsSync(draftPath) ? readFileSync(draftPath, "utf-8") : markdown;
    const remainingFindings = auditMarkdownEvidence(revisedMarkdown, documentConfig).slice(0, maxRewrites);
    const reportPath = join(context.outputDir, section.evidenceReportPath);
    const reviewPath = join(context.outputDir, section.evidenceReviewPath);
    await writeDocumentSectionEvidenceArtifacts({
      section,
      draftPath,
      reportPath,
      reviewPath,
      findings: remainingFindings,
      truncated: false,
      documentConfig,
      context,
      rewriteCount: revision.rewriteCount,
      skippedCount: revision.skippedCount
    });
    results.push({
      section,
      draftPath,
      reportPath,
      reviewPath,
      findings: remainingFindings,
      truncated: false,
      missingDraft: false,
      rewriteCount: revision.rewriteCount,
      skippedCount: revision.skippedCount
    });
  }
  await ensureDocumentSectionManifest(context);

  const batchReportPath = createOutputPath(context, sanitizeFileName(`${manifest.title}-章节证据修订汇总.md`));
  await writeUtf8File(
    batchReportPath,
    formatDocumentSectionEvidenceBatchReport({
      title: manifest.title,
      results,
      includeCleanSections: true,
      context,
      mode: "revise"
    })
  );
  const rewriteCount = results.reduce((sum, result) => sum + (result.rewriteCount ?? 0), 0);
  const skippedCount = results.reduce((sum, result) => sum + (result.skippedCount ?? 0), 0);
  const remainingRiskCount = results.reduce((sum, result) => sum + result.findings.length, 0);

  return {
    toolName: "revise_document_sections_evidence",
    summary: rewriteCount ? `已修订 ${rewriteCount} 行章节证据风险` : "没有可自动修订的章节正文行",
    content: [
      "revise_document_sections_evidence completed",
      `章节数：${results.length}`,
      `已修订行：${rewriteCount}`,
      skippedCount ? `跳过行：${skippedCount}（表格或全局项需人工确认）` : "",
      `剩余风险项：${remainingRiskCount}`,
      `汇总报告：${formatOutputRelativePath(batchReportPath, context)}`,
      "",
      "章节结果：",
      ...results.map((result) =>
        `- ${result.section.id} | ${result.section.number} ${result.section.title} | 修订 ${result.rewriteCount ?? 0} | 剩余风险 ${result.findings.length}${
          result.reportPath ? ` | ${formatOutputRelativePath(result.reportPath, context)}` : result.missingDraft ? " | 草稿缺失" : ""
        }`
      ),
      "",
      "next_audit_document_sections_call:",
      JSON.stringify({ section_ids: results.map((result) => result.section.id) }, null, 2),
      "",
      "next_assemble_document_sections_call:",
      JSON.stringify({}, null, 2)
    ].filter(Boolean).join("\n"),
    artifactPath: batchReportPath,
    schemeProgressUpdates: results.map((result) => ({
      section: result.section.id,
      anchorIds: result.section.bodyTag ? [result.section.bodyTag] : [result.section.id],
      status: result.missingDraft ? "failed" : result.findings.length ? "drafted" : "completed",
      detail: result.missingDraft
        ? `${result.section.number} ${result.section.title} 草稿缺失`
        : `${result.section.number} ${result.section.title} 已修订 ${result.rewriteCount ?? 0} 行，剩余风险 ${result.findings.length} 条`,
      artifactName: result.reportPath ? basename(result.reportPath) : undefined
    }))
  };
}

function resolveDocumentSectionQualityTargets(
  args: Record<string, unknown>,
  manifest: DocumentSectionManifest,
  context: AgentToolExecutionContext,
  toolName: string
): { sections: DocumentSectionManifestItem[]; summary: string; error?: string } {
  const requestedSectionIds = readStringListArg(args, "section_ids").length
    ? readStringListArg(args, "section_ids")
    : readStringListArg(args, "sectionIds");
  const resolved = resolveDocumentSectionManifestItems(requestedSectionIds, manifest.sections);
  if (resolved.unknown.length) {
    return {
      sections: [],
      summary: "章节不在模板中",
      error: `${toolName} failed: unknown section_ids ${resolved.unknown.join("、")}; use sections[].id from list_document_sections`
    };
  }
  if (resolved.ambiguous.length) {
    return {
      sections: [],
      summary: "章节匹配不唯一",
      error: `${toolName} failed: ambiguous section_ids ${resolved.ambiguous.join("、")}; use exact sections[].id`
    };
  }
  const sections = requestedSectionIds.length
    ? resolved.sections
    : manifest.sections.filter((section) => existsSync(join(context.outputDir, section.draftPath)));
  return { sections, summary: "" };
}

async function writeDocumentSectionEvidenceArtifacts(input: {
  section: DocumentSectionManifestItem;
  draftPath: string;
  reportPath: string;
  reviewPath: string;
  findings: DocumentEvidenceFinding[];
  truncated: boolean;
  documentConfig?: DocumentConfig;
  context: AgentToolExecutionContext;
  rewriteCount?: number;
  skippedCount?: number;
}): Promise<void> {
  await writeUtf8File(
    input.reportPath,
    formatDocumentEvidenceAuditReport({
      markdownPath: input.draftPath,
      documentConfig: input.documentConfig,
      findings: input.findings,
      truncated: input.truncated,
      context: input.context
    })
  );
  await writeUtf8File(
    input.reviewPath,
    `${JSON.stringify(
      {
        sectionId: input.section.id,
        number: input.section.number,
        title: input.section.title,
        reviewedAt: new Date().toISOString(),
        draftPath: formatOutputRelativePath(input.draftPath, input.context),
        findingCount: input.findings.length,
        highRiskCount: input.findings.filter((finding) => finding.level === "high").length,
        truncated: input.truncated,
        rewriteCount: input.rewriteCount ?? 0,
        skippedCount: input.skippedCount ?? 0,
        findings: input.findings
      },
      null,
      2
    )}\n`
  );
}

async function writeDocumentSectionMetaUpdate(
  section: DocumentSectionManifestItem,
  context: AgentToolExecutionContext,
  status: DocumentSectionDraftStatus,
  changeSummary: string
): Promise<void> {
  const metaPath = join(context.outputDir, section.metaPath);
  const existing = existsSync(metaPath) ? readJsonObjectIfPresent(metaPath) : undefined;
  await writeUtf8File(
    metaPath,
    `${JSON.stringify(
      {
        ...(existing ?? {}),
        sectionId: section.id,
        number: section.number,
        title: section.title,
        status,
        updatedAt: new Date().toISOString(),
        changeSummary
      },
      null,
      2
    )}\n`
  );
}

function readJsonObjectIfPresent(filePath: string): Record<string, unknown> | undefined {
  try {
    const parsed = JSON.parse(readFileSync(filePath, "utf-8")) as unknown;
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? (parsed as Record<string, unknown>) : undefined;
  } catch {
    return undefined;
  }
}

function formatDocumentSectionEvidenceBatchReport(input: {
  title: string;
  results: DocumentSectionEvidenceResult[];
  includeCleanSections: boolean;
  context: AgentToolExecutionContext;
  mode: "audit" | "revise";
}): string {
  const visibleResults = input.includeCleanSections
    ? input.results
    : input.results.filter((result) => result.findings.length || result.missingDraft);
  const totalFindings = input.results.reduce((sum, result) => sum + result.findings.length, 0);
  const totalRewrites = input.results.reduce((sum, result) => sum + (result.rewriteCount ?? 0), 0);
  const lines = [
    `# ${input.title}章节证据${input.mode === "audit" ? "审查" : "修订"}汇总`,
    "",
    `章节数：${input.results.length}`,
    `风险项：${totalFindings}`,
    input.mode === "revise" ? `已修订行：${totalRewrites}` : "",
    "",
    "| 章节 | 状态 | 风险项 | 高风险 | 修订行 | 报告 |",
    "| --- | --- | --- | --- | --- | --- |"
  ].filter(Boolean);

  for (const result of visibleResults) {
    const highRiskCount = result.findings.filter((finding) => finding.level === "high").length;
    const status = result.missingDraft ? "草稿缺失" : result.findings.length ? "存在风险" : "通过";
    lines.push(
      `| ${escapeMarkdownTableCell(`${result.section.number} ${result.section.title}`)} | ${status} | ${result.findings.length} | ${highRiskCount} | ${
        result.rewriteCount ?? 0
      } | ${result.reportPath ? escapeMarkdownTableCell(formatOutputRelativePath(result.reportPath, input.context)) : ""} |`
    );
  }

  const riskyResults = input.results.filter((result) => result.findings.length);
  if (riskyResults.length) {
    lines.push("", "## 风险章节", "");
    for (const result of riskyResults) {
      lines.push(`### ${result.section.number} ${result.section.title}`, "");
      for (const finding of result.findings.slice(0, 20)) {
        lines.push(`- ${finding.level === "high" ? "高" : "中"}风险，第 ${finding.lineNumber || "全局"} 行：${finding.reason}。${finding.text}`);
      }
      if (result.findings.length > 20) {
        lines.push(`- 其余 ${result.findings.length - 20} 条见章节报告。`);
      }
      lines.push("");
    }
  }

  return lines.join("\n").trim();
}

interface PolishDocumentSectionResult {
  section: DocumentSectionManifestItem;
  status: "polished" | "skipped" | "failed";
  filePath?: string;
  content?: string;
  error?: string;
}

async function executePolishDocumentSections(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const { manifest } = await ensureDocumentSectionManifest(context);
  const resolved = resolveDocumentSectionQualityTargets(args, manifest, context, "polish_document_sections");
  if (resolved.error) {
    return {
      toolName: "polish_document_sections",
      summary: resolved.summary,
      content: resolved.error
    };
  }
  if (!resolved.sections.length) {
    return {
      toolName: "polish_document_sections",
      summary: "没有可润色章节",
      content: "polish_document_sections skipped: document-sections 中没有已有 draft.md 的章节；请先调用 draft_document_sections 或 update_document_section_draft。"
    };
  }

  const requestedParallel = readNumberArg(args, "max_parallel", Number.NaN);
  const maxParallel = Number.isFinite(requestedParallel)
    ? clampDraftSectionParallelism(requestedParallel)
    : clampDraftSectionParallelism(context.settings.agent.draftSectionParallelism);
  const sections = resolved.sections.slice(0, DRAFT_SECTION_PARALLELISM_MAX);
  const documentConfig = loadDocumentConfig(getDocumentConfigPath(context.outputDir));
  const projectContext = readStringArg(args, "project_context");
  const styleRules = readStringListArg(args, "style_rules").length
    ? readStringListArg(args, "style_rules")
    : readStringListArg(args, "styleRules");
  const results: PolishDocumentSectionResult[] = await mapWithConcurrency(
    sections,
    maxParallel,
    async (section): Promise<PolishDocumentSectionResult> => {
      try {
        const draftPath = join(context.outputDir, section.draftPath);
        if (!existsSync(draftPath)) {
          return { section, status: "skipped", error: "draft.md 不存在" };
        }
        const currentBody = readDocumentSectionDraftBody(draftPath);
        if (!currentBody.trim()) {
          return { section, status: "skipped", error: "draft.md 正文为空" };
        }
        const content = await polishDocumentSection(section, currentBody, {
          documentConfig,
          context,
          projectContext,
          styleRules
        });
        await writeUtf8File(draftPath, formatDocumentSectionDraftMarkdown(section, content));
        await writeDocumentSectionMetaUpdate(section, context, "edited", "polish_document_sections 优化章节语言表达");
        return { section, status: "polished", filePath: draftPath, content };
      } catch (error) {
        return {
          section,
          status: "failed",
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
  );
  await ensureDocumentSectionManifest(context);

  const batchPath = createOutputPath(context, sanitizeFileName(`${manifest.title}-章节语言优化汇总.md`));
  await writeUtf8File(batchPath, formatPolishDocumentSectionBatchMarkdown(results, context));
  const polishedCount = results.filter((result) => result.status === "polished").length;
  const failedCount = results.filter((result) => result.status === "failed").length;
  const skippedCount = results.filter((result) => result.status === "skipped").length;

  return {
    toolName: "polish_document_sections",
    summary: failedCount
      ? `已润色 ${polishedCount}/${results.length} 个章节，失败 ${failedCount} 个`
      : `已润色 ${polishedCount} 个章节${skippedCount ? `，跳过 ${skippedCount} 个` : ""}`,
    content: [
      "polish_document_sections completed",
      `并行度：${maxParallel}`,
      `章节数：${results.length}`,
      `已润色：${polishedCount}`,
      skippedCount ? `已跳过：${skippedCount}` : "",
      failedCount ? `失败：${failedCount}` : "",
      `汇总：${formatOutputRelativePath(batchPath, context)}`,
      "",
      "章节结果：",
      ...results.map((result) =>
        `- ${result.section.id} | ${result.section.number} ${result.section.title} | ${result.status}${
          result.filePath ? ` | ${formatOutputRelativePath(result.filePath, context)}` : result.error ? ` | ${result.error}` : ""
        }`
      ),
      "",
      "next_audit_document_sections_call:",
      JSON.stringify({ section_ids: results.filter((result) => result.status === "polished").map((result) => result.section.id) }, null, 2),
      "",
      "next_assemble_document_sections_call:",
      JSON.stringify({}, null, 2)
    ].filter(Boolean).join("\n"),
    artifactPath: batchPath,
    schemeProgressUpdates: results.map((result) => ({
      section: result.section.id,
      anchorIds: result.section.bodyTag ? [result.section.bodyTag] : [result.section.id],
      status: result.status === "failed" ? "failed" : result.status === "skipped" ? "skipped" : "drafted",
      detail:
        result.status === "polished"
          ? `${result.section.number} ${result.section.title} 已完成语言优化`
          : result.status === "skipped"
            ? `${result.section.number} ${result.section.title} 已跳过语言优化：${result.error || ""}`
            : `${result.section.number} ${result.section.title} 语言优化失败：${result.error || "未知错误"}`,
      artifactName: result.filePath ? basename(result.filePath) : undefined
    }))
  };
}

async function polishDocumentSection(
  section: DocumentSectionManifestItem,
  body: string,
  input: {
    documentConfig?: DocumentConfig;
    context: AgentToolExecutionContext;
    projectContext: string;
    styleRules: string[];
  }
): Promise<string> {
  if (input.context.signal?.aborted) {
    throw new Error("polish_document_sections aborted");
  }

  const apiKey = input.context.settings.openai.apiKeyConfigured ? process.env.OPENAI_API_KEY : undefined;
  if (!apiKey) {
    return fallbackPolishDocumentSectionContent(body);
  }

  const client = new OpenAI({
    apiKey,
    baseURL: input.context.settings.openai.baseUrl,
    timeout: input.context.settings.openai.requestTimeoutMs
  });
  const response = await client.chat.completions.create(
    {
      model: input.context.settings.openai.chatModel,
      messages: buildPolishDocumentSectionMessages(section, body, input),
      max_tokens: Math.min(Math.max(Math.floor(input.context.settings.openai.maxOutputTokens / 6), 900), 3200)
    },
    { signal: input.context.signal }
  );
  const content = response.choices[0]?.message?.content?.trim();
  if (!content) {
    throw new Error("模型未返回章节润色结果");
  }
  return normalizeDraftContent(content);
}

function buildPolishDocumentSectionMessages(
  section: DocumentSectionManifestItem,
  body: string,
  input: {
    documentConfig?: DocumentConfig;
    context: AgentToolExecutionContext;
    projectContext: string;
    styleRules: string[];
  }
): ChatCompletionMessageParam[] {
  const contextText = compactText(
    [
      cleanDocumentContextText(input.context.userPrompt),
      cleanDocumentContextText(input.projectContext),
      cleanDocumentContextText(input.context.memory),
      input.documentConfig?.sourceSummary
    ].filter(Boolean).join("\n\n"),
    10000
  );
  return [
    {
      role: "system",
      content: [
        "你是专业文档章节润色编辑，只优化表达质量，不新增项目事实。",
        "输出要求：只输出润色后的本章节 Markdown 正文，不要代码块，不要解释，不要输出章节标题、section id、Word tag 或内部规则。",
        "必须保持原文中的事实边界：待补充/需确认/未明确/暂无 等缺口标记必须保留，不得改成确定事实。",
        "不得新增建设单位、系统名称、产品型号、算法、部署位置、网络区域、接口、数量、责任主体或完成状态。",
        "可以做的事：压缩套话、统一术语、调整句序、增强专业书面语、让段落更连贯、修正明显病句。",
        "不能做的事：补写缺失事实、扩展未确认方案、把建议/要求改写成项目现状。"
      ].join("\n")
    },
    {
      role: "user",
      content: [
        `章节：${section.number} ${section.title}`,
        section.sectionGroupTitle ? `所属章节组：${section.sectionGroupTitle}` : "",
        "",
        "语言优化规则：",
        ...(input.styleRules.length ? input.styleRules.map((rule) => `- ${rule}`) : ["- 使用正式、克制、专业的方案文档表达。", "- 减少空泛套话和重复政策背景。", "- 保留所有待补充/需确认标记。"]),
        "",
        "已确认事实：",
        ...(input.documentConfig ? formatDocumentConfigFactLines(input.documentConfig) : ["- 无"]),
        "",
        "待补充/需确认：",
        ...(input.documentConfig ? formatDocumentConfigGapLines(input.documentConfig) : ["- 无"]),
        "",
        "项目上下文摘要：",
        contextText || "暂无明确项目事实。",
        "",
        "当前章节正文：",
        body
      ].filter(Boolean).join("\n")
    }
  ];
}

function fallbackPolishDocumentSectionContent(body: string): string {
  return normalizeDraftContent(
    body
      .replace(/\r\n/g, "\n")
      .replace(/本节围绕“([^”]+)”展开，重点是/g, "本节聚焦“$1”，重点说明")
      .replace(/应继续核实/g, "需继续核实")
      .replace(/一律标注/g, "均标注")
      .replace(/后续补充\/确认/g, "后续补充或确认")
      .replace(/避免空泛套话/g, "减少空泛表述")
      .replace(/\n{3,}/g, "\n\n")
      .trim()
  );
}

function formatPolishDocumentSectionBatchMarkdown(
  results: PolishDocumentSectionResult[],
  context: AgentToolExecutionContext
): string {
  const lines = [
    "# 章节语言优化汇总",
    "",
    "| 章节 | 状态 | 草稿 | 说明 |",
    "| --- | --- | --- | --- |"
  ];
  for (const result of results) {
    lines.push(
      `| ${escapeMarkdownTableCell(`${result.section.number} ${result.section.title}`)} | ${result.status} | ${
        result.filePath ? escapeMarkdownTableCell(formatOutputRelativePath(result.filePath, context)) : ""
      } | ${escapeMarkdownTableCell(result.error || "")} |`
    );
  }
  return lines.join("\n");
}

async function executeAssembleDocumentSections(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const { manifest } = await ensureDocumentSectionManifest(context);
  const requestedSectionIds = readStringListArg(args, "section_ids").length
    ? readStringListArg(args, "section_ids")
    : readStringListArg(args, "sectionIds");
  const resolved = resolveDocumentSectionManifestItems(requestedSectionIds, manifest.sections);
  if (resolved.unknown.length) {
    return {
      toolName: "assemble_document_sections",
      summary: "章节不在模板中",
      content: `assemble_document_sections failed: unknown section_ids ${resolved.unknown.join("、")}; use sections[].id from list_document_sections`
    };
  }
  if (resolved.ambiguous.length) {
    return {
      toolName: "assemble_document_sections",
      summary: "章节匹配不唯一",
      content: `assemble_document_sections failed: ambiguous section_ids ${resolved.ambiguous.join("、")}; use exact sections[].id`
    };
  }

  const sections = resolved.sections.length ? resolved.sections : manifest.sections;
  const includeDraft = args.include_draft !== false;
  const documentConfig = loadDocumentConfig(getDocumentConfigPath(context.outputDir));
  const title = documentConfig?.title || manifest.title || formatDefaultDocumentTitle(context.sessionTitle, "文档");
  const rules = [...DEFAULT_DOCUMENT_CONTENT_RULES, ...(documentConfig?.globalRules ?? []), ...readStringListArg(args, "document_rules")];
  const safeFinalName = sanitizeFileName(readStringArg(args, "name") || `${title}-章节终稿.md`);
  const finalName = safeFinalName.toLowerCase().endsWith(".md") ? safeFinalName : `${safeFinalName}.md`;
  const outputStamp = createToolArtifactStamp();
  const loadedSections = sections.map((section) => loadDocumentSectionForAssembly(section, context));
  const draftMarkdown = formatMergedDocumentSectionMarkdown(loadedSections, { title: `${title}章节总草稿` });
  const finalSourceMarkdown = formatMergedDocumentSectionMarkdown(loadedSections, { title });
  const finalMarkdown = finalizeDocumentMarkdown(finalSourceMarkdown, rules);
  const draftPath = createOutputPath(context, sanitizeFileName(`${title}-章节总草稿.md`), outputStamp);
  const finalPath = createOutputPath(context, finalName, outputStamp);

  if (includeDraft) {
    await writeUtf8File(draftPath, draftMarkdown);
  }
  await writeUtf8File(finalPath, finalMarkdown);
  const missingSections = loadedSections.filter((section) => !section.exists);

  return {
    toolName: "assemble_document_sections",
    summary: `已合并 ${sections.length} 个章节 Markdown`,
    content: [
      "assemble_document_sections completed",
      `章节数：${sections.length}`,
      `缺失草稿章节：${missingSections.length}`,
      includeDraft ? `总草稿：${formatOutputRelativePath(draftPath, context)}` : "总草稿：未输出",
      `终稿：${formatOutputRelativePath(finalPath, context)}`,
      "",
      "应用文档规则：",
      ...rules.map((rule) => `- ${rule}`),
      "",
      "章节来源：",
      ...loadedSections.map((section) =>
        `- ${section.item.id} | ${section.item.number} ${section.item.title}: ${
          section.exists ? formatOutputRelativePath(join(context.outputDir, section.item.draftPath), context) : "缺失，已写入待补充/需确认占位"
        }`
      ),
      "",
      "next_audit_call:",
      JSON.stringify({ markdown_path: formatOutputRelativePath(finalPath, context) }, null, 2),
      "",
      "next_plan_document_assets_call:",
      JSON.stringify({ section_ids: sections.map((section) => section.id) }, null, 2),
      "",
      "next_word_call:",
      JSON.stringify({ markdown_path: formatOutputRelativePath(finalPath, context) }, null, 2)
    ].join("\n"),
    artifactPath: finalPath
  };
}

async function ensureDocumentSectionManifest(context: AgentToolExecutionContext): Promise<{
  manifest: DocumentSectionManifest;
  parsed: SchemeTemplateTaskJson;
  tables: SchemeTemplateTaskTable[];
  figures: SchemeTemplateTaskFigure[];
}> {
  const loaded = await loadDocumentTemplateAssetTaskJson(context);
  const sections = readTemplateTaskSections(loaded.parsed.sections);
  const tables = readTemplateTaskTables(loaded.parsed.tables);
  const figures = readTemplateTaskFigures(loaded.parsed.figures);
  const tableMap = new Map(tables.map((table) => [table.id, table]));
  const figureMap = new Map(figures.map((figure) => [figure.id, figure]));
  const documentConfig = loadDocumentConfig(getDocumentConfigPath(context.outputDir));
  const previousManifest = readDocumentSectionManifestIfPresent(getDocumentSectionManifestPath(context));
  const manifest: DocumentSectionManifest = {
    version: 1,
    title: documentConfig?.title || formatDefaultDocumentTitle(context.sessionTitle, "文档"),
    ...(documentConfig?.profile ? { profile: documentConfig.profile } : {}),
    generatedAt: new Date().toISOString(),
    templateSource: loaded.sourceLabel,
    ...(existsSync(getDocumentConfigPath(context.outputDir)) ? { configPath: formatOutputRelativePath(getDocumentConfigPath(context.outputDir), context) } : {}),
    sections: sections.map((section, index) =>
      buildDocumentSectionManifestItem(section, index, {
        context,
        documentConfig,
        previousManifest,
        tableMap,
        figureMap
      })
    )
  };
  await writeUtf8File(getDocumentSectionManifestPath(context), `${JSON.stringify(manifest, null, 2)}\n`);
  return { manifest, parsed: loaded.parsed, tables, figures };
}

function buildDocumentSectionManifestItem(
  section: SchemeTemplateTaskSection,
  index: number,
  input: {
    context: AgentToolExecutionContext;
    documentConfig?: DocumentConfig;
    previousManifest?: DocumentSectionManifest;
    tableMap: Map<string, SchemeTemplateTaskTable>;
    figureMap: Map<string, SchemeTemplateTaskFigure>;
  }
): DocumentSectionManifestItem {
  const sectionDir = buildDocumentSectionDirectoryName(section, index);
  const basePath = `${DOCUMENT_SECTION_DRAFT_DIRNAME}/${sectionDir}`;
  const draftPath = `${basePath}/draft.md`;
  const metaPath = `${basePath}/meta.json`;
  const tablesPath = `${basePath}/tables.json`;
  const figuresPath = `${basePath}/figures.json`;
  const evidenceReportPath = `${basePath}/evidence-report.md`;
  const evidenceReviewPath = `${basePath}/evidence-review.json`;
  const assetsDir = `${basePath}/assets`;
  const sectionGroupConfig = findDocumentConfigSectionGroupForSection(input.documentConfig, section.id);
  const meta = readDocumentSectionMetaIfPresent(join(input.context.outputDir, metaPath));
  const previous = input.previousManifest?.sections.find((item) => item.id === section.id);
  const draftExists = existsSync(join(input.context.outputDir, draftPath));
  const status = meta?.status || (draftExists ? previous?.status && previous.status !== "pending" ? previous.status : "drafted" : "pending");

  return {
    id: section.id,
    order: index + 1,
    number: section.number,
    title: section.title,
    headingLevel: section.number.split(".").filter(Boolean).length || 1,
    ...(section.anchors?.body?.tag ? { bodyTag: section.anchors.body.tag } : {}),
    ...(section.anchors?.body?.alias ? { bodyAlias: section.anchors.body.alias } : {}),
    ...(section.writingHint ? { writingHint: section.writingHint } : {}),
    paragraphTasks: buildSectionParagraphTasks(section, { tableMap: input.tableMap, figureMap: input.figureMap }),
    placeholders: section.placeholders ?? [],
    relatedTables: section.relatedTables ?? [],
    relatedFigures: section.relatedFigures ?? [],
    relatedTableSummaries: buildRelatedTableSummaries(section.relatedTables, input.tableMap) ?? [],
    relatedFigureSummaries: buildRelatedFigureSummaries(section.relatedFigures, input.figureMap) ?? [],
    ...(sectionGroupConfig?.id ? { sectionGroupId: sectionGroupConfig.id } : {}),
    ...(sectionGroupConfig?.title ? { sectionGroupTitle: sectionGroupConfig.title } : {}),
    writingRules: sectionGroupConfig?.writingRules ?? [],
    evidenceRules: sectionGroupConfig?.evidenceRules ?? [],
    requiredFacts: sectionGroupConfig?.requiredFacts ?? [],
    draftPath,
    metaPath,
    tablesPath,
    figuresPath,
    evidenceReportPath,
    evidenceReviewPath,
    assetsDir,
    status,
    ...(meta?.updatedAt ? { updatedAt: meta.updatedAt } : previous?.updatedAt ? { updatedAt: previous.updatedAt } : {})
  };
}

function getDocumentSectionManifestPath(context: AgentToolExecutionContext): string {
  return join(context.outputDir, DOCUMENT_SECTION_DRAFT_DIRNAME, DOCUMENT_SECTION_MANIFEST_FILENAME);
}

function buildDocumentSectionDirectoryName(section: SchemeTemplateTaskSection, index: number): string {
  const order = (index + 1).toString().padStart(3, "0");
  return sanitizeFileName(`${order}-${section.id}-${section.title}`) || `${order}-${section.id}`;
}

function readDocumentSectionManifestIfPresent(filePath: string): DocumentSectionManifest | undefined {
  if (!existsSync(filePath)) return undefined;
  try {
    const parsed = JSON.parse(readFileSync(filePath, "utf-8")) as unknown;
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return undefined;
    const record = parsed as Record<string, unknown>;
    if (record.version !== 1 || !Array.isArray(record.sections)) return undefined;
    return parsed as DocumentSectionManifest;
  } catch {
    return undefined;
  }
}

function readDocumentSectionMetaIfPresent(filePath: string): DocumentSectionMeta | undefined {
  if (!existsSync(filePath)) return undefined;
  try {
    const parsed = JSON.parse(readFileSync(filePath, "utf-8")) as unknown;
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return undefined;
    const record = parsed as Record<string, unknown>;
    const status = normalizeDocumentSectionStatus(typeof record.status === "string" ? record.status : "");
    return {
      sectionId: typeof record.sectionId === "string" ? record.sectionId : undefined,
      status: status && status !== "all" ? status : undefined,
      updatedAt: typeof record.updatedAt === "string" ? record.updatedAt : undefined,
      changeSummary: typeof record.changeSummary === "string" ? record.changeSummary : undefined
    };
  } catch {
    return undefined;
  }
}

function normalizeDocumentSectionStatus(value: string): DocumentSectionDraftStatus | "all" | undefined {
  const normalized = value.trim().toLowerCase();
  if (normalized === "all") return "all";
  if (normalized === "pending" || normalized === "drafted" || normalized === "edited" || normalized === "reviewed") return normalized;
  return undefined;
}

function matchesDocumentSectionListQuery(section: DocumentSectionManifestItem, query: string): boolean {
  if (!query) return true;
  const normalized = normalizeDraftSectionLookup(query);
  if (!normalized) return true;
  return [section.id, section.number, section.title, `${section.number} ${section.title}`]
    .map(normalizeDraftSectionLookup)
    .some((candidate) => candidate.includes(normalized) || normalized.includes(candidate));
}

function findDocumentSectionManifestMatches(
  value: string,
  sections: DocumentSectionManifestItem[]
): DocumentSectionManifestItem[] {
  const normalized = normalizeDraftSectionLookup(value);
  if (!normalized) return [];
  const exact = sections.filter((section) =>
    [section.id, section.number, section.title, `${section.number}${section.title}`, `${section.number} ${section.title}`]
      .map(normalizeDraftSectionLookup)
      .some((candidate) => candidate === normalized)
  );
  if (exact.length) return exact;
  return sections.filter((section) =>
    [section.id, section.number, section.title, `${section.number}${section.title}`, `${section.number} ${section.title}`]
      .map(normalizeDraftSectionLookup)
      .some((candidate) => candidate.includes(normalized) || normalized.includes(candidate))
  );
}

function resolveDocumentSectionManifestItems(
  inputs: string[],
  sections: DocumentSectionManifestItem[]
): { sections: DocumentSectionManifestItem[]; unknown: string[]; ambiguous: string[] } {
  if (!inputs.length) return { sections, unknown: [], ambiguous: [] };
  const selected: DocumentSectionManifestItem[] = [];
  const unknown: string[] = [];
  const ambiguous: string[] = [];
  const seen = new Set<string>();
  for (const input of inputs) {
    const matches = findDocumentSectionManifestMatches(input, sections);
    if (!matches.length) {
      unknown.push(input);
      continue;
    }
    if (matches.length > 1) {
      ambiguous.push(input);
      continue;
    }
    const section = matches[0];
    if (seen.has(section.id)) continue;
    seen.add(section.id);
    selected.push(section);
  }
  selected.sort((left, right) => left.order - right.order || left.id.localeCompare(right.id, "zh-CN"));
  return { sections: selected, unknown, ambiguous };
}

function formatDocumentSectionList(
  manifest: DocumentSectionManifest,
  sections: DocumentSectionManifestItem[],
  options: {
    query: string;
    status: DocumentSectionDraftStatus | "all";
    includeAssets: boolean;
    omitted: number;
    manifestPath: string;
    context: AgentToolExecutionContext;
  }
): string {
  const statusCounts = manifest.sections.reduce<Record<string, number>>((counts, section) => {
    counts[section.status] = (counts[section.status] ?? 0) + 1;
    return counts;
  }, {});
  const lines = [
    "list_document_sections completed",
    `文档：${manifest.title}`,
    manifest.profile ? `profile：${manifest.profile}` : "",
    `模板来源：${manifest.templateSource}`,
    `manifest：${formatOutputRelativePath(options.manifestPath, options.context)}`,
    `过滤：query=${options.query || "无"} status=${options.status}`,
    `状态统计：pending=${statusCounts.pending ?? 0} drafted=${statusCounts.drafted ?? 0} edited=${statusCounts.edited ?? 0} reviewed=${statusCounts.reviewed ?? 0}`,
    "",
    "sections:"
  ].filter(Boolean);

  for (const section of sections) {
    lines.push(
      [
        `- ${section.id} | ${section.number} ${section.title} | status=${section.status} | draft=${section.draftPath}`,
        section.bodyTag ? `  bodyTag: ${section.bodyTag}` : "",
        section.sectionGroupId ? `  sectionGroup: ${section.sectionGroupId} ${section.sectionGroupTitle || ""}`.trimEnd() : "",
        options.includeAssets && section.relatedTables.length ? `  tables: ${section.relatedTables.join("、")}` : "",
        options.includeAssets && section.relatedFigures.length ? `  figures: ${section.relatedFigures.join("、")}` : "",
        options.includeAssets ? `  assetsDir: ${section.assetsDir}` : ""
      ]
        .filter(Boolean)
        .join("\n")
    );
  }

  if (options.omitted > 0) {
    lines.push(`- 另有 ${options.omitted} 个章节未显示；可调整 max_sections/query/status。`);
  }
  lines.push(
    "",
    "用法：",
    "- 查询章节详情用 get_document_section({ section: sections[].id })。",
    "- 用户或 AI 微调正文后，用 update_document_section_draft 原地保存 draft.md。",
    "- 全部章节完成后，先用 audit_document_sections / revise_document_sections_evidence 做章节级审查和修订，再用 assemble_document_sections 合并终稿 Markdown。"
  );
  return compactText(lines.join("\n"), 80000);
}

function formatDocumentSectionDetail(
  section: DocumentSectionManifestItem,
  input: {
    context: AgentToolExecutionContext;
    draft?: string;
    includeAssets: boolean;
    tableMap: Map<string, SchemeTemplateTaskTable>;
    figureMap: Map<string, SchemeTemplateTaskFigure>;
  }
): string {
  const tablesJsonPath = join(input.context.outputDir, section.tablesPath);
  const figuresJsonPath = join(input.context.outputDir, section.figuresPath);
  const assetsDirPath = join(input.context.outputDir, section.assetsDir);
  const assetFiles = input.includeAssets && existsSync(assetsDirPath)
    ? readdirSync(assetsDirPath, { withFileTypes: true }).filter((entry) => entry.isFile()).map((entry) => entry.name)
    : [];
  const sectionContext = {
    id: section.id,
    number: section.number,
    title: section.title,
    order: section.order,
    headingLevel: section.headingLevel,
    status: section.status,
    bodyTag: section.bodyTag,
    bodyAlias: section.bodyAlias,
    writingHint: section.writingHint,
    paragraphTasks: section.paragraphTasks,
    placeholders: section.placeholders,
    sectionGroup: section.sectionGroupId ? { id: section.sectionGroupId, title: section.sectionGroupTitle } : undefined,
    writingRules: section.writingRules,
    evidenceRules: section.evidenceRules,
    requiredFacts: section.requiredFacts,
    paths: {
      draft: section.draftPath,
      meta: section.metaPath,
      tables: section.tablesPath,
      figures: section.figuresPath,
      evidenceReport: section.evidenceReportPath,
      evidenceReview: section.evidenceReviewPath,
      assets: section.assetsDir
    },
    relatedTables: section.relatedTables.map((id) => {
      const table = input.tableMap.get(id);
      return {
        id,
        caption: table?.caption,
        purpose: table?.purpose,
        header: table?.header,
        writeStrategy: table?.writeStrategy
      };
    }),
    relatedFigures: section.relatedFigures.map((id) => {
      const figure = input.figureMap.get(id);
      return {
        id,
        label: figure ? formatFigureDisplayName(figure) : id,
        caption: figure?.caption,
        purpose: figure?.purpose
      };
    })
  };
  const lines = [
    "get_document_section completed",
    `章节：${section.id} | ${section.number} ${section.title}`,
    `状态：${section.status}`,
    `草稿：${section.draftPath}${existsSync(join(input.context.outputDir, section.draftPath)) ? "" : "（未生成）"}`,
    "",
    "section_context:",
    JSON.stringify(sectionContext, null, 2),
    "",
    input.draft !== undefined ? "draft.md:" : "",
    input.draft !== undefined ? compactText(input.draft, 30000) : "",
    input.includeAssets && existsSync(tablesJsonPath) ? "" : "",
    input.includeAssets && existsSync(tablesJsonPath) ? "tables.json:" : "",
    input.includeAssets && existsSync(tablesJsonPath) ? compactText(readFileSync(tablesJsonPath, "utf-8"), 20000) : "",
    input.includeAssets && existsSync(figuresJsonPath) ? "" : "",
    input.includeAssets && existsSync(figuresJsonPath) ? "figures.json:" : "",
    input.includeAssets && existsSync(figuresJsonPath) ? compactText(readFileSync(figuresJsonPath, "utf-8"), 20000) : "",
    input.includeAssets ? "" : "",
    input.includeAssets ? `asset_files: ${assetFiles.join("、") || "无"}` : ""
  ].filter((line) => line !== "");
  return compactText(lines.join("\n"), 100000);
}

function formatDocumentSectionDraftMarkdown(section: DocumentSectionManifestItem, content: string): string {
  const body = stripDocumentSectionLeadingHeading(section, normalizeDraftContent(content));
  const rules = [
    section.writingHint ? `写作提示：${section.writingHint}` : "",
    ...section.paragraphTasks.map((task, index) => `段落 ${index + 1}：${task}`),
    section.relatedTableSummaries.length ? `关联表格：${section.relatedTableSummaries.join("；")}（正文只做自然引入，表格原始内容可存 tables.json）` : "",
    section.relatedFigureSummaries.length ? `关联图示：${section.relatedFigureSummaries.join("；")}（正文只做自然引入，图片资产放 assets/）` : "",
    section.bodyTag ? `Word body tag：${section.bodyTag}` : ""
  ].filter(Boolean);
  return [
    `# ${section.number} ${section.title}`,
    "",
    "<!-- document-section-meta",
    `id: ${section.id}`,
    `number: ${section.number}`,
    `title: ${section.title}`,
    ...rules.map((rule) => `- ${rule}`),
    "-->",
    "",
    body || "本节内容待补充/需确认。"
  ].join("\n");
}

function stripDocumentSectionLeadingHeading(section: DocumentSectionManifestItem, content: string): string {
  const lines = content.replace(/\r\n/g, "\n").split("\n");
  const firstNonEmptyIndex = lines.findIndex((line) => line.trim());
  if (firstNonEmptyIndex < 0) return "";
  const first = lines[firstNonEmptyIndex].trim();
  if (!/^#{1,6}\s+/.test(first)) return content.trim();
  const heading = first.replace(/^#{1,6}\s+/, "").trim();
  const normalizedHeading = normalizeDraftSectionLookup(heading);
  const normalizedSection = normalizeDraftSectionLookup(`${section.number} ${section.title}`);
  if (normalizedHeading !== normalizeDraftSectionLookup(section.title) && normalizedHeading !== normalizedSection) {
    return content.trim();
  }
  const nextLines = [...lines.slice(0, firstNonEmptyIndex), ...lines.slice(firstNonEmptyIndex + 1)];
  return nextLines.join("\n").trim();
}

function readDocumentSectionDraftBody(filePath: string): string {
  const raw = readFileSync(filePath, "utf-8");
  const withoutMeta = raw.replace(/<!--\s*document-section-meta[\s\S]*?-->\s*/gi, "");
  const lines = withoutMeta.replace(/\r\n/g, "\n").split("\n");
  const firstHeadingIndex = lines.findIndex((line) => /^#\s+/.test(line.trim()));
  if (firstHeadingIndex >= 0) {
    return normalizeDraftContent(lines.slice(firstHeadingIndex + 1).join("\n"));
  }
  return normalizeDraftContent(withoutMeta);
}

function parseJsonStringArg(value: string, label: string): unknown {
  try {
    return JSON.parse(value) as unknown;
  } catch {
    throw new Error(`update_document_section_draft.${label} 不是合法 JSON`);
  }
}

function loadDocumentSectionForAssembly(
  item: DocumentSectionManifestItem,
  context: AgentToolExecutionContext
): { item: DocumentSectionManifestItem; body: string; exists: boolean } {
  const draftPath = join(context.outputDir, item.draftPath);
  if (!existsSync(draftPath)) {
    return {
      item,
      body: "本节内容待补充/需确认。",
      exists: false
    };
  }
  const body = readDocumentSectionDraftBody(draftPath);
  return {
    item,
    body: body || "本节内容待补充/需确认。",
    exists: true
  };
}

function formatMergedDocumentSectionMarkdown(
  sections: Array<{ item: DocumentSectionManifestItem; body: string }>,
  options: { title: string }
): string {
  const lines = [`# ${options.title}`];
  for (const section of sections) {
    lines.push("", `## ${section.item.number} ${section.item.title}`, "", section.body.trim() || "本节内容待补充/需确认。");
  }
  return normalizeFinalMarkdown(lines.join("\n"));
}

function formatDocumentConfigFactLines(config: DocumentConfig): string[] {
  return config.facts.length
    ? config.facts.map((fact) => `- ${fact.key}：${fact.value}${fact.source ? `（来源：${fact.source}）` : ""}`)
    : ["- 无"];
}

function formatDocumentConfigGapLines(config: DocumentConfig): string[] {
  return config.gaps.length ? config.gaps.map((gap) => `- ${gap}`) : ["- 无"];
}

function formatDocumentGenerationPlanningLines(planning: DocumentConfig["planning"]): string[] {
  if (!planning) return ["- 无"];
  return [
    `- 规划摘要：${planning.summary}`,
    ...planning.assumptions.map((item) => `- 假设：${item}`),
    ...planning.risks.map((item) => `- 风险：${item}`)
  ];
}

function uniqueTextItems(values: string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const value of values.map((item) => item.trim()).filter(Boolean)) {
    const normalized = normalizeDraftSectionLookup(value);
    if (seen.has(normalized)) continue;
    seen.add(normalized);
    result.push(value);
  }
  return result;
}

interface DocumentEvidenceFinding {
  lineNumber: number;
  level: "high" | "medium";
  reason: string;
  text: string;
  suggestion: string;
}

async function executeAuditDocumentEvidence(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const markdownInputPath = readStringArg(args, "markdown_path");
  if (!markdownInputPath) {
    return {
      toolName: "audit_document_evidence",
      summary: "缺少 Markdown 路径",
      content: "audit_document_evidence failed: missing markdown_path"
    };
  }

  const markdownPath = resolveExistingOutputToolPath(markdownInputPath, context, "audit_document_evidence");
  if (!existsSync(markdownPath)) {
    throw new Error(`audit_document_evidence Markdown 不存在：${basename(markdownPath)}`);
  }
  if (extname(markdownPath).toLowerCase() !== ".md") {
    throw new Error("audit_document_evidence.markdown_path 只能读取 .md 文件");
  }

  const documentConfig = loadDocumentConfig(getDocumentConfigPath(context.outputDir));
  const markdown = readFileSync(markdownPath, "utf-8");
  const maxFindings = Math.min(Math.max(Math.trunc(readNumberArg(args, "max_findings", 80)), 1), 200);
  const allFindings = auditMarkdownEvidence(markdown, documentConfig);
  const findings = allFindings.slice(0, maxFindings);
  const reportName = sanitizeFileName(readStringArg(args, "name") || `${basename(markdownPath, extname(markdownPath))}-证据审查报告.md`);
  const reportPath = createOutputPath(context, reportName.toLowerCase().endsWith(".md") ? reportName : `${reportName}.md`);
  const report = formatDocumentEvidenceAuditReport({
    markdownPath,
    documentConfig,
    findings,
    truncated: allFindings.length > findings.length,
    context
  });
  await writeUtf8File(reportPath, report);

  const highRiskCount = findings.filter((finding) => finding.level === "high").length;
  return {
    toolName: "audit_document_evidence",
    summary: findings.length
      ? `发现 ${findings.length} 条疑似证据风险，其中高风险 ${highRiskCount} 条`
      : "未发现明显无证据确定表述",
    content: [
      "audit_document_evidence completed",
      `审查文件：${formatOutputRelativePath(markdownPath, context)}`,
      `报告：${formatOutputRelativePath(reportPath, context)}`,
      `已确认事实数：${documentConfig?.facts.length ?? 0}`,
      `待补充信息数：${documentConfig?.gaps.length ?? 0}`,
      `风险项：${findings.length}`,
      findings.length ? "注意：存在风险项时，不要直接进入最终 Word 交付；应先修正文稿或向用户确认。" : "可以继续进入 write_document_word。"
    ].join("\n"),
    artifactPath: reportPath
  };
}

function auditMarkdownEvidence(markdown: string, documentConfig?: DocumentConfig): DocumentEvidenceFinding[] {
  const factValues = buildConfirmedFactTokens(documentConfig);
  const gapTokens = buildGapTokens(documentConfig);
  const findings: DocumentEvidenceFinding[] = [];
  const lines = markdown.replace(/\r\n/g, "\n").split("\n");

  lines.forEach((line, index) => {
    const trimmed = line.trim();
    if (!shouldAuditMarkdownLine(trimmed)) return;

    const normalizedLine = normalizeEvidenceLookup(trimmed);
    const matchedGap = gapTokens.find((gap) => normalizedLine.includes(gap.normalized));
    if (matchedGap && !/待补充|需确认|待确认|未明确|缺少|暂无/.test(trimmed)) {
      findings.push({
        lineNumber: index + 1,
        level: "high",
        reason: `待补充信息被写成确定表述：${matchedGap.label}`,
        text: trimmed,
        suggestion: "将该句改为待补充/需确认，或先让用户/附件提供明确证据。"
      });
      return;
    }

    if (!looksLikeDefinitiveProjectFact(trimmed)) return;
    if (hasConfirmedEvidence(trimmed, factValues)) return;

    findings.push({
      lineNumber: index + 1,
      level: containsHighRiskDefinitivePhrase(trimmed) ? "high" : "medium",
      reason: "疑似无已确认事实支撑的确定性项目现状",
      text: trimmed,
      suggestion: "补充事实来源；无法确认时改为待补充/需确认，避免写成已建设或已采用。"
    });
  });

  if (!documentConfig?.facts.length) {
    findings.unshift({
      lineNumber: 0,
      level: "high",
      reason: "document-config.json 中没有已确认事实",
      text: "当前配置未记录任何已确认事实。",
      suggestion: "先通过 remember_project 或附件读取沉淀事实，再生成正式交付文档。"
    });
  }

  return findings;
}

function buildConfirmedFactTokens(documentConfig?: DocumentConfig): string[] {
  return uniqueTextItems(
    (documentConfig?.facts ?? [])
      .flatMap((fact) => [fact.value, `${fact.key}${fact.value}`, `${fact.key}：${fact.value}`])
      .map((value) => value.trim())
      .filter((value) => value.length >= 2)
  ).map((value) => normalizeEvidenceLookup(value));
}

function buildGapTokens(documentConfig?: DocumentConfig): Array<{ label: string; normalized: string }> {
  return uniqueTextItems(documentConfig?.gaps ?? [])
    .map((gap) => ({
      label: gap,
      normalized: normalizeEvidenceLookup(gap)
    }))
    .filter((gap) => gap.normalized.length >= 2);
}

function shouldAuditMarkdownLine(line: string): boolean {
  if (!line || line.length < 8) return false;
  if (/^#{1,6}\s+/.test(line)) return false;
  if (/^[-*]\s*$/.test(line)) return false;
  if (/^\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)+\|?$/.test(line)) return false;
  if (/待补充|需确认|待确认|未明确|暂无/.test(line)) return false;
  return true;
}

function looksLikeDefinitiveProjectFact(line: string): boolean {
  if (/应当|应|需|需要|建议|宜|可|拟|计划|目标|原则|要求|用于|可用于|本方案|后续/.test(line)) return false;
  return /(已|现已|目前|实际|采用|部署|配置|建设|接入|实现|完成|位于|包含|包括|提供|使用|存储|传输|运行|承载|安装|设置)/.test(line);
}

function containsHighRiskDefinitivePhrase(line: string): boolean {
  return /(已部署|已建设|已接入|已采用|已配置|已完成|实际部署|正式运行|位于|型号|服务器密码机|签名验签|密码产品)/.test(line);
}

function hasConfirmedEvidence(line: string, factTokens: string[]): boolean {
  const normalizedLine = normalizeEvidenceLookup(line);
  return factTokens.some((token) => token.length >= 2 && normalizedLine.includes(token));
}

function normalizeEvidenceLookup(value: string): string {
  return value
    .toLowerCase()
    .replace(/[\s.。．、,，:：;；()（）\[\]【】《》"'“”‘’_-]+/g, "")
    .trim();
}

function formatDocumentEvidenceAuditReport(input: {
  markdownPath: string;
  documentConfig?: DocumentConfig;
  findings: DocumentEvidenceFinding[];
  truncated: boolean;
  context: AgentToolExecutionContext;
}): string {
  const lines = [
    "# 文档证据审查报告",
    "",
    `审查文件：${formatOutputRelativePath(input.markdownPath, input.context)}`,
    `文档配置：${input.documentConfig ? "已读取 document-config.json" : "未读取 document-config.json"}`,
    `已确认事实数：${input.documentConfig?.facts.length ?? 0}`,
    `待补充信息数：${input.documentConfig?.gaps.length ?? 0}`,
    `风险项：${input.findings.length}${input.truncated ? "（已截断）" : ""}`,
    "",
    "## 结论",
    "",
    input.findings.length
      ? "存在疑似缺少证据支撑的确定性表述，建议先修正文稿或向用户确认，再生成最终 Word。"
      : "未发现明显无证据确定表述，可继续生成 Word。",
    "",
    "## 风险项"
  ];

  if (!input.findings.length) {
    lines.push("", "无。");
  } else {
    lines.push("", "| 行号 | 风险 | 原文 | 建议 |", "| --- | --- | --- | --- |");
    for (const finding of input.findings) {
      lines.push(
        `| ${finding.lineNumber || "全局"} | ${finding.level === "high" ? "高" : "中"}：${escapeMarkdownTableCell(finding.reason)} | ${escapeMarkdownTableCell(finding.text)} | ${escapeMarkdownTableCell(finding.suggestion)} |`
      );
    }
  }

  lines.push(
    "",
    "## 审查依据",
    "",
    "已确认事实：",
    ...(input.documentConfig?.facts.length ? input.documentConfig.facts.map((fact) => `- ${fact.key}：${fact.value}${fact.source ? `（来源：${fact.source}）` : ""}`) : ["- 无"]),
    "",
    "待补充/需确认：",
    ...(input.documentConfig?.gaps.length ? input.documentConfig.gaps.map((gap) => `- ${gap}`) : ["- 无"])
  );

  return lines.join("\n");
}

function escapeMarkdownTableCell(value: string): string {
  return value.replace(/\|/g, "\\|").replace(/\r?\n/g, " ").trim();
}

async function executeReviseDocumentEvidence(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const markdownInputPath = readStringArg(args, "markdown_path");
  if (!markdownInputPath) {
    return {
      toolName: "revise_document_evidence",
      summary: "缺少 Markdown 路径",
      content: "revise_document_evidence failed: missing markdown_path"
    };
  }

  const markdownPath = resolveExistingOutputToolPath(markdownInputPath, context, "revise_document_evidence");
  if (!existsSync(markdownPath)) {
    throw new Error(`revise_document_evidence Markdown 不存在：${basename(markdownPath)}`);
  }
  if (extname(markdownPath).toLowerCase() !== ".md") {
    throw new Error("revise_document_evidence.markdown_path 只能读取 .md 文件");
  }

  const documentConfig = loadDocumentConfig(getDocumentConfigPath(context.outputDir));
  const markdown = readFileSync(markdownPath, "utf-8");
  const maxRewrites = Math.min(Math.max(Math.trunc(readNumberArg(args, "max_rewrites", 80)), 1), 200);
  const findings = auditMarkdownEvidence(markdown, documentConfig);
  const revision = reviseMarkdownEvidenceLines(markdown, findings, maxRewrites);
  const reportName = sanitizeFileName(readStringArg(args, "name") || `${basename(markdownPath, extname(markdownPath))}-证据修订.md`);
  const outputPath = createOutputPath(context, reportName.toLowerCase().endsWith(".md") ? reportName : `${reportName}.md`);

  if (!revision.rewriteCount) {
    return {
      toolName: "revise_document_evidence",
      summary: "没有可自动修订的正文行",
      content: [
        "revise_document_evidence completed",
        `源文件：${formatOutputRelativePath(markdownPath, context)}`,
        `风险项：${findings.length}`,
        "没有写出新文件；如果仍有风险项，可能位于全局配置或表格行，需要人工确认。"
      ].join("\n")
    };
  }

  await writeUtf8File(outputPath, revision.markdown);

  return {
    toolName: "revise_document_evidence",
    summary: `已修订 ${revision.rewriteCount} 行无证据确定表述`,
    content: [
      "revise_document_evidence completed",
      `源文件：${formatOutputRelativePath(markdownPath, context)}`,
      `修订文件：${formatOutputRelativePath(outputPath, context)}`,
      `风险项：${findings.length}`,
      `已修订行：${revision.rewriteCount}`,
      revision.skippedCount ? `跳过行：${revision.skippedCount}（表格或全局项需人工确认）` : "",
      "",
      "next_audit_call:",
      JSON.stringify({ markdown_path: formatOutputRelativePath(outputPath, context) }, null, 2),
      "",
      "next_word_call:",
      JSON.stringify({ markdown_path: formatOutputRelativePath(outputPath, context) }, null, 2)
    ]
      .filter(Boolean)
      .join("\n"),
    artifactPath: outputPath
  };
}

function reviseMarkdownEvidenceLines(
  markdown: string,
  findings: DocumentEvidenceFinding[],
  maxRewrites: number
): { markdown: string; rewriteCount: number; skippedCount: number } {
  const lines = markdown.replace(/\r\n/g, "\n").split("\n");
  const rewrittenLines = new Set<number>();
  let skippedCount = 0;

  for (const finding of findings) {
    if (rewrittenLines.size >= maxRewrites) break;
    if (finding.lineNumber <= 0) {
      skippedCount += 1;
      continue;
    }
    const index = finding.lineNumber - 1;
    const currentLine = lines[index];
    if (!currentLine || !canRewriteEvidenceLine(currentLine)) {
      skippedCount += 1;
      continue;
    }
    lines[index] = rewriteEvidenceLine(currentLine, finding);
    rewrittenLines.add(index);
  }

  return {
    markdown: normalizeFinalMarkdown(lines.join("\n")),
    rewriteCount: rewrittenLines.size,
    skippedCount
  };
}

function canRewriteEvidenceLine(line: string): boolean {
  const trimmed = line.trim();
  if (!trimmed) return false;
  if (/^#{1,6}\s+/.test(trimmed)) return false;
  if (/^\|/.test(trimmed)) return false;
  if (/待补充|需确认|待确认|未明确|暂无/.test(trimmed)) return false;
  return true;
}

function rewriteEvidenceLine(line: string, finding: DocumentEvidenceFinding): string {
  const indent = line.match(/^\s*/)?.[0] ?? "";
  const bullet = line.slice(indent.length).match(/^((?:[-*]|\d+[.)])\s+)/)?.[1] ?? "";
  const prefix = `${indent}${bullet}`;
  const original = line.slice(prefix.length).trim();
  const revised = stripDefinitiveEvidencePhrases(original);
  return `${prefix}待补充/需确认：${revised}（${finding.reason}，需补充事实来源后再写成确定表述。）`;
}

function stripDefinitiveEvidencePhrases(value: string): string {
  return value
    .replace(/已部署/g, "是否部署")
    .replace(/已建设/g, "是否建设")
    .replace(/已接入/g, "是否接入")
    .replace(/已采用/g, "是否采用")
    .replace(/已配置/g, "是否配置")
    .replace(/已完成/g, "是否完成")
    .replace(/现已/g, "是否")
    .replace(/目前/g, "当前待确认")
    .trim();
}

async function executeWriteDocumentWord(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const markdownInputPath = readStringArg(args, "markdown_path");
  if (!markdownInputPath) {
    return {
      toolName: "write_document_word",
      summary: "缺少 Markdown 路径",
      content: "write_document_word failed: missing markdown_path"
    };
  }

  const markdownPath = resolveExistingOutputToolPath(markdownInputPath, context, "write_document_word");
  if (!existsSync(markdownPath)) {
    throw new Error(`write_document_word Markdown 不存在：${basename(markdownPath)}`);
  }
  if (extname(markdownPath).toLowerCase() !== ".md") {
    throw new Error("write_document_word.markdown_path 只能读取 .md 文件");
  }

  const documentConfig = loadDocumentConfig(getDocumentConfigPath(context.outputDir));
  const wordTemplate = mergeWordTemplateConfig(args, documentConfig?.wordTemplate);
  const templatePath = resolveWordTemplateToolPath(wordTemplate.templatePath, context, "write_document_word.template_path");
  const templateJsonPath = wordTemplate.templateJsonPath
    ? resolveOptionalWordTemplateJsonToolPath(wordTemplate.templateJsonPath, context, "write_document_word.template_json_path")
    : undefined;
  const renderMode = wordTemplate.renderMode || "full_document";
  const markdown = readFileSync(markdownPath, "utf-8");
  const templateTables = readTemplateTablesArg(args);
  const explicitTemplateCells = readTemplateCellsArg(args);
  const templateCells =
    explicitTemplateCells ??
    (templateTables?.length ? undefined : await buildFallbackTemplateCellsForCurrentDocumentTemplate(context));
  const contentControls = readContentControlsArg(args);
  const diagrams = readDiagramAssetsArg(args, context, "write_document_word.diagrams");
  const title = documentConfig?.title || context.sessionTitle || basename(markdownPath, extname(markdownPath));
  const outputName = sanitizeFileName(
    readStringArg(args, "name") ||
      `${title}${wordTemplate.outputNameSuffix && !title.includes(wordTemplate.outputNameSuffix) ? wordTemplate.outputNameSuffix : ""}.docx`
  );
  const finalName = outputName.toLowerCase().endsWith(".docx") ? outputName : `${outputName}.docx`;
  const outputPath = createOutputPath(context, finalName);
  const result = await writeSchemeDocxFromTemplate(templatePath, outputPath, {
    prompt: context.userPrompt || title,
    memory: context.memory,
    generatedMarkdown: markdown,
    fields: readObjectArg(args, "fields"),
    templateFields: readTemplateFieldsArg(args),
    templateTables,
    templateCells,
    contentControls,
    diagrams,
    renderMode,
    templateJsonPath
  });

  return {
    toolName: "write_document_word",
    summary: `已生成 Word ${result.fileName}`,
    content: [
      `write_document_word completed: ${result.outputPath}`,
      `Markdown：${formatOutputRelativePath(markdownPath, context)}`,
      `Word 模板：${formatToolSourcePath(templatePath, context)}`,
      templateJsonPath ? `模板结构 JSON：${formatToolSourcePath(templateJsonPath, context)}` : "模板结构 JSON：未使用，已从 docx 解析锚点",
      `写入模式：${result.renderMode}`,
      `追加/替换正文：${result.appendedMarkdown ? "是" : "否"}`,
      `模板占位符替换次数：${result.templateReplacementCount}`,
      `模板整表替换次数：${result.templateTableReplacementCount}`,
      `模板表格单元格替换次数：${result.templateCellReplacementCount}`,
      `Content Control 替换次数：${result.contentControlReplacementCount}`,
      `嵌入图示：${result.embeddedDiagrams.join("、") || "无"}`,
      `模板锚点：${formatTemplateAnchorsUsed(result.templateAnchorsUsed)}`
    ].join("\n"),
    artifactPath: result.outputPath
  };
}

async function buildFallbackTemplateCellsForCurrentDocumentTemplate(
  context: AgentToolExecutionContext
): Promise<TemplateCellReplacementInput[] | undefined> {
  let loaded: { parsed: SchemeTemplateTaskJson; sourceLabel: string } | undefined;
  try {
    loaded = await loadDocumentTemplateAssetTaskJson(context);
  } catch {
    return undefined;
  }

  const cells = readTemplateTaskTables(loaded.parsed.tables)
    .filter((table) => !isDynamicTemplateTable(table))
    .flatMap((table) =>
      (table.rows ?? [])
        .flatMap((row) => row.cells)
        .filter((cell) => isFillableTemplateCell(cell))
        .map((cell) => ({
          tableId: table.id,
          ...(table.caption ? { caption: table.caption } : {}),
          rowIndex: cell.rowIndex,
          cellIndex: cell.cellIndex,
          columnIndex: cell.columnIndex,
          value: inferFallbackTemplateCellValue(table, cell)
        }))
    );
  return cells.length ? cells : undefined;
}

function inferFallbackTemplateCellValue(table: SchemeTemplateTaskTable, cell: SchemeTemplateTaskCell): string {
  const header = table.header?.[cell.columnIndex] || "";
  if (/^序号$/.test(header)) return `${cell.rowIndex}.`;
  return "待补充/需确认";
}

function formatTemplateAnchorsUsed(anchorIds: string[]): string {
  const uniqueIds = Array.from(new Set(anchorIds.filter(Boolean)));
  if (!uniqueIds.length) return "未使用";
  const preview = uniqueIds.slice(0, 20).join("、");
  return uniqueIds.length > 20 ? `${uniqueIds.length} 个（${preview} 等）` : preview;
}

function mergeWordTemplateConfig(
  args: Record<string, unknown>,
  config: DocumentWordTemplateConfig | undefined
): DocumentWordTemplateConfig {
  const requestedRenderMode = readOptionalRenderModeArg(args);
  const renderMode = requestedRenderMode || config?.renderMode || "full_document";
  return {
    templatePath: readStringArg(args, "template_path") || config?.templatePath,
    templateJsonPath: readStringArg(args, "template_json_path") || config?.templateJsonPath,
    renderMode,
    outputNameSuffix: config?.outputNameSuffix
  };
}

function readOptionalRenderModeArg(args: Record<string, unknown>): "template_sections" | "append" | "full_document" | undefined {
  const value = readStringArg(args, "render_mode") || readStringArg(args, "renderMode");
  if (value === "append" || value === "full_document" || value === "template_sections") return value;
  return undefined;
}

function resolveWordTemplateToolPath(inputPath: string | undefined, context: AgentToolExecutionContext, toolName: string): string {
  if (!inputPath) return getBuiltInTemplateDocxPath(context.docsDir);
  const filePath = resolveDocsOrOutputToolPath(inputPath, context);
  assertPathInside(filePath, getReadableToolSourceDirs(context), toolName);
  if (!existsSync(filePath)) {
    throw new Error(`${toolName} 文件不存在：${basename(filePath)}`);
  }
  return filePath;
}

function resolveDocumentProfileToolPath(inputPath: string, context: AgentToolExecutionContext, toolName: string): string {
  const filePath = resolveDocsOrOutputToolPath(inputPath, context);
  assertPathInside(filePath, getReadableToolSourceDirs(context), toolName);
  if (!existsSync(filePath)) {
    throw new Error(`${toolName} 文件不存在：${basename(filePath)}`);
  }
  if (extname(filePath).toLowerCase() !== ".json") {
    throw new Error(`${toolName} 只能读取 .json 文件`);
  }
  return filePath;
}

function resolveOptionalWordTemplateJsonToolPath(inputPath: string, context: AgentToolExecutionContext, toolName: string): string | undefined {
  const filePath = resolveDocsOrOutputToolPath(inputPath, context);
  assertPathInside(filePath, getReadableToolSourceDirs(context), toolName);
  return existsSync(filePath) ? filePath : undefined;
}

function resolveDocsOrOutputToolPath(inputPath: string, context: AgentToolExecutionContext): string {
  const docsOrRootPath = resolveToolPath(inputPath, context);
  if (isPathInside(docsOrRootPath, getReadableToolSourceDirs(context))) return docsOrRootPath;
  return resolveOutputToolPath(inputPath, context);
}

function formatToolSourcePath(filePath: string, context: AgentToolExecutionContext): string {
  if (isPathInside(filePath, [context.outputDir])) return formatOutputRelativePath(filePath, context);
  if (context.globalOutputDir && isPathInside(filePath, [context.globalOutputDir])) {
    return relative(context.globalOutputDir, filePath).replace(/\\/g, "/");
  }
  if (isPathInside(filePath, [context.docsDir])) return `docs/${relative(context.docsDir, filePath).replace(/\\/g, "/")}`;
  return filePath;
}

function stripDocumentRuleComments(content: string): string {
  return content.replace(/<!--\s*document-section-meta[\s\S]*?-->\s*/gi, "");
}

function finalizeDocumentMarkdown(markdown: string, rules: string[]): string {
  const withoutRules = stripDocumentRuleComments(markdown);
  const withoutToolStatus = withoutRules
    .split(/\r?\n/)
    .filter((line) => !/^状态：/.test(line.trim()))
    .filter((line) => !/^章节文件：/.test(line.trim()))
    .join("\n");
  const strippedIds = stripRawTemplatePlanningText(withoutToolStatus);
  return normalizeFinalMarkdown(strippedIds);
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

function normalizeFinalMarkdown(markdown: string): string {
  return markdown
    .replace(/^```(?:markdown|md)?\s*/i, "")
    .replace(/```$/i, "")
    .replace(/\r\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
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

async function executeWriteFile(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const name = sanitizeFileName(readStringArg(args, "name") || "agent-output.md");
  const content = readStringArg(args, "content");
  if (!content) {
    return { toolName: "write_file", summary: "缺少写入内容", content: "write_file failed: missing content" };
  }

  const outputPath = createOutputPath(context, name);
  await writeUtf8File(outputPath, content);
  return {
    toolName: "write_file",
    summary: `已写入 ${basename(outputPath)}`,
    content: compactText(content, 12000),
    artifactPath: outputPath
  };
}

async function executeUpdateDocumentProfile(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const content = readStringArg(args, "content") || readStringArg(args, "profile_json") || readStringArg(args, "profileJson");
  if (!content) {
    return {
      toolName: "update_document_profile",
      summary: "缺少 profile 内容",
      content: "update_document_profile failed: missing content"
    };
  }

  const profilePath = resolveProfileUpdatePath(args, context);
  const templateId = basename(dirname(profilePath));
  if (templateId === DEFAULT_DOCUMENT_TEMPLATE_ID) {
    throw new Error("update_document_profile 不允许修改默认内置模板");
  }
  if (!existsSync(profilePath)) {
    throw new Error(`update_document_profile profile 不存在：${basename(dirname(profilePath))}`);
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(content) as unknown;
  } catch {
    throw new Error("update_document_profile.content 不是合法 JSON");
  }
  const normalizedProfile = parseDocumentProfileJson(JSON.stringify(parsed), templateId);
  if (!normalizedProfile) {
    throw new Error("update_document_profile.content 不是有效 DocumentProfile：至少需要 version/profile/sectionRules");
  }

  await writeUtf8File(profilePath, `${JSON.stringify(parsed, null, 2)}\n`);
  const changeSummary = readStringArg(args, "change_summary") || readStringArg(args, "changeSummary");
  return {
    toolName: "update_document_profile",
    summary: `已更新模板 profile：${templateId}`,
    content: [
      `update_document_profile completed: ${formatToolSourcePath(profilePath, context)}`,
      `模板 ID：${templateId}`,
      `章节规则数：${normalizedProfile.sectionRules.length}`,
      changeSummary ? `调整说明：${changeSummary}` : ""
    ].filter(Boolean).join("\n"),
    artifactPath: profilePath
  };
}

async function executeBuildDocumentConfig(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext
): Promise<AgentToolExecutionResult> {
  const profilePathArg = readStringArg(args, "profile_path") || readStringArg(args, "profilePath");
  const profile = readStringArg(args, "profile") || "generic_document";
  const projectContext = readStringArg(args, "project_context");
  const planning = readDocumentGenerationPlanningArg(args);
  const sectionGroupPlans = readDocumentSectionGroupPlansArg(args);
  const profilePath = profilePathArg ? resolveDocumentProfileToolPath(profilePathArg, context, "build_document_config.profile_path") : undefined;
  const namedProfilePath = getDocumentProfilePath(context.docsDir, profile);
  let profileSourcePath = profilePath ?? (existsSync(namedProfilePath) ? namedProfilePath : undefined);
  let profileConfig = profilePath ? loadDocumentProfileFromFile(profilePath, basename(profilePath, extname(profilePath))) : loadDocumentProfile(context.docsDir, profile);
  if (profilePath && !profileConfig) {
    throw new Error(`build_document_config.profile_path 无法读取有效 profile：${basename(profilePath)}`);
  }
  if (!profileConfig) {
    throw new Error(`build_document_config.profile 无法读取有效 profile：${profile}`);
  }
  const title =
    readStringArg(args, "title") ||
    profileConfig.title ||
    formatDefaultDocumentTitle(context.sessionTitle, profileConfig.titleSuffix || "密码应用方案");
  const sourceSummary = buildDocumentSourceSummary(
    {
      userPrompt: context.userPrompt,
      projectContext,
      memory: context.memory
    },
    12000
  );
  const profileWordTemplate = inferWordTemplateForProfile(profileConfig, profileSourcePath, context);
  const profileTemplateSections = await loadDocumentProfileTemplateSections(profileWordTemplate, context);
  const documentConfig = buildDocumentConfigFromProfile({
    profile: profileConfig,
    title,
    sourceSummary,
    memory: context.memory,
    planning,
    sectionGroupPlans,
    templateSections: profileTemplateSections,
    wordTemplate: profileWordTemplate
  });
  const configPath = getDocumentConfigPath(context.outputDir);
  await writeUtf8File(configPath, `${JSON.stringify(documentConfig, null, 2)}\n`);

  return {
    toolName: "build_document_config",
    summary: `已生成 document-config.json（${documentConfig.sectionGroups.length} 个章节组）`,
    content: [
      `profile_source: ${profileSourcePath ? formatToolSourcePath(profileSourcePath, context) : `docs/${relative(context.docsDir, getDocumentProfilePath(context.docsDir, profile)).replace(/\\/g, "/")}`}`,
      "",
      buildDocumentConfigSummary(documentConfig)
    ].filter((line) => line !== "").join("\n"),
    artifactPath: configPath
  };
}

function readDocumentGenerationPlanningArg(args: Record<string, unknown>): Partial<DocumentGenerationPlanning> | undefined {
  const summary = readStringArg(args, "generation_plan") || readStringArg(args, "generationPlan") || readStringArg(args, "plan");
  const assumptions = readStringListArg(args, "planning_assumptions").length
    ? readStringListArg(args, "planning_assumptions")
    : readStringListArg(args, "planningAssumptions");
  const risks = readStringListArg(args, "planning_risks").length
    ? readStringListArg(args, "planning_risks")
    : readStringListArg(args, "planningRisks");
  if (!summary && !assumptions.length && !risks.length) return undefined;
  return { summary, assumptions, risks };
}

function readDocumentSectionGroupPlansArg(args: Record<string, unknown>): DocumentConfigSectionGroupPlanInput[] {
  const value = args.section_group_plans ?? args.sectionGroupPlans;
  if (!Array.isArray(value)) return [];
  return value
    .map((item): DocumentConfigSectionGroupPlanInput | undefined => {
      if (!item || typeof item !== "object" || Array.isArray(item)) return undefined;
      const record = item as Record<string, unknown>;
      return {
        id: readLooseRecordString(record, "id"),
        sectionGroupId:
          readLooseRecordString(record, "section_group_id") ||
          readLooseRecordString(record, "sectionGroupId"),
        match: readLooseRecordString(record, "match"),
        title: readLooseRecordString(record, "title"),
        order: readLooseRecordNumber(record, "order"),
        objective: readLooseRecordString(record, "objective"),
        outline: readLooseRecordStringList(record, "outline"),
        keyPoints: readLooseRecordStringList(record, "key_points").length
          ? readLooseRecordStringList(record, "key_points")
          : readLooseRecordStringList(record, "keyPoints"),
        evidenceNeeds: readLooseRecordStringList(record, "evidence_needs").length
          ? readLooseRecordStringList(record, "evidence_needs")
          : readLooseRecordStringList(record, "evidenceNeeds"),
        openQuestions: readLooseRecordStringList(record, "open_questions").length
          ? readLooseRecordStringList(record, "open_questions")
          : readLooseRecordStringList(record, "openQuestions")
      } satisfies DocumentConfigSectionGroupPlanInput;
    })
    .filter((plan): plan is DocumentConfigSectionGroupPlanInput => Boolean(plan));
}

function readLooseRecordString(record: Record<string, unknown>, key: string): string {
  const value = record[key];
  return typeof value === "string" ? value.trim() : "";
}

function readLooseRecordStringList(record: Record<string, unknown>, key: string): string[] {
  const value = record[key];
  if (!Array.isArray(value)) return [];
  return value.map((item) => (typeof item === "string" ? item.trim() : "")).filter(Boolean);
}

function readLooseRecordNumber(record: Record<string, unknown>, key: string): number | undefined {
  const value = record[key];
  if (typeof value === "number" && Number.isFinite(value)) return Math.trunc(value);
  if (typeof value === "string" && value.trim() && Number.isFinite(Number(value))) return Math.trunc(Number(value));
  return undefined;
}

async function loadDocumentProfileTemplateSections(
  inferredWordTemplate: DocumentWordTemplateConfig | undefined,
  context: AgentToolExecutionContext
): Promise<DocumentConfigSectionLike[] | undefined> {
  const wordTemplate = inferredWordTemplate;
  if (!wordTemplate?.templatePath && !wordTemplate?.templateJsonPath) {
    const parsed = await loadSchemeTemplateTaskJson(context);
    return readTemplateTaskSections(parsed.sections).map((section) => buildDocumentConfigSectionLike(section));
  }

  const templateJsonPath = wordTemplate.templateJsonPath
    ? resolveOptionalWordTemplateJsonToolPath(wordTemplate.templateJsonPath, context, "build_document_config.wordTemplate.templateJsonPath")
    : undefined;
  if (templateJsonPath) {
    const parsedJson = readSchemeTemplateTaskJsonIfPresent(templateJsonPath);
    const sections = readTemplateTaskSections(parsedJson?.sections);
    if (sections.length) return sections.map((section) => buildDocumentConfigSectionLike(section));
  }

  const templateDocxPath = resolveWordTemplateToolPath(wordTemplate.templatePath, context, "build_document_config.wordTemplate.templatePath");
  try {
    const parsedDocx = await parseWordTemplateAnchorsFromDocx(templateDocxPath);
    const taskJson = buildSchemeTemplateTaskJsonFromWordTemplate(parsedDocx);
    return readTemplateTaskSections(taskJson.sections).map((section) => buildDocumentConfigSectionLike(section));
  } catch {
    return undefined;
  }
}

function inferWordTemplateForProfile(
  profile: DocumentProfile,
  profileSourcePath: string | undefined,
  context: AgentToolExecutionContext
): DocumentWordTemplateConfig | undefined {
  if (!profileSourcePath) return undefined;

  if (isPathInside(profileSourcePath, [join(context.docsDir, "document-profiles")])) {
    return {
      templatePath: BUILT_IN_TEMPLATE_DOCX_RELATIVE_PATH,
      templateJsonPath: BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH,
      renderMode: "template_sections",
      outputNameSuffix: profile.titleSuffix
    };
  }

  const templateDir = dirname(profileSourcePath);
  const templatePath = join(templateDir, "template.docx");
  if (!existsSync(templatePath)) return undefined;
  const templateJsonPath = join(templateDir, "template.json");
  const renderMode = inferTemplateRenderMode(templateJsonPath);
  return {
    templatePath: formatToolSourcePath(templatePath, context),
    ...(existsSync(templateJsonPath) ? { templateJsonPath: formatToolSourcePath(templateJsonPath, context) } : {}),
    renderMode,
    outputNameSuffix: profile.titleSuffix
  };
}

function inferTemplateRenderMode(templateJsonPath: string): DocumentWordTemplateConfig["renderMode"] {
  if (!existsSync(templateJsonPath)) return "full_document";
  try {
    const parsed = JSON.parse(readFileSync(templateJsonPath, "utf-8")) as { strategy?: { defaultWriteMode?: string }; sections?: unknown[] };
    const mode = parsed.strategy?.defaultWriteMode;
    if (mode === "template_sections" || mode === "append" || mode === "full_document") return mode;
    return Array.isArray(parsed.sections) && parsed.sections.length ? "template_sections" : "full_document";
  } catch {
    return "full_document";
  }
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

function createOutputFileName(name: string, stamp = createToolArtifactStamp()): string {
  const safeName = sanitizeFileName(name) || "agent-output";
  const extension = extname(safeName);
  const stem = extension ? basename(safeName, extension) : safeName;
  return `${stem || "agent-output"}-${stamp}${extension}`;
}

function createOutputPath(context: AgentToolExecutionContext, name: string, stamp?: string): string {
  return join(context.outputDir, createOutputFileName(name, stamp));
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
  if (normalizedInput === "document-templates" || normalizedInput.startsWith("document-templates/")) {
    return resolve(getTemplateLibraryOutputDir(context), normalizedInput);
  }

  return resolve(context.rootDir, inputPath);
}

function getTemplateLibraryOutputDir(context: AgentToolExecutionContext): string {
  return context.globalOutputDir || context.outputDir;
}

function getTemplateLibraryDir(context: AgentToolExecutionContext): string {
  return join(getTemplateLibraryOutputDir(context), "document-templates");
}

function getReadableToolSourceDirs(context: AgentToolExecutionContext): string[] {
  return Array.from(new Set([context.docsDir, context.outputDir, getTemplateLibraryDir(context)].map((dir) => resolve(dir))));
}

function resolveOutputToolPath(inputPath: string, context: AgentToolExecutionContext): string {
  if (isAbsolute(inputPath)) return resolve(inputPath);

  const normalizedInput = inputPath.replace(/\\/g, "/");
  if (normalizedInput === "document-templates" || normalizedInput.startsWith("document-templates/")) {
    return resolve(getTemplateLibraryOutputDir(context), normalizedInput);
  }

  const rootRelativePath = resolve(context.rootDir, inputPath);
  if (isPathInside(rootRelativePath, [context.outputDir])) {
    return rootRelativePath;
  }

  return resolve(context.outputDir, inputPath);
}

function resolveProfileUpdatePath(args: Record<string, unknown>, context: AgentToolExecutionContext): string {
  const profilePathArg = readStringArg(args, "profile_path") || readStringArg(args, "profilePath") || readStringArg(args, "path");
  const templateId = readStringArg(args, "template_id") || readStringArg(args, "templateId");
  const profilePath = profilePathArg
    ? resolveDocsOrOutputToolPath(profilePathArg, context)
    : templateId
      ? join(getTemplateLibraryDir(context), templateId, "profile.json")
      : undefined;

  if (!profilePath) {
    throw new Error("update_document_profile 需要 profile_path 或 template_id");
  }
  assertPathInside(profilePath, [getTemplateLibraryDir(context)], "update_document_profile.profile_path");
  if (basename(profilePath) !== "profile.json") {
    throw new Error("update_document_profile 只能更新 document-templates/<id>/profile.json");
  }
  if (extname(profilePath).toLowerCase() !== ".json") {
    throw new Error("update_document_profile.profile_path 只能是 .json 文件");
  }
  return profilePath;
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
    .filter((entry) => entry.isFile() && isGeneratedOutputFileNameMatch(entry.name, requestedName))
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

function isGeneratedOutputFileNameMatch(fileName: string, requestedName: string): boolean {
  if (fileName === requestedName) return true;
  if (fileName.endsWith(`-${requestedName}`)) return true;
  return stripGeneratedOutputSuffix(fileName) === requestedName;
}

function stripGeneratedOutputSuffix(fileName: string): string {
  const extension = extname(fileName);
  const stem = extension ? basename(fileName, extension) : fileName;
  return `${stem.replace(/-[a-z0-9]+-[a-z0-9]{6}$/i, "")}${extension}`;
}

function formatOutputRelativePath(filePath: string, context: AgentToolExecutionContext): string {
  if (!filePath) return "";
  return relative(context.outputDir, filePath).replace(/\\/g, "/");
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

function readKeyValueListArg(args: Record<string, unknown>, key: string): Array<{ key: string; value: string; source?: string }> {
  const value = args[key];
  if (!Array.isArray(value)) return [];
  return value
    .map((item) => {
      if (!item || typeof item !== "object" || Array.isArray(item)) return undefined;
      const record = item as Record<string, unknown>;
      const itemKey = typeof record.key === "string" ? record.key.trim() : "";
      const itemValue = typeof record.value === "string" ? record.value.trim() : "";
      const source = typeof record.source === "string" ? record.source.trim() : "";
      return itemKey && itemValue ? { key: itemKey, value: itemValue, ...(source ? { source } : {}) } : undefined;
    })
    .filter((item): item is { key: string; value: string; source?: string } => Boolean(item));
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

function readDiagramAssetsArg(
  args: Record<string, unknown>,
  context: AgentToolExecutionContext,
  toolName = "write_document_word.diagrams"
): SchemeDiagramAsset[] {
  const value = args.diagrams;
  const explicitDiagrams = Array.isArray(value)
    ? value
        .map((item) => normalizeDiagramAsset(item, context, toolName))
        .filter((item): item is SchemeDiagramAsset => Boolean(item))
    : [];
  if (explicitDiagrams.length) return explicitDiagrams;
  return dedupeDiagramAssets([
    ...extractDiagramAssetsFromMemory(context.memory, context, toolName),
    ...discoverDiagramAssetsFromOutput(context)
  ]);
}

function normalizeDiagramAsset(
  item: unknown,
  context: AgentToolExecutionContext,
  toolName = "write_document_word.diagrams"
): SchemeDiagramAsset | undefined {
  if (!item || typeof item !== "object" || Array.isArray(item)) return undefined;
  const record = item as Record<string, unknown>;
  const label = typeof record.label === "string" ? record.label.trim() : "";
  const rawPath = typeof record.path === "string" ? record.path.trim() : "";
  if (!label || !rawPath) return undefined;

  const path = resolveOutputToolPath(rawPath, context);
  assertPathInside(path, [context.outputDir], toolName);
  if (!existsSync(path)) return undefined;
  return {
    label,
    kind: typeof record.kind === "string" ? record.kind.trim() : undefined,
    figureId: typeof (record.figure_id ?? record.figureId) === "string" ? String(record.figure_id ?? record.figureId).trim() : undefined,
    path
  };
}

function extractDiagramAssetsFromMemory(
  memory: string,
  context: AgentToolExecutionContext,
  toolName = "write_document_word.diagrams"
): SchemeDiagramAsset[] {
  const assets: SchemeDiagramAsset[] = [];
  const pattern = /image_generate completed:\s*(.+?)\s*\n图示名称：(.+?)(?:\n图示类型：(.+?))?(?=\n## |\nimage_generate completed:|$)/gs;
  for (const match of memory.matchAll(pattern)) {
    const rawPath = match[1]?.trim();
    const label = match[2]?.trim();
    if (!rawPath || !label) continue;
    const asset = normalizeDiagramAsset({ path: rawPath, label, kind: match[3]?.trim() }, context, toolName);
    if (asset && !assets.some((item) => item.path === asset.path)) assets.push(asset);
  }
  return assets;
}

function discoverDiagramAssetsFromOutput(context: AgentToolExecutionContext): SchemeDiagramAsset[] {
  if (!existsSync(context.outputDir)) return [];
  const assets = readdirSync(context.outputDir, { withFileTypes: true })
    .filter((entry) => entry.isFile() && isDiagramAssetFileName(entry.name))
    .map((entry) => {
      const filePath = join(context.outputDir, entry.name);
      const label = inferDiagramLabelFromFileName(entry.name, context);
      if (!label) return undefined;
      return normalizeDiagramAsset(
        {
          path: filePath,
          label,
          kind: inferDiagramKindFromLabel(label)
        },
        context,
        "write_document_word.diagrams"
      );
    })
    .filter((asset): asset is SchemeDiagramAsset => Boolean(asset));
  return dedupeDiagramAssets(assets);
}

function isDiagramAssetFileName(fileName: string): boolean {
  const extension = extname(fileName).toLowerCase();
  return extension === ".png" || extension === ".jpg" || extension === ".jpeg" || extension === ".webp" || extension === ".svg";
}

function inferDiagramLabelFromFileName(fileName: string, context: AgentToolExecutionContext): string {
  const extension = extname(fileName);
  const stripped = basename(stripGeneratedOutputSuffix(fileName), extension);
  const sessionTitle = sanitizeFileName(context.sessionTitle || "");
  const withoutSession =
    sessionTitle && stripped.startsWith(`${sessionTitle}-`) ? stripped.slice(sessionTitle.length + 1) : stripped;
  const parts = withoutSession.split("-").map((part) => part.trim()).filter(Boolean);
  const figurePart = [...parts].reverse().find((part) => /图|架构|框架|流程|拓扑/.test(part));
  return (figurePart || withoutSession).replace(/[_\s]+/g, " ").trim();
}

function inferDiagramKindFromLabel(label: string): string {
  return /流程|过程|调用|读取|写入|签名|验签|鉴别/.test(label) ? "flow" : "architecture";
}

function dedupeDiagramAssets(assets: SchemeDiagramAsset[]): SchemeDiagramAsset[] {
  const seen = new Set<string>();
  const result: SchemeDiagramAsset[] = [];
  for (const asset of assets) {
    const key = resolve(asset.path).toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(asset);
  }
  return result;
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
