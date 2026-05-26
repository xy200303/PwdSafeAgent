export interface SchemeSectionRequirement {
  key: string;
  title: string;
  patterns: RegExp[];
}

export interface SchemeDiagramRequirement {
  key: string;
  label: string;
  patterns: RegExp[];
}

export const REQUIRED_SCHEME_SECTIONS: SchemeSectionRequirement[] = [
  {
    key: "background",
    title: "1. 背景",
    patterns: [/背景/, /系统建设规划/, /法律法规要求/]
  },
  {
    key: "overview",
    title: "2. 系统概述",
    patterns: [/系统概述/, /基本情况/, /计算平台现状/, /网络环境/, /业务应用现状/]
  },
  {
    key: "requirement-analysis",
    title: "3. 密码应用需求分析",
    patterns: [/密码应用需求分析/, /风险分析/, /物理和环境安全/, /网络和通[讯信]安全/, /应用和数据安全/]
  },
  {
    key: "goals-principles",
    title: "4. 安全目标及设计原则",
    patterns: [/安全目标/, /设计原则/, /建设依据/, /密码技术标准/]
  },
  {
    key: "crypto-design",
    title: "5. 密码应用设计",
    patterns: [/密码应用设计/, /密码应用技术框架/, /计算平台密码应用方案/, /业务应用的密码应用方案/]
  },
  {
    key: "key-management",
    title: "5.3.7. 密钥管理方式",
    patterns: [/密钥管理/, /密钥生存周期/, /密钥生成/, /密钥分发/, /密钥销毁/]
  },
  {
    key: "management",
    title: "6. 安全管理方案",
    patterns: [/安全管理方案/, /管理制度/, /人员管理/, /建设运行/, /应急处置/]
  },
  {
    key: "compliance",
    title: "7. 安全与合规性分析",
    patterns: [/安全与合规性分析/, /合规性分析/, /符合性/, /GB\/T\s*39786|39786/]
  },
  {
    key: "implementation",
    title: "8. 实施保障方案",
    patterns: [/实施保障方案/, /实施内容/, /实施计划/, /保障措施/, /经费概算/]
  }
];

export const REQUIRED_SCHEME_DIAGRAMS: SchemeDiagramRequirement[] = [
  {
    key: "network-architecture",
    label: "网络架构图",
    patterns: [/网络架构图/, /网络框架图/]
  },
  {
    key: "network-topology",
    label: "网络拓扑图",
    patterns: [/网络拓扑图/, /拓扑及边界/]
  },
  {
    key: "crypto-architecture",
    label: "密码应用技术架构图",
    patterns: [/密码应用技术架构图/, /密码应用技术框架图/, /密码服务架构图/]
  },
  {
    key: "business-flow",
    label: "典型业务密码应用流程图",
    patterns: [/典型业务密码应用流程图/, /密码应用工作流程图/, /业务流程图/]
  }
];

export function renderSchemeChapterGuide(): string {
  return [
    "Word 模板章节参考（整篇生成时使用模板 JSON 的真实 sections，write_word 会按不可见 SDT 锚写入 Word 样式模板）：",
    ...REQUIRED_SCHEME_SECTIONS.map((section) => `- ${section.title}`),
    "常用图示清单（需要图片时先调用 image_generate 生成，再按需传给 write_word）：",
    ...REQUIRED_SCHEME_DIAGRAMS.map((diagram) => `- ${diagram.label}`),
    "模板精确标注：docs/密码应用方案.template.json 包含不可见 SDT 锚、表格单元格、图片和题注锚点。",
    "推荐节奏：整篇交付时先读取 docs/密码应用方案.template.json，再调用 plan_scheme_batches 生成批次，create_word 创建模板副本，随后按批次调用 draft_scheme_sections 并行起草正文，再按 sections 顺序合并到 write_word.sections 批量写入；不要一次性生成整篇长文。",
    "表格和图示节奏：正文写完后调用 plan_scheme_assets，按 template_cells_plan 填表，按 image_generate_plan 生图，并用 diagrams.figure_id 精确嵌入对应图片锚点。",
    "写作要求：每节只处理对应模板节的事实，资料不足处标注待补充，避免通用套话；表格和图示放到正文写完后统一生成和嵌入。"
  ].join("\n");
}

export function findMissingSchemeSections(text: string): string[] {
  return REQUIRED_SCHEME_SECTIONS.filter((section) => !section.patterns.some((pattern) => pattern.test(text))).map(
    (section) => section.title
  );
}

export function findMissingSchemeDiagrams(text: string): string[] {
  return REQUIRED_SCHEME_DIAGRAMS.filter((diagram) => !diagram.patterns.some((pattern) => pattern.test(text))).map(
    (diagram) => diagram.label
  );
}
