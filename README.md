# PwdSafeAgent

PwdSafeAgent 是一个基于 Electron、React、Redux Toolkit、Radix UI、Incremark、OpenAI SDK 和 electron-vite 构建的桌面端密码应用方案生成 Agent。它面向密码应用方案编制场景，通过对话式交互收集项目资料，由 Agent 自主调用工具读取模板、分析附件、生成 Word/PDF/Markdown/图片等交付物，并在前端以类 Codex 的聊天流展示消息、工具调用过程和文件卡片。

## 核心能力

- 以 `docs/密码应用方案.docx` 作为方案模板资源，支持按需读取 Word/PDF/文本资料。
- Agent 对话运行时强制使用 `@mariozechner/pi-coding-agent`，模型配置保存在本地 `.env.local`。
- 不再内置或回退到自实现 OpenAI Chat Runtime；Agent 循环推理和工具调用决策由 Pi Agent 的 `AgentSession` 完成。
- 内置工具包括 `time`、`web_search`、`read_file`、`read_word`、`read_pdf`、`read_image`、`write_file`、`update_document_profile`、`build_document_config`、`list_document_sections`、`get_document_section`、`draft_document_sections`、`update_document_section_draft`、`audit_document_sections`、`revise_document_sections_evidence`、`polish_document_sections`、`assemble_document_sections`、`audit_document_evidence`、`revise_document_evidence`、`plan_document_assets`、`write_document_word`、`write_pdf`、`image_generate`、`exec_bash`、`send_file`。
- 整篇方案生成采用统一 section-first 流程：先从输入材料、项目档案和模板 profile 生成 `document-config.json`，再按 `template.json` 或 Word 动态解析出的真实章节生成 `document-sections/<section>/draft.md`，审查和微调后由 `assemble_document_sections` 合并终稿 Markdown，最后补齐图表并写入 Word。
- 支持 `docs/document-profiles/<profile>.json` 定义不同文档类型的生成规则、行文规则、证据规则、必需事实和 Word 渲染方式；内置 `generic_document` profile 基于 `docs/templates/密码应用方案.docx` 生成。
- 支持用户在设置中上传全局 Word 模板，按标题结构写入隐藏章节锚点，并在 `data/output/document-templates/<id>/` 生成 `template.json` 与可编辑的 `profile.json`；用户可通过修改 profile 自定义生成规则，后续按该 profile 完成章节起草和 Word 输出。
- 项目档案区分已确认事实、事实来源和待补充信息；生成正文时待补充项不会被改写成确定事实。
- 每个会话使用独立工作区 `data/workspaces/<sessionId>/`，附件导入、章节草稿和最终产物按会话隔离。
- 前端支持工具调用折叠详情、文件卡片、系统打开、文件夹定位、PDF.js / docx-preview 预览。
- 支持粘贴、拖拽和选择附件，并对直接导入附件做数量、单文件大小、总大小和 Windows 文件名安全校验。
- 支持独立生图 `OPENAI_IMAGE_API_KEY` / `OPENAI_IMAGE_BASE_URL`，以及独立识图 `OPENAI_VISION_API_KEY` / `OPENAI_VISION_BASE_URL` / `OPENAI_VISION_MODEL`。
- Windows 打包可携带 `runtime/win/python` 或历史兼容目录 `runtime/win/pyhton`，为无系统 Python 的用户提供 `exec_bash` Python 环境。

## 技术栈

- 桌面端：Electron、electron-vite、electron-builder
- 前端：React 19、TypeScript、Redux Toolkit、React Redux
- UI：Radix UI、lucide-react、Incremark
- 模型与生图：OpenAI JavaScript SDK
- 文档与预览：docx-templates、PizZip、mammoth、pdf-parse、PDF.js、docx-preview
- 测试：Vitest

## 目录结构

```text
src/
  main/        Electron Main、Agent Runtime、工具注册、文件与运行时能力
  preload/     安全 IPC Bridge
  renderer/    React UI、Redux Store、聊天与预览组件
  shared/      前后端共享类型和限制
tests/
  main/        Main 进程能力单测
  renderer/    Renderer 侧纯函数与组件辅助单测
docs/          方案模板和设计文档
runtime/       可选内置运行时资源，例如 Windows Python
```

## 环境要求

- Node.js 20+ 推荐
- npm
- Windows 桌面环境用于完整打包验收
- 可选：LibreOffice，用于 Word 转 PDF
- 可选：Windows Python runtime，放入 `runtime/win/python` 或 `runtime/win/pyhton`

## 快速开始

```bash
npm install
copy .env.example .env.local
npm run dev
```

开发态启动后，渲染端通过 Electron preload 暴露的 `window.pwdSafeAgent` 与 Main 进程通信。不要直接用浏览器打开渲染页，否则无法使用 IPC、文件和 Agent 能力。

## 配置说明

`.env.local` 会优先于 `.env` 加载。也可以在应用内“设置”面板保存配置，配置会写入本地 `.env.local`。

打包时会把项目根目录的 `.env` 一起复制到安装包 `resources/.env`，用于给终端用户提供默认模型配置。请注意：如果 `.env` 内包含真实 API Key，安装包接收者可以提取并使用这些密钥，建议使用单独创建的低额度 Key。

```env
OPENAI_API_KEY=
OPENAI_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1
OPENAI_IMAGE_BASE_URL=
OPENAI_VISION_BASE_URL=
OPENAI_IMAGE_API_KEY=
OPENAI_VISION_API_KEY=
OPENAI_CHAT_MODEL=gpt-5.5
OPENAI_CHAT_IMAGE_INPUT_ENABLED=false
OPENAI_IMAGE_MODEL=gpt-image-2
OPENAI_VISION_MODEL=
OPENAI_IMAGE_SIZE=1536x1024
OPENAI_IMAGE_QUALITY=high
OPENAI_AUTO_IMAGE_GENERATION=true
OPENAI_REQUEST_TIMEOUT_MS=120000
OPENAI_IMAGE_REQUEST_TIMEOUT_MS=300000
OPENAI_MAX_OUTPUT_TOKENS=16000
AGENT_EXEC_BASH_ENABLED=true
AGENT_DRAFT_SECTION_PARALLELISM=20
AGENT_IMAGE_GENERATION_PARALLELISM=10
AGENT_AUTO_PDF_EXPORT=false
LIBREOFFICE_PATH=
```

常用配置：

- `OPENAI_API_KEY`：文本模型 API Key。
- `OPENAI_BASE_URL`：OpenAI 兼容接口地址。
- `OPENAI_IMAGE_API_KEY`：生图 API Key；留空时沿用 `OPENAI_API_KEY`。
- `OPENAI_IMAGE_BASE_URL`：生图接口地址；留空时沿用 `OPENAI_BASE_URL`。
- `OPENAI_CHAT_IMAGE_INPUT_ENABLED`：Chat 模型是否支持直接输入图片；为 `true` 时图片附件会随对话直接传入模型。
- `OPENAI_VISION_API_KEY`：识图工具 API Key；留空时沿用 `OPENAI_API_KEY`。
- `OPENAI_VISION_BASE_URL`：识图工具接口地址；留空时沿用 `OPENAI_BASE_URL`。
- `OPENAI_VISION_MODEL`：识图工具模型；留空时沿用 `OPENAI_CHAT_MODEL`。
- `OPENAI_IMAGE_REQUEST_TIMEOUT_MS`：单次生图请求超时，默认 `300000` ms（5 分钟）。
- `AGENT_DRAFT_SECTION_PARALLELISM`：同时起草的方案章节数，默认 `20`。
- `AGENT_IMAGE_GENERATION_PARALLELISM`：同时运行的生图任务数，默认 `10`，可在设置页调低。
- `AGENT_EXEC_BASH_ENABLED`：是否允许 Agent 使用命令执行工具，默认启用。
- `AGENT_AUTO_PDF_EXPORT`：生成 Word 后是否自动尝试导出 PDF。
- `LIBREOFFICE_PATH`：LibreOffice 可执行文件路径；留空时自动查找。

## 使用方式

1. 启动应用后，在左侧新建或选择对话。
2. 在输入区描述系统背景、建设单位、网络拓扑、等保级别、关键数据和密码应用需求。
3. 可通过选择、粘贴或拖拽上传 Word、PDF、图片和文本资料。
4. 按 `Enter` 发送消息，按 `Shift+Enter` 换行。
5. Agent 会按任务需要自主调用工具，不会在普通问答阶段默认读取模板或生成图片。
6. 生成整篇方案时，Agent 会走“文档配置 → 章节索引 → 章节 md 草稿 → 章节审查/修订/润色 → 合并 md 终稿 → 证据复核 → 图表规划 → Word 写入/PDF”的 section-first 流程。
7. 当 Agent 生成文件并调用 `send_file` 后，聊天流中会出现文件卡片，可预览、系统打开或在文件夹中定位。

## 可用脚本

```bash
npm run dev          # 开发模式启动 Electron + Vite
npm run build        # 类型检查并构建 Main / Preload / Renderer
npm run pack         # 构建并生成 win-unpacked 目录
npm run dist         # 构建安装包
npm run preview      # 预览构建产物
npm run test         # 运行全部 Vitest 单测
npm run test:watch   # 监听模式运行测试
npm run typecheck    # TypeScript 类型检查
npm run icon:generate # 生成应用图标
```

## 打包与资源

`electron-builder` 会把以下目录作为 `extraResources` 复制到安装包资源目录：

- `runtime/`：内置运行时，例如 Windows Python。
- `docs/`：方案模板和项目文档。
- `build/icon.png`、`build/icon.ico`：应用图标。

Windows 安装包目标包括 `nsis` 和 `zip`。开发态数据和生成文件位于项目本地数据目录；打包态会使用 Electron `userData` 目录存储可写数据。

## 测试

单测集中放在 `tests/`：

- `tests/main/`：Agent Runtime、工具注册、文档导出、运行时诊断、会话持久化、安全头等。
- `tests/renderer/`：Redux Store、Bridge、流排序、附件 payload、预览数据转换等。

推荐提交前运行：

```bash
npm run typecheck
npm test
npm run build
```

## 安全说明

- Renderer 不直接读取 API Key，密钥只由 Electron Main 进程加载。
- IPC 通过 preload bridge 暴露有限 API，Renderer 不启用 Node 集成。
- 生产态 CSP 保持严格策略；开发态仅为 Vite React Refresh 放行必要能力。
- `exec_bash` 默认启用以满足本地工具能力，但仍建议只在可信任务和可观察工具调用过程下使用。
- 工具文件访问限制在允许目录和用户附件范围内，避免任意路径读取。

## 文档

- [技术方案](docs/密码方案生成Agent技术方案.md)
- [接口文档](docs/密码方案生成Agent-接口文档.md)
- [UI 设计文档](docs/密码方案生成Agent-UI设计文档.md)
- [密码应用方案模板](docs/密码应用方案.docx)
