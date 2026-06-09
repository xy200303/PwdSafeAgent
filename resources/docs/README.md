# 运行时文档资源

本目录存放会随安装包进入生产环境的只读文档资源。

- 源码位置：`resources/docs/`
- 打包后位置：`<resourcesPath>/docs/`
- 运行时对 Agent 暴露的逻辑路径：`docs/...`

## 目录结构

- `templates/`
  - 内置 Word 模板；Word 写入工具会优先从 .docx 动态解析不可见锚点
- `references/`
  - 标准规范、外部参考资料
- `document-profiles/`
  - 基于 Word 模板生成的文档生成规则、行文规则和 Word 渲染方式 profile，可通过 `build_document_config.profile` 使用
- `examples/`
  - 示例输出文档

## 当前关键文件

- Word 模板：`docs/templates/密码应用方案.docx`
- 模板结构 JSON：模板索引文件；未提供时从 `.docx` 动态解析锚点
- 密码应用方案模板 profile：`docs/document-profiles/generic_document.json`
- 用户上传的 Word 模板不会写入本目录；运行时通过设置页上传入口注册到全局 `data/output/document-templates/<id>/`，按标题结构写入隐藏章节锚点，并生成同目录的 `template.json` 与可编辑 `profile.json`
