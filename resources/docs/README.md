# 运行时文档资源

本目录存放会随安装包进入生产环境的只读文档资源。

- 源码位置：`resources/docs/`
- 打包后位置：`<resourcesPath>/docs/`
- 运行时对 Agent 暴露的逻辑路径：`docs/...`

## 目录结构

- `templates/`
  - 内置 Word 模板与模板标注 JSON
- `references/`
  - 标准规范、外部参考资料
- `examples/`
  - 示例输出文档

## 当前关键文件

- Word 模板：`docs/templates/密码应用方案.docx`
- 模板标注 JSON：`docs/templates/密码应用方案.template.json`
