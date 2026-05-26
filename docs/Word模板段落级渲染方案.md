# Word 模板不可见锚与段落级渲染方案

## 1. 背景

《密码应用方案》属于复杂格式报告，Word 模板中已经包含章节标题、编号、正文样式、表格样式、页眉页脚、目录、图注和页面设置。文档生成链路的目标不是重新排版一份 Word，而是在权威模板上像人工填写一样写入内容。

现有基于关键词或可见标记的替换方式存在稳定性问题：

1. `${字段}`、`[[PS:field:字段]]` 这类旧式字段会增加兼容分支，后续不再保留。
2. `[[PS:sec/table/fig...]]` 可见结构锚会污染模板，也容易被人工编辑误删、移动或残留到最终文档。
3. 只靠章节标题、编号或关键词搜索定位不可靠。标题可能改名，编号可能重排，同名章节可能重复，Word 也可能把文本拆成多个 run。
4. `docx-templates` 的普通字段替换发生在命令所在的 `w:r` 内，长文本会继承占位符局部格式，容易导致字体、加粗、缩进、换行混乱。

因此，本方案将定位和渲染分开：**定位靠不可见结构锚，格式靠原始 docx 模板，内容写入靠段落级/块级克隆渲染**。

## 2. 核心目标

1. `docs/密码应用方案.docx` 是唯一权威格式源。
2. 生成后的 Word 在样式体系、段落格式、表格格式、标题编号、页眉页脚和整体版式上与原模板保持一致。
3. `docs/密码应用方案.template.json` 只作为结构地图，不作为重建 Word 的模板本体。
4. 章节、表格、图片槽位通过不可见结构锚精确定位，不依赖标题编号搜索。
5. 标题和编号只用于人工可读语义、模板校验和异常报告。
6. 简单字段统一为 `{字段名}` 或内部字段 ID，不再保留 `${字段}`、`[[PS:field:字段]]` 兼容。
7. 可见 `[[PS:sec/table/fig...]]` 结构锚下线。

## 3. 核心结论

推荐最终路线：

```text
原始 docx 模板 = 唯一格式源
不可见锚 = 稳定定位点
template.json = 结构地图和渲染说明
代码渲染器 = 克隆模板段落/表格/图片槽位并替换内容
docx-templates = 可选辅助，只处理短字段或简单命令
```

不要把 `template.json` 做成“重新生成 Word 的模板”。如果从 JSON 重新生成 Word，样式 ID、编号定义、表格边框、段落缩进、行距、页眉页脚、目录域、图片关系等细节很容易丢失，难以做到和原模板一致。

真正稳定的做法是：**解包原 docx，在原 document.xml 里找到不可见锚，然后复制原模板里的样板段落、样板表格行、图片槽位，只替换内容，不重建格式**。

## 4. 不可见锚设计

### 4.1 首选：Content Control / SDT

Word 内容控件对应 OOXML 的 `w:sdt`，可以携带不可见的结构 tag。用户正常编辑 Word 时不会看到 `[[PS:...]]` 文本，但程序可以稳定定位。

章节正文锚示例：

```xml
<w:sdt>
  <w:sdtPr>
    <w:alias w:val="2.2.1 物理环境正文"/>
    <w:tag w:val="ps:section:sec_2_2_1:body"/>
  </w:sdtPr>
  <w:sdtContent>
    <w:p>...</w:p>
    <w:p>...</w:p>
  </w:sdtContent>
</w:sdt>
```

表格锚示例：

```xml
<w:sdt>
  <w:sdtPr>
    <w:alias w:val="服务器清单表"/>
    <w:tag w:val="ps:table:server_list"/>
  </w:sdtPr>
  <w:sdtContent>
    <w:tbl>...</w:tbl>
  </w:sdtContent>
</w:sdt>
```

动态表格行可以使用行级或内容级锚：

```text
ps:row:server_list:item
```

图片槽位可以使用：

```text
ps:figure:network_topology
```

### 4.2 备选：Bookmark

Word bookmark 也是不可见定位点，适合标记起止位置。缺点是人工编辑时较容易破坏起止边界，且嵌套和跨表格场景处理更麻烦。

### 4.3 不推荐：隐藏文本锚

隐藏文本比可见 `[[PS:...]]` 干净，但仍然是正文文本的一部分，可能被查找、复制、清理格式或另存时影响。因此只适合作为兼容兜底，不作为主方案。

## 5. 定位策略

渲染时不搜索章节标题，也不搜索编号，而是直接查找不可见锚：

```text
w:sdtPr / w:tag[@w:val="ps:section:sec_2_2_1:body"]
```

定位顺序：

1. 解包 docx，读取 `word/document.xml`。
2. 遍历 `w:body` 下的块级节点，包括 `w:p`、`w:tbl`、`w:sdt`。
3. 建立 tag 到节点的索引。
4. 根据 `template.json` 中的结构 ID 找到目标锚。
5. 校验锚附近的标题、编号、样式或模板 hash。
6. 替换 `w:sdtContent` 内部内容，保留外层 `w:sdt` 和结构 tag。

标题编号只作为校验：

```text
找到 ps:section:sec_2_2_1:body 后，检查附近标题是否仍然接近 “2.2.1 物理环境”。
```

如果校验失败，输出模板异常报告，不继续盲写。

## 6. 段落级渲染原则

格式保真的关键不是生成新格式，而是克隆原模板格式块。

章节正文渲染规则：

1. 保留章节标题段落，不重建标题。
2. 找到章节正文的 `w:sdtContent`。
3. 选择模板中的样板正文段落作为 `paragraphPrototype`。
4. 每生成一段正文，就复制一个样板段落。
5. 只替换样板段落中的 `w:t` 文本。
6. 保留 `w:pPr`、`w:rPr`、样式 ID、缩进、行距、编号属性等格式。

表格渲染规则：

1. 固定结构表格保留原 `w:tbl`，只替换单元格文本。
2. 动态行表格复制样板行 `w:tr`。
3. 保留单元格边框、底纹、宽度、合并、段落样式。
4. 只替换单元格中的文本或图片关系。

图片渲染规则：

1. 优先使用模板中的图片槽位。
2. 替换 image part 和 relationship，不重新创建整套版式。
3. 图注段落从模板图注样板克隆。

## 7. template.json 的角色

`template.json` 是结构地图，不是 Word 格式源。

推荐记录：

```ts
interface TemplateAnchor {
  id: string;
  type: "section" | "table" | "row" | "figure" | "field";
  tag: string;
  title?: string;
  number?: string;
  expectedStyleId?: string;
  prototype?: {
    paragraphTag?: string;
    rowTag?: string;
    captionTag?: string;
  };
  validation?: {
    nearbyHeadingText?: string;
    documentSha256?: string;
    contentHash?: string;
  };
}
```

章节示例：

```json
{
  "id": "sec_2_2_1",
  "type": "section",
  "tag": "ps:section:sec_2_2_1:body",
  "number": "2.2.1",
  "title": "物理环境",
  "prototype": {
    "paragraphTag": "ps:prototype:paragraph:normal"
  }
}
```

表格示例：

```json
{
  "id": "server_list",
  "type": "table",
  "tag": "ps:table:server_list",
  "prototype": {
    "rowTag": "ps:row:server_list:item"
  }
}
```

## 8. 与 docx-templates 的边界

`docx-templates` 可以继续使用，但不作为核心定位机制。

适合使用 `docx-templates` 的场景：

1. 项目名称、建设单位、地址、等级保护级别等短字段。
2. 简单条件显示。
3. 简单列表或低风险段落。

不建议交给 `docx-templates` 的场景：

1. 按章节标题编号定位。
2. 替换整章正文。
3. 复杂表格动态行。
4. 图片和图注版式保真。
5. 已渲染文档的增量章节替换。

如果继续使用 `docx-templates`，命令也应放在独立段落或独立单元格内，不把大块正文塞进普通句子中的字段。

## 9. 为什么这种方案更稳

标题编号搜索的问题是语义定位，不是结构定位。语义会变，结构 ID 不应变。

不可见锚方案的稳定性来自三点：

1. 锚是机器用的结构 ID，不依赖用户可见标题。
2. 锚不污染正文，不会出现在最终报告里。
3. 锚包住真实 Word 块，替换时可以直接操作 `w:sdtContent`。

格式保真来自三点：

1. 原 docx 保留为唯一格式源。
2. 新内容通过复制原模板样板块生成。
3. 渲染器只替换文本、图片关系和必要的行数量，不重新设计样式。

## 10. 渲染流程

推荐流程如下：

1. 模板预处理。
   给章节正文、表格、动态行、图片槽位、样板段落写入不可见 SDT tag。

2. 生成 template.json。
   记录 tag、章节语义、样板块关系、校验信息和模板 hash。

3. Agent 生成结构化内容。
   输出章节 AST、表格数据、图片引用，而不是直接输出 Word。

4. 解包原 docx。
   读取 `document.xml`、styles、numbering、relationships、media。

5. 按不可见锚定位。
   使用 `w:tag` 精确找到章节、表格、图片槽位。

6. 克隆模板块并替换内容。
   段落克隆 `w:p`，表格行克隆 `w:tr`，图片替换 relationship 和 media part。

7. 校验输出。
   检查锚是否完整、标题是否匹配、字段是否已替换、docx 是否可打开。

8. 打包 docx。
   生成最终 Word 文件。

## 11. 迁移计划

第一阶段：清理旧字段。

1. 移除 `${字段}` 兼容。
2. 移除 `[[PS:field:字段]]` 兼容。
3. 简单字段统一为 `{字段名}` 或内部字段 tag。
4. 更新字段替换测试。

第二阶段：引入不可见锚。

1. 编写模板预处理器，给 docx 写入 SDT tag。
2. 为章节正文、表格、动态行、图片槽位、样板段落生成稳定 ID。
3. 更新 `template.json` schema，记录 tag 和 prototype。
4. 保留标题编号字段作为校验信息。

第三阶段：移除可见结构锚。

1. 停止生成 `[[PS:sec/table/fig...]]`。
2. 从模板和 JSON 中逐步移除 visibleMarkers。
3. 更新依赖 visibleMarkers 的测试。

第四阶段：实现块级渲染器。

1. 按 SDT tag 定位 `w:sdtContent`。
2. 克隆样板段落写入章节正文。
3. 克隆样板行写入动态表格。
4. 替换图片关系和图注。

第五阶段：回归验证。

1. 用真实复杂报告样例测试。
2. 对比渲染前后样式 ID、编号属性、表格属性是否保留。
3. 检查最终 docx 在 Word/WPS 中打开效果。
4. 对模板锚缺失、重复、错位输出明确错误报告。

## 12. 稳定性边界

可以稳定保证：

1. 样式体系来自原模板。
2. 标题、页眉页脚、节属性、目录域保留。
3. 正文段落和表格行继承原模板格式。
4. 机器定位不依赖标题编号搜索。
5. 不可见锚不污染最终正文。

不能绝对保证：

1. 内容变长后分页位置完全不变。
2. Word 和 WPS 在像素级渲染完全一致。
3. 用户手工删除 SDT tag 后仍能无损定位。
4. 修订模式、批注、复杂域代码在所有场景下都不受影响。

因此，本方案的目标是：**格式体系和排版规则与原模板一致，内容长度导致的自然换页允许变化**。

## 13. 最终效果

采用不可见锚后，文档生成链路变为：

```text
机器按不可见锚定位
代码按模板样板块克隆
Agent 只提供结构化内容
标题编号只负责人工可读和校验
```

这样可以实现复杂方案报告格式保留，避免生成文本破坏 Word 原格式，同时去掉旧字段兼容和可见结构锚带来的不稳定因素。
