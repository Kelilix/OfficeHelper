---
name: word-page-operator
description: Word 文档页面设置技能。当用户需要对 Word 文档的页面布局进行操作时触发，包括页边距、纸张大小、纸张方向、分栏、页眉页脚、分页符、节属性等设置。**本技能也负责分节符的插入与删除操作。**
---

# Word 页面设置技能

## 核心概念

### ⚡ 页面设置的两大入口：PageSetup 与 Sections

Word 文档的页面设置通过两个核心对象完成：

- **PageSetup**：控制单个节的页面布局（页边距、纸张、方向、分栏、水印等）
- **Sections / SectionFormat**：控制文档节的数量、节属性和节类型

每个 Word 文档至少有一个节（Section 1）。通过插入分节符可以将文档分为多个节，每个节拥有独立的页面设置。

本技能封装了 `Document.Sections` 及其 `PageSetup` 属性，提供完整的页面操作能力。

### 本技能与 word-paragraph-operator / word-text-operator 的关系

| 维度 | word-text-operator | word-paragraph-operator | word-page-operator |
|---|---|---|---|
| 入口对象 | Range / Selection | Paragraphs / Paragraph | Sections / PageSetup |
| 典型场景 | 文本内容、字体格式 | 段落格式、对齐缩进 | 页边距、纸张方向、分栏 |
| 作用层次 | 字符级 | 段落级 | 页面级 / 节级 |

## 执行器

`PageOperator` 封装了以下功能：

| 属性/方法 | 说明 | 典型场景 |
|---|---|---|
| `page` | `PageSetupOperator` | 页边距、纸张、方向、水印 |
| `section` | `SectionOperator` | 分节、节格式、节起始 |
| `count()` | 节总数 | 了解文档分节情况 |
| `get(index)` | 按索引获取节 | 精确定位到第 N 节 |
| `all()` | 所有节列表 | 遍历处理多节文档 |

## 可用操作（按分组）

### 节基础访问

#### 1. get_section_count

返回文档中的节总数。

```python
{"action": "get_section_count", "params": {}, "description": "获取节总数"}
```

#### 2. get_section_by_index

按索引获取节。`index` 从 1 开始，支持负数（-1 = 最后一节）。

```python
{"action": "get_section_by_index", "params": {"index": 1}, "description": "获取第1节"}
{"action": "get_section_by_index", "params": {"index": -1}, "description": "获取最后一节"}
```

#### 3. get_current_section_index

获取当前 Selection（光标）所在节的索引。

```python
{"action": "get_current_section_index", "params": {}, "description": "获取光标所在节"}
```

---

### 页面设置读取

#### 4. get_page_setup_info

读取指定节的完整页面设置信息（页边距、纸张、方向、分栏等）。

```python
{"action": "get_page_setup_info", "params": {"index": 1}, "description": "读取页面设置"}
```

#### 5. get_page_margins

读取页边距值（厘米）。

```python
{"action": "get_page_margins", "params": {"index": 1}, "description": "读取页边距"}
```

#### 6. get_paper_size

读取纸张大小（宽高，厘米）。

```python
{"action": "get_paper_size", "params": {"index": 1}, "description": "读取纸张大小"}
```

#### 7. get_orientation

读取纸张方向：`"portrait"`（纵向）或 `"landscape"`（横向）。

```python
{"action": "get_orientation", "params": {"index": 1}, "description": "读取纸张方向"}
```

#### 8. get_column_count

读取分栏数。

```python
{"action": "get_column_count", "params": {"index": 1}, "description": "读取分栏数"}
```

#### 9. get_column_info

读取指定节的分栏详细信息（栏数、栏宽、栏间距、是否等宽栏）。

```python
{"action": "get_column_info", "params": {"index": 1}, "description": "读取分栏详情"}
```

#### 10. get_section_start_type

读取节的起始类型：`"continuous"` `"new_page"` `"even_page"` `"odd_page"`。

```python
{"action": "get_section_start_type", "params": {"index": 1}, "description": "读取节起始类型"}
```

---

### 页边距设置

#### 11. set_page_margins

设置页边距。参数为厘米值，可传部分参数（其他保持不变）。

```python
# 全部设置
{"action": "set_page_margins", "params": {"index": 1, "top": 2.54, "bottom": 2.54, "left": 3.17, "right": 3.17}, "description": "设置页边距"}
# 仅设置上边距
{"action": "set_page_margins", "params": {"index": 1, "top": 3.0}, "description": "设置上边距"}
```

#### 12. set_page_margins_by_inch

用英寸设置页边距（方便习惯英制单位的场景）。

```python
{"action": "set_page_margins_by_inch", "params": {"index": 1, "top": 1.0, "bottom": 1.0, "left": 1.25, "right": 1.25}, "description": "英寸设置页边距"}
```

#### 13. set_page_margins_preset

使用预设方案设置页边距。

```python
{"action": "set_page_margins_preset", "params": {"index": 1, "preset": "normal"}, "description": "普通边距"}
{"action": "set_page_margins_preset", "params": {"index": 1, "preset": "narrow"}, "description": "窄边距"}
{"action": "set_page_margins_preset", "params": {"index": 1, "preset": "wide"}, "description": "宽边距"}
{"action": "set_page_margins_preset", "params": {"index": 1, "preset": "mirrored"}, "description": "对称边距"}
```

预设值说明：
- `"normal"`：上下 2.54cm，左右 3.17cm（Word 默认）
- `"narrow"`：上下左右均为 1.27cm
- `"wide"`：上下 2.54cm，左右 5.08cm
- `"mirrored"`：内外 2.54cm，左 3.17cm，右 3.17cm（用于双面打印）

---

### 纸张设置

#### 14. set_paper_size

设置纸张大小。参数为宽高（厘米）。

```python
{"action": "set_paper_size", "params": {"index": 1, "width": 21.0, "height": 29.7}, "description": "设为A4"}
{"action": "set_paper_size", "params": {"index": 1, "width": 8.5, "height": 11}, "description": "设为Letter"}
```

#### 15. set_paper_size_preset

使用预设纸张类型。

```python
{"action": "set_paper_size_preset", "params": {"index": 1, "preset": "A4"}, "description": "设为A4"}
{"action": "set_paper_size_preset", "params": {"index": 1, "preset": "Letter"}, "description": "设为Letter"}
{"action": "set_paper_size_preset", "params": {"index": 1, "preset": "A3"}, "description": "设为A3"}
{"action": "set_paper_size_preset", "params": {"index": 1, "preset": "A5"}, "description": "设为A5"}
```

常用预设：`A0` `A1` `A2` `A3` `A4` `A5` `B4` `B5` `Letter` `Legal` `Tabloid` `Executive`

#### 16. set_orientation

设置纸张方向。

```python
{"action": "set_orientation", "params": {"index": 1, "orientation": "landscape"}, "description": "设为横向"}
{"action": "set_orientation", "params": {"index": 1, "orientation": "portrait"}, "description": "设为纵向"}
```

---

### 分栏操作

#### 17. set_columns

设置分栏数。

```python
{"action": "set_columns", "params": {"index": 1, "count": 2}, "description": "设为两栏"}
{"action": "set_columns", "params": {"index": 1, "count": 3, "equal_width": true}, "description": "三栏等宽"}
```

#### 18. set_columns_with_gutter

设置分栏数并指定栏间距（厘米）。

```python
{"action": "set_columns_with_gutter", "params": {"index": 1, "count": 2, "spacing": 0.75}, "description": "两栏间距0.75cm"}
```

#### 19. set_columns_equal_width

将分栏设为等宽栏。

```python
{"action": "set_columns_equal_width", "params": {"index": 1}, "description": "设为等宽栏"}
```

#### 20. set_column_width

设置指定栏的宽度（厘米）。`column` 从 1 开始，`width` 为栏宽。

```python
{"action": "set_column_width", "params": {"index": 1, "column": 1, "width": 8.0}, "description": "设置第1栏宽度"}
```

#### 21. apply_two_column_layout

应用两栏布局（带分隔线）。

```python
{"action": "apply_two_column_layout", "params": {"index": 1, "with_line": true}, "description": "两栏带分隔线"}
{"action": "apply_two_column_layout", "params": {"index": 1, "with_line": false}, "description": "两栏无分隔线"}
```

---

### 分节符与节操作

#### 22. insert_section_break

在当前 Selection 处插入分节符。`type`：`"continuous"` `"new_page"` `"even_page"` `"odd_page"`

```python
{"action": "insert_section_break", "params": {"type": "new_page"}, "description": "插入分节符（下一页）"}
{"action": "insert_section_break", "params": {"type": "continuous"}, "description": "插入分节符（连续）"}
```

#### 23. set_section_start_type

设置指定节的起始类型。

```python
{"action": "set_section_start_type", "params": {"index": 1, "type": "new_page"}, "description": "节起始于新页"}
```

#### 24. set_section_start_new_page

快捷方式：将节设为从新页开始。

```python
{"action": "set_section_start_new_page", "params": {"index": 1}, "description": "节从新页开始"}
```

#### 25. set_section_start_continuous

快捷方式：将节设为连续（无分页）。

```python
{"action": "set_section_start_continuous", "params": {"index": 1}, "description": "连续节"}
```

#### 26. set_first_page_different

设置首页不同（首页使用不同的页眉页脚）。

```python
{"action": "set_first_page_different", "params": {"index": 1, "on": true}, "description": "首页不同"}
{"action": "set_first_page_different", "params": {"index": 1, "on": false}, "description": "取消首页不同"}
```

#### 27. set_odd_and_even_pages

设置奇偶页不同（奇偶页使用不同的页眉页脚）。

```python
{"action": "set_odd_and_even_pages", "params": {"index": 1, "on": true}, "description": "奇偶页不同"}
{"action": "set_odd_and_even_pages", "params": {"index": 1, "on": false}, "description": "取消奇偶页不同"}
```

---

### 分页控制

#### 28. insert_page_break

在 Range 处插入手动分页符（硬分页符）。

```python
{"action": "insert_page_break", "params": {"rng": "[100, 100]"}, "description": "插入分页符"}
```

#### 29. insert_line_break

在 Range 处插入换行符（软分页符，跨行不断行）。

```python
{"action": "insert_line_break", "params": {"rng": "[50, 50]"}, "description": "插入换行符"}
```

#### 30. insert_column_break

在多栏文档中插入栏分隔符（跳到下一栏）。

```python
{"action": "insert_column_break", "params": {"rng": "[50, 50]"}, "description": "插入栏分隔符"}
```

#### 31. remove_page_break

移除 Range 处的分页符。

```python
{"action": "remove_page_break", "params": {"rng": "[100, 102]"}, "description": "移除分页符"}
```

#### 32. get_page_count

获取文档的总页数。

```python
{"action": "get_page_count", "params": {}, "description": "获取总页数"}
```

#### 33. get_page_of_range

获取指定 Range 所在页的页码。

```python
{"action": "get_page_of_range", "params": {"rng": "[0, 10]"}, "description": "获取所在页码"}
```

---

### 页眉页脚操作

#### 34. set_header

设置页眉内容。`position`：`"primary"` `"first"` `"even_odd"`；`alignment`：`"left"` `"center"` `"right"`

```python
{"action": "set_header", "params": {"index": 1, "position": "primary", "text": "文档标题"}, "description": "设置页眉"}
{"action": "set_header", "params": {"index": 1, "position": "first", "text": "首页页眉"}, "description": "设置首页页眉"}
{"action": "set_header", "params": {"index": 1, "position": "even_odd", "text": "偶数页页眉"}, "description": "设置偶数页眉"}
```

#### 35. get_header

读取页眉内容。

```python
{"action": "get_header", "params": {"index": 1, "position": "primary"}, "description": "读取页眉"}
```

#### 36. clear_header

清除页眉内容。

```python
{"action": "clear_header", "params": {"index": 1, "position": "primary"}, "description": "清除页眉"}
```

#### 37. set_footer

设置页脚内容。

```python
{"action": "set_footer", "params": {"index": 1, "position": "primary", "text": "第 {PAGE} 页"}, "description": "设置页脚"}
```

#### 38. get_footer

读取页脚内容。

```python
{"action": "get_footer", "params": {"index": 1, "position": "primary"}, "description": "读取页脚"}
```

#### 39. clear_footer

清除页脚内容。

```python
{"action": "clear_footer", "params": {"index": 1, "position": "primary"}, "description": "清除页脚"}
```

#### 40. insert_page_number_in_header

在页眉中插入页码字段。

```python
{"action": "insert_page_number_in_header", "params": {"index": 1, "position": "primary", "alignment": "right"}, "description": "页眉右侧插入页码"}
```

#### 41. insert_page_number_in_footer

在页脚中插入页码字段。

```python
{"action": "insert_page_number_in_footer", "params": {"index": 1, "position": "primary", "alignment": "center"}, "description": "页脚居中插入页码"}
```

---

### 页面级格式设置

#### 42. set_vertical_alignment

设置页面内容的垂直对齐方式。`align`：`"top"` `"center"` `"justify"` `"bottom"`

```python
{"action": "set_vertical_alignment", "params": {"index": 1, "align": "center"}, "description": "内容垂直居中"}
```

#### 43. set_page_border

设置页面边框（整页四周的边框）。

```python
{"action": "set_page_border", "params": {"index": 1, "side": "all", "line_style": 1, "line_width": 6, "color": "black"}, "description": "设置页面边框"}
```

#### 44. clear_page_border

清除页面边框。

```python
{"action": "clear_page_border", "params": {"index": 1}, "description": "清除页面边框"}
```

#### 45. set_page_shading

设置整页背景填充色。

```python
{"action": "set_page_shading", "params": {"index": 1, "fill_color": "lightBlue"}, "description": "设置页面背景色"}
```

#### 46. clear_page_shading

清除页面背景色。

```python
{"action": "clear_page_shading", "params": {"index": 1}, "description": "清除页面背景"}
```

---

### 文档级操作

#### 47. apply_page_setup_to_all

将当前节的页面设置应用到所有节。

```python
{"action": "apply_page_setup_to_all", "params": {"index": 1}, "description": "应用设置到全文"}
```

#### 48. copy_page_setup

将源节的页面设置复制到目标节。

```python
{"action": "copy_page_setup", "params": {"from_index": 1, "to_index": 2}, "description": "复制页面设置"}
```

#### 49. reset_page_setup

重置指定节的页面设置为 Word 默认值。

```python
{"action": "reset_page_setup", "params": {"index": 1}, "description": "重置页面设置"}
```

---

## 执行流程

仔细阅读用户需求，决定需要调用哪些 action。只返回 JSON 数组：

```json
[
  {"action": "get_page_setup_info", "params": {"index": 1}, "description": "读取页面设置"},
  {"action": "set_page_margins", "params": {"index": 1, "top": 2.0, "bottom": 2.0, "left": 3.0, "right": 3.0}, "description": "设置页边距"}
]
```

**典型场景「设置全文页边距为2厘米」**：

```json
[
  {"action": "get_section_count", "params": {}, "description": "获取节总数"},
  {"action": "set_page_margins", "params": {"index": 1, "top": 2.0, "bottom": 2.0, "left": 2.0, "right": 2.0}, "description": "设置页边距"}
]
```

如果文档只有单节，可以直接写 `index: 1`。如果文档有多节，请先调用 `get_section_count` 查询节数量，再决定操作哪个节。

## 典型案例库

以下案例展示常见需求的正确 action 组合方式。遇到类似需求时，可参考对应案例的行动序列。

---

### 案例1 【将第一页设置为A3】

#### 分析

- 用户要求「第一页」设置纸张大小，属于页面级操作，不涉及文本范围
- `set_paper_size_preset` 等页面设置 action **不需要**先调用 `get_full_text` 查范围，直接调用即可
- 节索引 `index: 1` 表示第一节（第一页所在节）

```json
[
  {"action": "set_paper_size_preset", "params": {"index": 1, "preset": "A3"}, "description": "将第一节纸张设为A3"}
]
```

### 案例2 【将全文页面大小设置为A4】

#### 分析

- 用户要求「全文」设置纸张大小，属于页面级操作，但是需要关注文档有几个节，并对所有节都做应用
- `set_paper_size_preset` 
- 节索引 `index: 1` 表示第一节（第一页所在节）

```json
[
  {"action": "set_paper_size_preset", "params": {"index": 1, "preset": "B5"}, "description": "将第一节纸张设为 B5"},
  {"action": "apply_page_setup_to_all", "params": {"index": 1}, "description": "应用页面设置到全文所有节"}
]
```

---

### 案例3 【删除第 2 节前的分节符】

#### 分析

- 用户说「删除第 2 节」，实际是删除第二节**前面**的分节符，使第二节内容合并到第一节
- `index: 1`（第一节）不能删除，因为首节前不存在分节符，调用会返回错误
- 属于**页面级操作**，不需要 `get_full_text` 查范围

```json
[
  {"action": "delete_section_break", "params": {"index": 2}, "description": "删除第二节前的分节符"}
]
```

### 案例4 【删除所有分节符】

#### 分析

- `index: 0` 表示删除所有分节符，从后往前逐个删除（避免索引偏移）
- 若文档只有一节，无分节符可删，操作正常返回

```json
[
  {"action": "delete_section_break", "params": {"index": 0}, "description": "删除所有分节符"}
]
```

---

## 注意事项

- 节索引从 1 开始（Word 原生），-1 表示最后一节，**0 表示所有节**
- **第一节不能删除**，因为首节前不存在分节符，删除会返回错误
- 多节文档中，每个节可以有独立的页面设置，通过 `index` 参数指定
- 删除分节符会将该节内容合并到前一节，同时该节的页面设置丢失
- `rng` 参数支持三种写法：`"full_document"`、`"[start, end]"`、省略（当前选区）
- 详细 API 参考见 [references/API_REFERENCE.md](references/API_REFERENCE.md)
