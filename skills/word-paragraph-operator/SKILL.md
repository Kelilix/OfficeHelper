---
name: word-paragraph-operator
description: Word 文档段落操作技能。当用户需要对 Word 文档中的段落进行读取、增删、合并、拆分、属性设置（对齐/行距/缩进/间距/样式）、编号列表、边框底纹、批量处理、结构分析等任何段落级操作时触发。
---

# Word 段落操作技能

## 核心概念

### ⚡ 选中内容的两步处理原则

当用户请求中同时涉及整体段落和选中部分时，请遵循 word-text-operator 中的两步处理原则。

本技能专注于段落级操作，若用户同时要求对选中文字施加字符格式（如颜色、加粗），需要与 word-text-operator 配合，拆为多个 step。

## 本技能与 word-text-operator 的关系

- `word-text-operator`：以 Range/Selection 为核心，操作文本内容、字符格式
- `word-paragraph-operator`：以 Paragraphs 集合为核心，操作段落本身

| 维度 | word-text-operator | word-paragraph-operator |
|---|---|---|
| 入口对象 | Range / Selection | Paragraphs / Paragraph |
| 典型场景 | 查找替换、字体格式、书签导航 | 按索引访问、段落 CRUD、批量格式 |
| 段落格式 | 需先展开 Range | 直接操作段落 |

## 执行器

`ParagraphOperator` 封装了以下功能：

| 属性/方法 | 说明 | 典型场景 |
|---|---|---|
| `count()` | 段落总数 | 了解文档规模 |
| `get(index)` | 按索引获取段落 | 精确定位到第 N 段 |
| `all()` | 所有段落列表 | 遍历处理 |
| `find_by_text(text)` | 按文本内容查找段落 | 找到含有关键词的段落 |
| `get_document_structure()` | 文档结构摘要 | 生成目录或分析大纲 |
| `find_empty_paragraphs()` | 空段落列表 | 清理文档 |
| `find_headings()` | 所有标题段落 | 提取大纲 |
| `apply_format_to_all()` | 批量设置格式 | 全文档统一排版 |

## 可用操作（按分组）

### 段落基础访问

#### 1. get_paragraph_count

返回文档中的段落总数。

```python
{"action": "get_paragraph_count", "params": {}, "description": "获取段落总数"}
```

#### 2. get_paragraph_by_index

按索引获取段落。`index` 从 1 开始，支持负数（-1 = 最后一段）。

```python
{"action": "get_paragraph_by_index", "params": {"index": 3}, "description": "获取第3段"}
{"action": "get_paragraph_by_index", "params": {"index": -1}, "description": "获取最后一段"}
```

#### 3. get_paragraph_text

读取指定段落的纯文本。

```python
{"action": "get_paragraph_text", "params": {"index": 1}, "description": "读取第1段文本"}
```

#### 4. get_paragraph_range

获取指定段落的字符位置（Range），返回 `[start, end]`。

```python
# 单段落：index 为 1-based 段落索引
{"action": "get_paragraph_range", "params": {"index": 1}, "description": "获取第1段字符位置"}
# 多段落：start~end 返回各段字符位置列表
{"action": "get_paragraph_range", "params": {"start": 1, "end": 3}, "description": "获取1~3段字符位置"}
```

---

### 段落识别与结构分析

#### 5. get_document_structure

返回文档中所有段落的结构信息（索引、文本摘要、样式、级别、是否标题/空/列表）。

```python
{"action": "get_document_structure", "params": {}, "description": "获取文档结构"}
```

#### 6. get_outline_summary

返回文档的大纲结构摘要（用于生成目录）。

```python
{"action": "get_outline_summary", "params": {}, "description": "获取大纲摘要"}
```

#### 7. find_empty_paragraphs

返回所有空段落（仅含段落标记）。

```python
{"action": "find_empty_paragraphs", "params": {}, "description": "查找空段落"}
```

#### 8. find_heading_paragraphs

返回所有标题段落。

```python
{"action": "find_heading_paragraphs", "params": {}, "description": "查找标题段落"}
```

#### 9. find_paragraphs_by_level

返回指定大纲级别的标题段落。`level` 为 1~9。

```python
{"action": "find_paragraphs_by_level", "params": {"level": 1}, "description": "查找1级标题"}
```

#### 10. find_paragraphs_by_text

按文本内容查找所有匹配的段落。

```python
{"action": "find_paragraphs_by_text", "params": {"text": "关键词", "whole_word": false}, "description": "查找含关键词的段落"}
```

---

### 段落属性读取

#### 11. get_paragraph_format_info

读取段落的完整格式属性（对齐、行距、缩进、间距等）。

```python
{"action": "get_paragraph_format_info", "params": {"index": 1}, "description": "读取段落格式"}
```

#### 12. get_paragraph_style

返回段落使用的样式名称。

```python
{"action": "get_paragraph_style", "params": {"index": 1}, "description": "获取样式名称"}
```

#### 13. get_outline_level

返回段落的大纲级别（0=正文，1~9=标题级别）。

```python
{"action": "get_outline_level", "params": {"index": 1}, "description": "获取大纲级别"}
```

#### 14. is_paragraph_list_item

判断段落是否属于编号列表。

```python
{"action": "is_paragraph_list_item", "params": {"index": 1}, "description": "是否编号列表项"}
```

#### 15. is_paragraph_in_table

判断段落是否在表格内。

```python
{"action": "is_paragraph_in_table", "params": {"index": 1}, "description": "是否在表格中"}
```

---

### 段落属性写入（单个）

#### 16. set_paragraph_alignment

设置段落对齐方式。`align`：`"left"` `"center"` `"right"` `"justify"` `"distribute"`

```python
{"action": "set_paragraph_alignment", "params": {"index": 1, "alignment": "center"}, "description": "居中对齐"}
```

#### 17. set_paragraph_line_spacing

设置段落行间距。`rule`：`"single"` `"1.5"` `"double"` `"at_least"` `"exact"`

```python
{"action": "set_paragraph_line_spacing", "params": {"index": 1, "spacing": 1.5, "rule": "1.5"}, "description": "1.5倍行距"}
{"action": "set_paragraph_line_spacing", "params": {"index": 1, "spacing": 20, "rule": "exact"}, "description": "固定值20磅"}
```

#### 18. set_paragraph_space_before

设置段前间距（磅值）。

```python
{"action": "set_paragraph_space_before", "params": {"index": 1, "points": 12.0}, "description": "段前12磅"}
```

#### 19. set_paragraph_space_after

设置段后间距（磅值）。

```python
{"action": "set_paragraph_space_after", "params": {"index": 1, "points": 6.0}, "description": "段后6磅"}
```

#### 20. set_paragraph_indent_left

左缩进。`characters`（字符数）或 `cm`（厘米），二选一。

```python
{"action": "set_paragraph_indent_left", "params": {"index": 1, "cm": 2.0}, "description": "左缩进2厘米"}
```

#### 21. set_paragraph_indent_right

右缩进。

```python
{"action": "set_paragraph_indent_right", "params": {"index": 1, "cm": 1.0}, "description": "右缩进1厘米"}
```

#### 22. set_paragraph_first_line_indent

首行缩进。传负值则为悬挂缩进。**推荐用 `characters` 参数（字符数），而非 `cm`**。

```python
# 首行缩进 2 字符（约 0.74cm）
{"action": "set_paragraph_first_line_indent", "params": {"index": 1, "characters": 2}, "description": "首行缩进2字符"}
# 悬挂缩进 2 字符
{"action": "set_paragraph_first_line_indent", "params": {"index": 1, "characters": -2}, "description": "悬挂缩进2字符"}
# 用厘米指定
{"action": "set_paragraph_first_line_indent", "params": {"index": 1, "cm": 0.74}, "description": "首行缩进0.74厘米"}
```

**注意**：`index` 只设置单个段落。要对所有段落设置时，请用下面的 `apply_format_to_range`。

#### 23. set_paragraph_hanging_indent

设置悬挂缩进。

```python
{"action": "set_paragraph_hanging_indent", "params": {"index": 1, "characters": 2}, "description": "悬挂缩进2字符"}
```

#### 24. set_paragraph_outline_level

设置大纲级别。

```python
{"action": "set_paragraph_outline_level", "params": {"index": 1, "level": 1}, "description": "设为1级大纲"}
```

#### 25. set_paragraph_keep_together

段内不分页。

```python
{"action": "set_paragraph_keep_together", "params": {"index": 1, "on": true}, "description": "段内不分页"}
```

#### 26. set_paragraph_keep_with_next

与下段同页。

```python
{"action": "set_paragraph_keep_with_next", "params": {"index": 1, "on": true}, "description": "与下段同页"}
```

#### 27. set_paragraph_style

应用样式到段落。

```python
{"action": "set_paragraph_style", "params": {"index": 1, "style_name": "标题 1"}, "description": "应用标题1样式"}
```

#### 28. reset_paragraph_format

将段落格式恢复为默认样式。

```python
{"action": "reset_paragraph_format", "params": {"index": 1}, "description": "重置段落格式"}
```

---

### 段落边框与底纹

#### 29. set_paragraph_border

给段落添加边框。`side`：`"top"` `"bottom"` `"left"` `"right"`

```python
{"action": "set_paragraph_border", "params": {"index": 1, "side": "bottom", "line_style": 1, "line_width": 4, "color": "black"}, "description": "添加下边框"}
```

#### 30. clear_paragraph_border

清除段落边框。

```python
{"action": "clear_paragraph_border", "params": {"index": 1}, "description": "清除边框"}
```

#### 31. set_paragraph_shading

设置段落底纹（背景填充色）。

```python
{"action": "set_paragraph_shading", "params": {"index": 1, "fill_color": "yellow"}, "description": "黄色背景"}
```

#### 32. clear_paragraph_shading

清除段落底纹。

```python
{"action": "clear_paragraph_shading", "params": {"index": 1}, "description": "清除底纹"}
```

---

### 编号列表操作

#### 33. get_list_paragraphs

返回所有编号/项目符号列表段落。

```python
{"action": "get_list_paragraphs", "params": {}, "description": "获取所有编号段落"}
```

#### 34. get_list_level

返回段落在列表中的级别（1-based）。

```python
{"action": "get_list_level", "params": {"index": 1}, "description": "获取列表级别"}
```

#### 35. set_list_level

设置段落在列表中的级别。

```python
{"action": "set_list_level", "params": {"index": 1, "level": 2}, "description": "设为2级列表"}
```

#### 36. apply_bullet_list

将段落转换为项目符号列表。

```python
{"action": "apply_bullet_list", "params": {"index": 1}, "description": "转为项目符号"}
```

#### 37. apply_numbered_list

将段落转换为编号列表。

```python
{"action": "apply_numbered_list", "params": {"index": 1, "number_format": "decimal", "start_at": 1}, "description": "转为编号列表"}
```

#### 38. remove_list_format

移除段落的列表格式（还原为普通段落）。

```python
{"action": "remove_list_format", "params": {"index": 1}, "description": "移除列表格式"}
```

---

### 段落内容操作（CRUD）

#### 39. set_paragraph_text

替换段落内容（保留段落格式）。

```python
{"action": "set_paragraph_text", "params": {"index": 1, "text": "新文本内容"}, "description": "替换段落文本"}
```

#### 40. insert_text_before_paragraph

在段落开头插入文本。

```python
{"action": "insert_text_before_paragraph", "params": {"index": 1, "text": "前缀文本"}, "description": "段落前插入文本"}
```

#### 41. insert_text_after_paragraph

在段落末尾插入文本。

```python
{"action": "insert_text_after_paragraph", "params": {"index": 1, "text": "后缀文本"}, "description": "段落后插入文本"}
```

#### 42. delete_paragraph

删除整个段落（慎用：会合并相邻段落）。

```python
{"action": "delete_paragraph", "params": {"index": 1}, "description": "删除段落"}
```

#### 43. clear_paragraph

清空段落内容（保留段落标记）。

```python
{"action": "clear_paragraph", "params": {"index": 1}, "description": "清空段落内容"}
```

#### 44. add_paragraph_after

在指定段落之后插入新段落。

```python
{"action": "add_paragraph_after", "params": {"index": 1, "text": "新段落内容"}, "description": "在段落后插入新段落"}
```

#### 45. add_paragraph_before

在指定段落之前插入新段落。

```python
{"action": "add_paragraph_before", "params": {"index": 1, "text": "新段落内容"}, "description": "在段落前插入新段落"}
```

#### 46. merge_with_next

将当前段落与下一段合并。

```python
{"action": "merge_with_next", "params": {"index": 1}, "description": "与下一段合并"}
```

#### 47. merge_with_previous

将当前段落与上一段合并。

```python
{"action": "merge_with_previous", "params": {"index": 2}, "description": "与上一段合并"}
```

#### 48. split_paragraph_by_separator

按分隔符将一个段落拆分为多个段落。默认分隔符为 Tab（`\t`）。

```python
{"action": "split_paragraph_by_separator", "params": {"index": 1, "separator": "\t"}, "description": "按Tab拆分段落"}
{"action": "split_paragraph_by_separator", "params": {"index": 1, "separator": ","}, "description": "按逗号拆分段落"}
```

---

### 批量操作

#### 49. apply_format_to_range

批量设置 [start, end] 范围内所有段落的格式。**适合"全文所有段落统一设置"的场景**。返回影响段落数量。

`start` / `end` 为 1-based 闭区间；**`end` 可为 `-1` 表示最后一段**（与单段 `index: -1` 约定一致）。勿把 `-1` 当作「全文」以外的含义误用——`start:1, end:-1` 即第 1 段到最后一段。

```python
# 全文首行缩进 2 字符（推荐：end 用段落总数，或 end=-1 表示到最后一段）
{"action": "apply_format_to_range", "params": {"start": 1, "end": -1, "first_line_indent_characters": 2}, "description": "全文首行缩进"}
{"action": "apply_format_to_range", "params": {"start": 1, "end": 10, "first_line_indent_characters": 2}, "description": "批量设置首行缩进"}
# 对齐+行距+首行缩进
{"action": "apply_format_to_range", "params": {"start": 1, "end": 10, "alignment": "justify", "spacing": 1.5, "rule": "1.5", "first_line_indent_characters": 2}, "description": "批量设置格式"}
```

#### 50. reverse_paragraph_order

将指定范围内的段落顺序反转（原地交换文本内容）。

```python
{"action": "reverse_paragraph_order", "params": {"start": 1, "end": 5}, "description": "反转段落顺序"}
```

#### 51. delete_empty_paragraphs

删除所有空段落。返回删除数量。

```python
{"action": "delete_empty_paragraphs", "params": {}, "description": "删除所有空段落"}
```

---

### Selection / Range 互操作

#### 52. get_paragraph_at_selection

获取当前 Selection 所在段落的段落对象。

```python
{"action": "get_paragraph_at_selection", "params": {}, "description": "获取光标所在段落"}
```

#### 53. select_paragraph

选中指定段落（影响界面 Selection）。**不传 `index` 时选中光标所在段落**（与 `move_to_document_start` + `move_down` 等导航配合时必用，勿默认当成第 1 段）。

```python
{"action": "select_paragraph", "params": {"index": 2}, "description": "选中第2段"}
{"action": "select_paragraph", "params": {}, "description": "选中光标所在整段"}
```

#### 54. select_paragraph_range

选中指定范围内的所有段落。

```python
{"action": "select_paragraph_range", "params": {"start": 1, "end": 5}, "description": "选中1~5段"}
```

---

## 执行流程

仔细阅读用户需求，决定需要调用哪些 action。只返回 JSON 数组：

```json
[
  {"action": "get_paragraph_count", "params": {}, "description": "获取段落总数"},
  {"action": "apply_format_to_range", "params": {"start": 1, "end": 5}, "description": "批量设置格式"}
]
```

**典型场景「所有段落设置首行缩进 2 字符」**：

```json
[
  {"action": "get_paragraph_count", "params": {}, "description": "获取段落总数"},
  {"action": "apply_format_to_range", "params": {"start": 1, "end": 段落总数, "first_line_indent_characters": 2}, "description": "全文首行缩进2字符"}
]
```

`end` 填入第一步获取到的段落数量。如果文档只有少量已知段落，可直接写具体数字（如 `end: 10`）。

如果只需要一个 action，也返回只含一个元素的数组。

## 典型案例库

以下案例展示常见需求的正确 action 组合方式。遇到类似需求时，可参考对应案例的行动序列。

---

暂无

---

## 注意事项

- 段落索引从 1 开始（Word 原生），-1 表示最后一段
- `delete_paragraph` 会删除段落标记，导致相邻段落合并，慎用
- `clear_paragraph` 只清空文本内容，保留段落标记
- 详细 API 参考见 [references/API_REFERENCE.md](references/API_REFERENCE.md)
