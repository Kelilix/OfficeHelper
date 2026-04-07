# Word Paragraph Operator API 参考

本文档是 `word-paragraph-operator` 技能的详细 API 参考，补充 SKILL.md 中的速查用法。

## 目录

1. [Word COM 常量](#1-word-com-常量)
2. [ParagraphOperator 类](#2-paragraphoperator-类)
3. [Action 注册表映射](#3-action-注册表映射)

---

## 1. Word COM 常量

本技能使用的 Word 内置常量定义如下：

### 对齐方式（wdParagraphAlignment）

| 常量名 | 值 | 说明 |
|---|---|---|
| `WD_ALIGN_PARAGRAPH_LEFT` | 0 | 左对齐 |
| `WD_ALIGN_PARAGRAPH_CENTER` | 1 | 居中对齐 |
| `WD_ALIGN_PARAGRAPH_RIGHT` | 2 | 右对齐 |
| `WD_ALIGN_PARAGRAPH_JUSTIFY` | 3 | 两端对齐 |
| `WD_ALIGN_PARAGRAPH_DISTRIBUTE` | 4 | 分散对齐 |

### 行距规则（wdLineSpacingType）

| 常量名 | 值 | 说明 |
|---|---|---|
| `WD_LINESPACING_SINGLE` | 0 | 单倍行距 |
| `WD_LINESPACING_ONE_POINT_FIVE` | 1 | 1.5 倍行距 |
| `WD_LINESPACING_DOUBLE` | 2 | 双倍行距 |
| `WD_LINESPACING_AT_LEAST` | 3 | 固定值（最小值） |
| `WD_LINESPACING_EXACT` | 4 | 固定值（精确值） |

### 列表类型（wdListType）

| 常量名 | 值 | 说明 |
|---|---|---|
| `WD_LIST_NONE` | 0 | 非列表 |
| `WD_LIST_BULLET` | 1 | 项目符号 |
| `WD_LIST_NUMBER` | 2 | 编号列表 |
| `WD_LIST_PICTURE_BULLET` | 3 | 图片项目符号 |
| `WD_LIST_MIXED` | -1 | 混合类型 |

### 其他常量

| 常量名 | 值 | 说明 |
|---|---|---|
| `WD_STYLE_NORMAL` | -1 | 正文样式 |
| `WD_STYLE_HEADING_1` | -2 | 标题 1 样式 |
| `WD_STYLE_HEADING_2` | -3 | 标题 2 样式 |
| `WD_STYLE_HEADING_3` | -4 | 标题 3 样式 |
| `WD_MOVE` | 0 | 移动（Collapse 方向） |
| `WD_EXTEND` | 1 | 扩展（Collapse 方向） |

---

## 2. ParagraphOperator 类

### 构造函数

```python
def __init__(self, word_base: WordBase):
    """
    Args:
        word_base: WordBase 实例（由 WordTextOperator 传入）
    """
```

### 内部工具方法

以下方法供其他公共方法内部调用，不建议直接使用：

#### `_paragraphs() -> CDispatch`

返回文档的 Paragraphs 集合对象。

#### `_list_paras() -> CDispatch`

返回文档的 ListParagraphs 集合对象。

#### `_resolve_color(color: Union[str, int]) -> int`

将颜色名（如 `"yellow"` `"black"`）或十六进制字符串（如 `"#FFFF00"`）或整数转为 Word 颜色值。

支持的颜色名：`black` `white` `red` `green` `blue` `yellow` `cyan` `magenta` `gray`

### 基础访问方法

#### `count() -> int`

返回文档中的段落总数。

#### `get(index: int) -> CDispatch`

按索引获取单个段落。Word 段落索引从 1 开始。

- `index=1` → 第一段
- `index=-1` → 最后一段
- `index=-2` → 倒数第二段

#### `all() -> List[CDispatch]`

返回所有段落对象的列表。

#### `first() -> CDispatch`

返回第一段。等价于 `get(1)`。

#### `last() -> CDispatch`

返回最后一段。等价于 `get(-1)`。

#### `range(start: int, end: int) -> List[CDispatch]`

返回 [start, end] 范围内的段落列表（含首尾）。支持负数索引。

```python
paras = op.para.range(1, 5)  # 第1~5段
paras = op.para.range(2, -1)  # 第2段到倒数第1段
```

#### `at_range(rng: CDispatch) -> List[CDispatch]`

返回指定 Range 范围内包含的所有段落。

### 段落属性读取方法

#### `get_text(para: CDispatch) -> str`

读取段落的纯文本内容（不含尾部换行符）。

#### `get_length(para: CDispatch) -> int`

返回段落文本的字符数（不含段落标记）。

#### `get_index(para: CDispatch) -> int`

返回段落在文档中的编号（从 1 开始）。

#### `get_style_name(para: CDispatch) -> str`

返回段落使用的样式本地名称（如 `"正文"` `"标题 1"`）。

#### `get_style_wd_name(para: CDispatch) -> int`

返回段落样式的 Word 内置类型编号。

#### `get_outline_level(para: CDispatch) -> int`

返回段落的大纲级别：
- `0` = 正文
- `1`~`9` = 标题级别（对应 Word 的 1~9 级大纲）

#### `is_heading(para: CDispatch) -> bool`

判断段落是否为标题样式（样式名含 "heading" 或 "标题"）。

#### `is_empty(para: CDispatch) -> bool`

判断段落是否为空（仅含段落标记 `\r` 或 `\x07`）。

#### `is_list_item(para: CDispatch) -> bool`

判断段落是否属于编号列表（ListFormat.ListType != WD_LIST_NONE）。

#### `is_in_table(para: CDispatch) -> bool`

判断段落是否在表格内。通过 `Range.Information(12)` 检测。

#### `get_format_info(para: CDispatch) -> dict`

读取段落的完整格式属性，返回字段：

```python
{
    "alignment": int,           # 对齐常量
    "alignment_name": str,      # 对齐名称
    "line_spacing": float,      # 行距值
    "line_spacing_rule": int,   # 行距规则常量
    "space_before": float,      # 段前间距（磅）
    "space_after": float,       # 段后间距（磅）
    "left_indent": float,       # 左缩进
    "right_indent": float,      # 右缩进
    "first_line_indent": float, # 首行缩进
    "outline_level": int,       # 大纲级别
    "widow_control": int,      # 孤行控制
    "keep_together": int,       # 段内不分页
    "keep_with_next": int,     # 与下段同页
    "page_break_before": int,  # 段前分页
    "style_name": str,         # 样式名
}
```

### 段落属性写入方法

#### `set_alignment(para: CDispatch, align: Union[str, int])`

设置段落对齐方式。

#### `set_line_spacing(para: CDispatch, value: Optional[float], rule: Union[str, int, None])`

设置行间距。常用组合：

```python
set_line_spacing(para, 1.5, "1.5")     # 1.5倍行距
set_line_spacing(para, 2.0, "double")  # 2倍行距
set_line_spacing(para, 20, "at_least")  # 固定值，最小20磅
set_line_spacing(para, 20, "exact")     # 固定值，精确20磅
set_line_spacing(para, None, "single")  # 单倍行距（仅设规则）
```

#### `set_space_before(para: CDispatch, points: float)`

设置段前间距（磅值）。

#### `set_space_after(para: CDispatch, points: float)`

设置段后间距（磅值）。

#### `set_indent_left(para: CDispatch, characters=None, cm=None)`

设置左缩进。优先使用厘米单位：

```python
set_indent_left(para, cm=2.0)      # 2厘米缩进
set_indent_left(para, characters=4) # 4字符缩进
```

#### `set_indent_right(para: CDispatch, characters=None, cm=None)`

设置右缩进。

#### `set_first_line_indent(para: CDispatch, characters=None, cm=None)`

设置首行缩进。传负值向左缩进 = 悬挂缩进：

```python
set_first_line_indent(para, cm=0.74)     # 首行缩进0.74厘米（标准中文段落）
set_first_line_indent(para, cm=-0.74)    # 悬挂缩进0.74厘米（代替首行缩进）
```

#### `set_hanging_indent(para: CDispatch, characters: float)`

设置悬挂缩进（首行外的其他行左缩进）。

原理：设置 `FirstLineIndent = LeftIndent - characters`。

#### `set_outline_level(para: CDispatch, level: int)`

设置大纲级别（0=正文，1~9=标题级别）。

#### `set_keep_together(para: CDispatch, on: bool = True)`

段内不分页（Keeps lines of the paragraph on the same page）。

#### `set_keep_with_next(para: CDispatch, on: bool = True)`

与下段同页。

#### `set_page_break_before(para: CDispatch, on: bool = True)`

段前分页。

#### `set_widow_control(para: CDispatch, on: bool = True)`

孤行控制（prevent single lines of a paragraph from appearing at the top/bottom of a page）。

#### `set_style(para: CDispatch, style_name: str)`

应用样式到段落。

#### `reset_format(para: CDispatch)`

将段落格式恢复为默认（清空直接格式）。

### 边框与底纹方法

#### `set_border(para: CDispatch, side="bottom", line_style=1, line_width=4, color=0x000000, space=6.0)`

给段落指定边添加边框。

- `side`: `"top"` `"bottom"` `"left"` `"right"`
- `line_style`: 0=无, 1=单线, 2=双线, 3=点线, 4=粗线...
- `line_width`: 1=0.5磅, 4=0.75磅, 6=1磅, 8=1.5磅, 18=2.25磅
- `color`: 整数颜色值或颜色名字符串
- `space`: 边框与文字间距（磅值）

#### `clear_border(para: CDispatch)`

清除段落的所有边框。

#### `set_shading(para: CDispatch, fill_color="yellow", texture=0)`

设置段落底纹（背景填充色）。

- `texture`: 0=纯色, 1=横线, 2=竖线, 3=斜线...

#### `clear_shading(para: CDispatch)`

清除段落底纹（设为白色）。

### 编号列表方法

#### `list_count() -> int`

返回文档中 ListParagraphs（编号段落）的数量。

#### `list_paragraphs() -> List[CDispatch]`

返回所有编号段落对象列表。

#### `is_list_paragraph(para: CDispatch) -> bool`

判断段落是否在 ListParagraphs 集合中。

#### `apply_bullet(para: CDispatch, bullet_type="bullet")`

将段落转换为项目符号列表项。

#### `apply_numbering(para: CDispatch, number_format="decimal", start_at=1)`

将段落转换为编号列表项。

#### `remove_list_format(para: CDispatch)`

移除段落的列表格式（还原为普通段落）。

#### `get_list_level(para: CDispatch) -> int`

返回段落在列表中的级别（1-based）。

#### `set_list_level(para: CDispatch, level: int)`

设置段落在列表中的级别。

#### `get_list_number(para: CDispatch) -> Optional[int]`

获取段落的当前编号值（如 "3." 中的 3）。列表项外返回 None。

### 段落内容操作方法（CRUD）

#### `set_text(para: CDispatch, text: str)`

替换段落文本内容，保留段落格式。

#### `insert_text_before(para: CDispatch, text: str) -> CDispatch`

在段落开头插入文本，返回段落 Range。

#### `insert_text_after(para: CDispatch, text: str) -> CDispatch`

在段落末尾（段落标记之前）插入文本，返回段落 Range。

#### `delete_paragraph(para: CDispatch)`

删除整个段落（内容 + 段落标记）。**慎用**：会合并相邻段落。

#### `clear_paragraph(para: CDispatch)`

清空段落文本内容（保留段落标记）。

#### `add_paragraph_after(para: CDispatch) -> CDispatch`

在段落之后插入新段落，返回新段落对象。

#### `add_paragraph_before(para: CDispatch) -> CDispatch`

在段落之前插入新段落，返回新段落对象。

#### `add_empty_paragraph_after(para: CDispatch) -> CDispatch`

在段落之后插入空段落，返回当前段落对象。

#### `merge_with_next(para: CDispatch) -> bool`

将当前段落与下一段合并。返回是否成功（最后一段无法合并）。

#### `merge_with_previous(para: CDispatch) -> bool`

将当前段落与上一段合并。返回是否成功（第一段无法合并）。

#### `split_paragraph(para: CDispatch, separator="\t") -> List[CDispatch]`

按分隔符将一个段落拆分为多个段落。返回拆分后的段落列表。

### 批量操作方法

#### `find_by_text(text: str, whole_word=False, match_case=False) -> List[CDispatch]`

在所有段落中查找包含指定文本的段落。

#### `find_empty_paragraphs() -> List[CDispatch]`

返回所有空段落。

#### `find_headings() -> List[CDispatch]`

返回所有标题段落。

#### `find_headings_by_level(level: int) -> List[CDispatch]`

返回指定大纲级别的标题段落。

#### `find_list_paragraphs() -> List[CDispatch]`

返回所有编号/项目符号列表段落。

#### `apply_format_to_all(align=None, line_spacing=None, line_spacing_rule=None, space_before=None, space_after=None, indent_left=None, indent_right=None, first_line_indent=None) -> int`

批量设置所有段落格式。返回修改的段落数量。

#### `reverse_order(start=1, end=None) -> List[CDispatch]`

将指定范围内段落的文本内容原地反转（先内容末变末、首变尾）。

### Range / Selection 互操作方法

#### `get_paragraph_at_selection() -> CDispatch`

获取当前 Selection 所在段落。

#### `get_paragraph_at_range(rng: CDispatch) -> CDispatch`

获取指定 Range 所在段落。

#### `select_paragraph(para: CDispatch)`

选中整个段落（切换到 Selection 模式，影响界面）。

#### `select_range_of_paragraphs(start: int, end: int)`

选中 start 到 end 范围内的所有段落。

### 文档结构方法

#### `get_outline_summary() -> List[dict]`

返回文档大纲摘要。返回格式：

```python
[
    {"level": 0, "text": "正文内容...", "index": 1},
    {"level": 1, "text": "第一章 概述", "index": 2},
    {"level": 2, "text": "1.1 背景", "index": 3},
    ...
]
```

#### `get_document_structure() -> List[dict]`

返回文档完整段落结构。返回格式：

```python
[
    {
        "index": 1,
        "text": "段落文本（前60字符）...",
        "style": "正文",
        "level": 0,
        "is_heading": False,
        "is_empty": False,
        "is_list": False,
    },
    ...
]
```

---

## 3. Action 注册表映射

每个 action 在 AI Agent 执行层（action_registry.py）的映射规则：

### 简单索引操作（单个段落）

索引参数 `index` 支持正数（1-based）和负数（-1 = 最后一段）：

| Action | 方法 | 参数 |
|---|---|---|
| `get_paragraph_count` | `count()` | - |
| `get_paragraph_by_index` | `get(index)` | index |
| `get_paragraph_text` | `get_text(para)` | index |
| `get_paragraph_style` | `get_style_name(para)` | index |
| `get_outline_level` | `get_outline_level(para)` | index |
| `get_list_level` | `get_list_level(para)` | index |
| `is_paragraph_list_item` | `is_list_item(para)` | index |
| `is_paragraph_in_table` | `is_in_table(para)` | index |
| `get_paragraph_format_info` | `get_format_info(para)` | index |
| `set_paragraph_text` | `set_text(para, text)` | index, text |
| `set_paragraph_alignment` | `set_alignment(para, align)` | index, alignment |
| `set_paragraph_line_spacing` | `set_line_spacing(para, value, rule)` | index, spacing, rule |
| `set_paragraph_space_before` | `set_space_before(para, points)` | index, points |
| `set_paragraph_space_after` | `set_space_after(para, points)` | index, points |
| `set_paragraph_indent_left` | `set_indent_left(para, characters, cm)` | index, characters, cm |
| `set_paragraph_indent_right` | `set_indent_right(para, characters, cm)` | index, characters, cm |
| `set_paragraph_first_line_indent` | `set_first_line_indent(para, characters, cm)` | index, characters, cm |
| `set_paragraph_hanging_indent` | `set_hanging_indent(para, characters)` | index, characters |
| `set_paragraph_outline_level` | `set_outline_level(para, level)` | index, level |
| `set_paragraph_keep_together` | `set_keep_together(para, on)` | index, on |
| `set_paragraph_keep_with_next` | `set_keep_with_next(para, on)` | index, on |
| `set_paragraph_style` | `set_style(para, style_name)` | index, style_name |
| `reset_paragraph_format` | `reset_format(para)` | index |
| `set_paragraph_border` | `set_border(para, side, ...)` | index, side, line_style, line_width, color |
| `clear_paragraph_border` | `clear_border(para)` | index |
| `set_paragraph_shading` | `set_shading(para, fill_color)` | index, fill_color |
| `clear_paragraph_shading` | `clear_shading(para)` | index |
| `apply_bullet_list` | `apply_bullet(para)` | index |
| `apply_numbered_list` | `apply_numbering(para, ...)` | index, number_format, start_at |
| `remove_list_format` | `remove_list_format(para)` | index |
| `set_list_level` | `set_list_level(para, level)` | index, level |
| `delete_paragraph` | `delete_paragraph(para)` | index |
| `clear_paragraph` | `clear_paragraph(para)` | index |
| `add_paragraph_after` | `add_paragraph_after(para)` | index |
| `add_paragraph_before` | `add_paragraph_before(para)` | index |
| `merge_with_next` | `merge_with_next(para)` | index |
| `merge_with_previous` | `merge_with_previous(para)` | index |
| `split_paragraph_by_separator` | `split_paragraph(para, separator)` | index, separator |
| `insert_text_before_paragraph` | `insert_text_before(para, text)` | index, text |
| `insert_text_after_paragraph` | `insert_text_after(para, text)` | index, text |
| `select_paragraph` | `select_paragraph(para)` | index |

### 范围操作

| Action | 方法 | 参数 |
|---|---|---|
| `get_paragraph_range` | `range(start, end)` | start, end |
| `apply_format_to_range` | `apply_format_to_all(...)` | start, end, 及其他格式参数 |
| `reverse_paragraph_order` | `reverse_order(start, end)` | start, end |
| `select_paragraph_range` | `select_range_of_paragraphs(start, end)` | start, end |

### 文档级操作（无索引）

| Action | 方法 | 参数 |
|---|---|---|
| `get_document_structure` | `get_document_structure()` | - |
| `get_outline_summary` | `get_outline_summary()` | - |
| `find_empty_paragraphs` | `find_empty_paragraphs()` | - |
| `find_heading_paragraphs` | `find_headings()` | - |
| `find_paragraphs_by_level` | `find_headings_by_level(level)` | level |
| `find_paragraphs_by_text` | `find_by_text(text, ...)` | text, whole_word, match_case |
| `get_list_paragraphs` | `list_paragraphs()` | - |
| `delete_empty_paragraphs` | `find_empty_paragraphs()` + delete | - |
| `get_paragraph_at_selection` | `get_paragraph_at_selection()` | - |
