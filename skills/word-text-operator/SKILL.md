---
name: word-text-operator
description: Word 文档文本操作技能。当用户需要对 Word 文档中的文本进行读取、查找、替换、格式化、书签管理、光标导航等任何文本级操作时触发。
---

# Word 文本操作技能

## 核心原则

**优先使用 Range，Selection 仅在用户交互场景下使用。**

- `Range`：文档中一对字符位置（start, end），可编程操作，不影响光标
- `Selection`：当前用户看到的高亮区域，会改变界面状态

---

## ⚡ 选中内容的两步处理原则（最高优先级）

> **当用户请求中同时包含"整体目标"和"选中部分的额外要求"时，必须分两步处理。先用 Range 操作整体目标，再用 Selection 操作选中部分。**

### 为什么必须分两步

一旦执行了移动光标、扩选、查找等操作，**当前选区就会丢失**，后续操作将无法正确作用于用户原本选中的内容。两步分离可以保证选中信息不丢失。

### 规则

1. **Step 1（整体操作）**：使用 `rng: "[start, end]"` 或 `rng: "full_document"` 操作整体范围，**不得使用省略 rng**（因为省略会使用已丢失的选区）
2. **Step 2（选中部分）**：先调用 `get_selection_info` 获取当前选区的 start/end，再在 Step 2 的 actions 中用 `rng: "[start, end]"` 精确指定
3. **必须在 plan 阶段标注**：当有选中内容时，plan 中对应 step 的 `description` 末尾必须附加以下标记：
   ```
   [选中信息] 选中内容: "xxx" | 段落: 第N段 | 字符位置: [start, end]
   ```
4. **禁止在两步之间调用任何导航类 action**：Step 1 完成后、Step 2 执行前，禁止调用 `move_*`、`select_*`、`expand_to_*`、`find_and_select` 等

### 典型场景库

#### 场景 1：整体段落 + 选中部分加格式
**用户请求**：将第1段字体设置为宋体，四号字，选中的部分要额外设置为红色，加粗。

**正确 plan**：
```json
{
  "plan": [
    {
      "step": 1,
      "skill": "word-text-operator",
      "description": "将第1段设置为宋体、四号字 [选中信息] 选中内容: "歌加拉看" | 段落: 第1段 | 字符位置: [5, 9]",
      "need_feedback": false
    },
    {
      "step": 2,
      "skill": "word-text-operator",
      "description": "将选中部分设置为宋体、四号字、红色、加粗",
      "need_feedback": false
    }
  ]
}
```

**对应 actions**：
- Step 1：
  ```python
  # 获取第1段（index=1）的字符位置，再设格式
  {"action": "get_paragraph_range", "params": {"index": 1}, "description": "获取第1段字符范围"}
  ```
  → 返回 `[start, end]`，如 `[0, 26]`（整段字符范围）
  ```python
  {"action": "set_font_name", "params": {"rng": "[start, end]", "font_name": "宋体"}, "description": "第1段设宋体"}
  {"action": "set_font_size", "params": {"rng": "[start, end]", "size": 14.0}, "description": "第1段设四号"}
  ```
- Step 2（用注入的选中位置 `[sel_start, sel_end]`）：
  ```python
  # 注意：不需要再次调用 get_selection_info，prompt 中已注入选中位置
  {"action": "set_font_name", "params": {"rng": "[sel_start, sel_end]", "font_name": "宋体"}, "description": "选中部分设宋体"}
  {"action": "set_font_size", "params": {"rng": "[5, 9]", "size": 14.0}, "description": "选中部分设四号"}
  {"action": "set_bold", "params": {"rng": "[5, 9]", "bold": true}, "description": "选中部分加粗"}
  {"action": "set_font_color", "params": {"rng": "[5, 9]", "color": "red"}, "description": "选中部分设红色"}
  ```

#### 场景 2：整体文档 + 选中部分特殊格式
**用户请求**：全文设为黑体，选中的文字设为红色。

**正确做法**：Step 1 用 `rng: "full_document"` 整体设黑体；Step 2 用选中位置 `[start, end]` 单独设红色。

#### 场景 3：全文替换 + 保留选中位置
**用户请求**：把所有"ABC"替换成"XYZ"，但选中的那个"ABC"保持红色。

**正确做法**：Step 1 全局替换；Step 2 先 `get_selection_info` 拿到选中位置，再对 `[start, end]` 单独设红色。

### `rng` 参数约定

| 写法 | 含义 | 使用场景 |
|---|---|---|
| `"full_document"` | 整篇文档 | 全局操作 |
| `"[start, end]"` | 指定字符范围 | 已知位置时（两步处理中必须用这个） |
| **省略 / null** | 当前用户选区 | 仅在用户有实时选中且未做导航时使用 |

---

## 执行器

`WordTextOperator` 统一封装了以下子模块：

| 属性 | 类 | 适用场景 |
|---|---|---|
| `op` | `WordTextOperator` | 高频操作（查找/替换/格式）直接调用 |
| `op.nav` | `RangeNavigator` | Range 移动/扩展/GoTo/比较 |
| `op.text` | `TextOperator` | 文本读写/插入/删除/大小写/统计 |
| `op.fmt` | `TextFormatter` | 字体/段落/边框/底纹格式 |
| `op.find` | `FindReplace` | 高级查找/通配符/批量 |
| `op.bm` | `BookmarkOperator` | 书签 CRUD |
| `op.sel` | `SelectionOperator` | 光标/选区扩展/Selection 专属操作 |

**Range 参数约定（重要）**：大多数操作的 `params` 中 `rng` 支持三种写法：
- `"full_document"` → 整篇文档
- `"[start, end]"` → 指定字符范围（start 和 end 为整数）
- **省略或 `null` → 当前用户鼠标选中的区域（默认！）**

> ⚠️ **常见错误**：当用户使用鼠标选中了部分文字后提问，LLM 在调用操作时**必须省略 rng 参数或设为 `null`**，让系统使用用户当前的选区。如果 LLM 传了 `"rng": "full_document"` 或没有传 rng 导致系统错误地使用全文档，用户选中的区域外的文字也会被修改。

---

## 可用操作（按子模块分组）

### 文本读取

#### 1. get_full_text

读取整篇文档的纯文本。

```python
{"action": "get_full_text", "params": {}, "description": "读取整篇文档文本"}
```

#### 2. get_text

读取指定 Range 的纯文本。

```python
{"action": "get_text", "params": {"rng": "full_document"}, "description": "读取文档文本"}
```

#### 3. get_selection_text

读取当前用户选中的文本。

```python
{"action": "get_selection_text", "params": {}, "description": "读取当前选中文本"}
```

---

### 查找

#### 4. find

在文档中查找第一个匹配项，返回匹配 Range 的位置信息。

```python
{"action": "find", "params": {"text": "关键词", "whole_word": false, "match_case": false}, "description": "查找第一个匹配"}
```

#### 5. find_all

在文档中查找所有匹配项，返回位置列表 `[{start, end, text}, ...]`。

```python
{"action": "find_all", "params": {"text": "关键词"}, "description": "查找所有匹配"}
```

#### 6. count_occurrences

统计关键词在文档中出现的次数。

```python
{"action": "count_occurrences", "params": {"text": "关键词"}, "description": "统计出现次数"}
```

#### 7. find_wildcards

使用通配符模式查找或替换。

```python
# 仅查找
{"action": "find_wildcards", "params": {"pattern": "[0-9]+"}, "description": "通配符查找数字"}
# 查找并替换
{"action": "find_wildcards", "params": {"pattern": "[0-9]+", "replace_text": "X"}, "description": "数字替换为X"}
```

常用通配符：`?`任意字符 `*`任意多个 `[abc]`括号内 `>`词尾 `<`词头 `{n,m}`次数

#### 8. find_with_format

按文本内容和格式同时筛选查找。

```python
{"action": "find_with_format", "params": {"text": "关键词", "bold": true, "italic": null}, "description": "查找加粗的关键词"}
```

#### 9. find_and_select (Selection)

在文档中查找并选中文本（影响界面）。

```python
{"action": "find_and_select", "params": {"text": "关键词", "whole_word": false, "match_case": false}, "description": "查找并选中"}
```

---

### 替换

#### 10. replace

全文档替换。返回替换次数。

```python
{"action": "replace", "params": {"find_text": "旧", "replace_text": "新", "whole_word": false, "match_case": false, "replace_all": true}, "description": "替换文本"}
```

#### 11. replace_with_format

替换文本并同时设置替换后文本的格式。

```python
{"action": "replace_with_format", "params": {"find_text": "关键词", "replace_text": "替换词", "bold": true, "italic": false}, "description": "替换并加粗"}
```

#### 12. batch_replace

批量替换多组词对。参数为 JSON 对象 `{原词: 新词, ...}`。

```python
{"action": "batch_replace", "params": {"replacements": {"A": "B", "X": "Y"}}, "description": "批量替换"}
```

#### 13. replace_in_selection (Selection)

在当前 Selection 中替换。

```python
{"action": "replace_in_selection", "params": {"find_text": "旧", "replace_text": "新", "replace_all": true}, "description": "在选区中替换"}
```

---

### 字体格式（op 直接调用）

#### 14. set_bold

```python
{"action": "set_bold", "params": {"rng": "full_document", "bold": true}, "description": "设置加粗"}
```

#### 15. set_italic

```python
{"action": "set_italic", "params": {"rng": "[0, 10]", "italic": true}, "description": "设置斜体"}
```

#### 16. set_underline

下划线类型可选：`"single"` `"double"` `"words"` `"dotted"` `"dash"` `"wavy"` 等。

```python
{"action": "set_underline", "params": {"rng": "[0, 10]", "underline": "single"}, "description": "添加单下划线"}
```

#### 17. set_font_color

颜色支持颜色名（`"red"` `"blue"` `"black"` `"green"` 等）或十六进制（`"#FF0000"`）。

```python
{"action": "set_font_color", "params": {"rng": "[0, 10]", "color": "red"}, "description": "设置红色"}
```

#### 18. set_font_name

```python
{"action": "set_font_name", "params": {"rng": "[0, 10]", "font_name": "宋体"}, "description": "设置宋体"}
```

#### 19. set_font_size

字号单位为磅（pt）。

```python
{"action": "set_font_size", "params": {"rng": "full_document", "size": 14.0}, "description": "设置为14磅"}
```

#### 20. set_highlight

高亮色可选：`"yellow"` `"red"` `"green"` `"blue"` `"cyan"` 等。

```python
{"action": "set_highlight", "params": {"rng": "[0, 10]", "highlight": "yellow"}, "description": "添加黄色高亮"}
```

---

### 段落格式（op.fmt 调用）

#### 21. set_paragraph_alignment

对齐方式可选：`"left"` `"center"` `"right"` `"justify"` `"distribute"`。

```python
{"action": "set_paragraph_alignment", "params": {"rng": "[0, 10]", "alignment": "center"}, "description": "居中对齐"}
```

#### 22. set_line_spacing

行距规则可选：`"single"` `"1.5"` `"double"` `"at_least"` `"exact"`。

```python
{"action": "set_line_spacing", "params": {"rng": "full_document", "spacing": 1.5, "rule": "1.5"}, "description": "设置1.5倍行距"}
```

#### 23. set_indent_left

左缩进，`characters`（字符数）或 `cm`（厘米），二选一。

```python
{"action": "set_indent_left", "params": {"rng": "full_document", "cm": 2.0}, "description": "左缩进2厘米"}
```

#### 24. set_indent_right

```python
{"action": "set_indent_right", "params": {"rng": "full_document", "cm": 1.0}, "description": "右缩进1厘米"}
```

#### 25. set_first_line_indent

首行缩进（厘米）。传负值则为悬挂缩进。

```python
{"action": "set_first_line_indent", "params": {"rng": "full_document", "cm": 0.74}, "description": "首行缩进0.74厘米"}
```

#### 26. set_space_before

段前间距（磅）。

```python
{"action": "set_space_before", "params": {"rng": "full_document", "points": 12.0}, "description": "段前12磅"}
```

#### 27. set_space_after

段后间距（磅）。

```python
{"action": "set_space_after", "params": {"rng": "full_document", "points": 6.0}, "description": "段后6磅"}
```

#### 28. set_outline_level

大纲级别（0=正文，1-9=标题级别）。

```python
{"action": "set_outline_level", "params": {"rng": "full_document", "level": 1}, "description": "设为1级大纲"}
```

#### 29. set_keep_together

段内不分页。

```python
{"action": "set_keep_together", "params": {"rng": "full_document", "on": true}, "description": "段内不分页"}
```

#### 30. set_keep_with_next

与下段同页。

```python
{"action": "set_keep_with_next", "params": {"rng": "full_document", "on": true}, "description": "与下段同页"}
```

---

### 边框与底纹（op.fmt 调用）

#### 31. set_border

给 Range 的指定边添加边框。`side`：`"top"` `"bottom"` `"left"` `"right"` `"inside"`。

```python
{"action": "set_border", "params": {"rng": "[0, 10]", "side": "bottom", "line_style": 1, "line_width": 4, "color": 0x000000}, "description": "添加下边框"}
```

#### 32. clear_border

清除边框。

```python
{"action": "clear_border", "params": {"rng": "[0, 10]"}, "description": "清除边框"}
```

#### 33. set_shading

设置文字底纹（背景填充色）。

```python
{"action": "set_shading", "params": {"rng": "[0, 10]", "fill_color": "yellow", "texture": 0}, "description": "设置黄色背景"}
```

#### 34. clear_shading

清除底纹。

```python
{"action": "clear_shading", "params": {"rng": "[0, 10]"}, "description": "清除底纹"}
```

---

### Range 导航（op 或 op.nav 调用）

#### 35. expand_to_word

将 Range 扩展到完整单词。

```python
{"action": "expand_to_word", "params": {"rng": "[5, 8]"}, "description": "扩展到完整单词"}
```

#### 36. expand_to_sentence

将 Range 扩展到完整句子。

```python
{"action": "expand_to_sentence", "params": {"rng": "[5, 8]"}, "description": "扩展到完整句子"}
```

#### 37. expand_to_paragraph

将 Range 扩展到完整段落。

```python
{"action": "expand_to_paragraph", "params": {"rng": "[5, 8]"}, "description": "扩展到完整段落"}
```

#### 38. collapse

折叠 Range 为空（光标点）。`direction`：`"start"` 或 `"end"`。

```python
{"action": "collapse", "params": {"rng": "[5, 20]", "direction": "start"}, "description": "折叠到起点"}
```

#### 39. move

整体移动 Range。`unit`：1=字符 2=词 3=句子 4=段落 5=行。

```python
{"action": "move", "params": {"rng": "[5, 20]", "unit": 4, "count": 1}, "description": "向后移动1段"}
```

#### 40. goto_bookmark

跳转到指定书签位置并选中。

```python
{"action": "goto_bookmark", "params": {"name": "书签1"}, "description": "跳转到书签"}
```

#### 41. goto_page

跳转到指定页码（从 1 开始）。

```python
{"action": "goto_page", "params": {"page": 5}, "description": "跳转到第5页"}
```

#### 42. goto_line

跳转到指定行号。

```python
{"action": "goto_line", "params": {"line": 10}, "description": "跳转到第10行"}
```

---

### 书签（op 或 op.bm 调用）

#### 43. create_bookmark

在指定字符范围创建书签。`name` 不能含空格和特殊字符。

```python
{"action": "create_bookmark", "params": {"name": "my_bookmark", "start": 0, "end": 10}, "description": "创建书签"}
```

#### 44. delete_bookmark

删除指定书签（不删内容）。

```python
{"action": "delete_bookmark", "params": {"name": "my_bookmark"}, "description": "删除书签"}
```

#### 45. delete_all_bookmarks

删除所有书签。

```python
{"action": "delete_all_bookmarks", "params": {}, "description": "删除所有书签"}
```

#### 46. rename_bookmark

```python
{"action": "rename_bookmark", "params": {"old_name": "旧名", "new_name": "新名"}, "description": "重命名书签"}
```

#### 47. list_bookmarks

列出所有书签，返回 `[{name, start, end, text}, ...]`。

```python
{"action": "list_bookmarks", "params": {}, "description": "列出所有书签"}
```

#### 48. export_bookmarks

将书签导出为 JSON 文件。

```python
{"action": "export_bookmarks", "params": {"path": "bookmarks.json"}, "description": "导出书签"}
```

#### 49. import_bookmarks

从 JSON 文件导入书签。

```python
{"action": "import_bookmarks", "params": {"path": "bookmarks.json"}, "description": "导入书签"}
```

#### 50. bookmark_text

查找文本并为其添加书签。

```python
{"action": "bookmark_text", "params": {"name": "my_bookmark", "text": "要书签化的文本"}, "description": "为文本添加书签"}
```

#### 51. wrap_with_bookmarks

在文本两侧创建成对书签。

```python
{"action": "wrap_with_bookmarks", "params": {"text": "目标文本", "open_name": "start_tag", "close_name": "end_tag"}, "description": "两端加书签"}
```

---

### Selection 专属操作（op.sel 调用）

#### 52. get_selection_info

获取当前 Selection 详细信息（text/start/end/type）。

```python
{"action": "get_selection_info", "params": {}, "description": "获取选区信息"}
```

#### 53. move_left

光标左移。`extend=true` 为扩展选区模式。

```python
{"action": "move_left", "params": {"unit": 1, "count": 3, "extend": false}, "description": "左移3字符"}
```

#### 54. move_right

```python
{"action": "move_right", "params": {"unit": 2, "count": 1, "extend": false}, "description": "右移1词"}
```

#### 55. move_up

```python
{"action": "move_up", "params": {"unit": 5, "count": 1, "extend": false}, "description": "上移1行"}
```

#### 56. move_down

```python
{"action": "move_down", "params": {"unit": 5, "count": 1, "extend": false}, "description": "下移1行"}
```

#### 57. move_to_line_start

移动到当前行首。

```python
{"action": "move_to_line_start", "params": {}, "description": "移到行首"}
```

#### 58. move_to_line_end

```python
{"action": "move_to_line_end", "params": {}, "description": "移到行尾"}
```

#### 59. move_to_document_start

```python
{"action": "move_to_document_start", "params": {}, "description": "移到文档开头"}
```

#### 60. move_to_document_end

```python
{"action": "move_to_document_end", "params": {}, "description": "移到文档末尾"}
```

#### 61. select_word

选中光标所在单词。

```python
{"action": "select_word", "params": {}, "description": "选中单词"}
```

#### 62. select_line

```python
{"action": "select_line", "params": {}, "description": "选中整行"}
```

#### 63. select_paragraph

```python
{"action": "select_paragraph", "params": {}, "description": "选中整段"}
```

#### 64. select_all

选中整个文档。

```python
{"action": "select_all", "params": {}, "description": "全选"}
```

#### 65. select

选中指定 Range（可用于将 op.find 返回的 Range 变为可见选区）。

```python
{"action": "select", "params": {"rng": "[0, 10]"}, "description": "选中指定范围"}
```

#### 66. type_text

在光标处输入文本（替换已选中内容）。

```python
{"action": "type_text", "params": {"text": "输入的内容"}, "description": "输入文本"}
```

#### 67. clear_formatting

清除 Selection 的所有格式。

```python
{"action": "clear_formatting", "params": {}, "description": "清除格式"}
```

---

### 文本插入（op 或 op.text 调用）

#### 68. insert_text

在 Range 处插入文本。`before=true` 插入到 Range 之前，`before=false` 插入到之后。

```python
{"action": "insert_text", "params": {"rng": "[0, 0]", "text": "插入文本", "before": true}, "description": "在位置前插入"}
```

#### 69. insert_page_break

在 Range 处插入分页符。

```python
{"action": "insert_page_break", "params": {"rng": "[100, 100]"}, "description": "插入分页符"}
```

#### 70. insert_file

在 Range 处插入另一个文件的内容。

```python
{"action": "insert_file", "params": {"rng": "[0, 0]", "file_path": "template.docx"}, "description": "插入文件"}
```

#### 71. insert_symbol

在 Range 处插入符号。`character_code`：字符代码，`font_name`：符号字体（如 `"Wingdings 2"`）。

```python
{"action": "insert_symbol", "params": {"rng": "[0, 0]", "character_code": 9744, "font_name": "Wingdings 2"}, "description": "插入符号"}
```

#### 72. insert_paragraph

在 Range 处插入段落标记。

```python
{"action": "insert_paragraph", "params": {"rng": "[0, 0]"}, "description": "插入段落"}
```

---

### 文本删除（op 或 op.text 调用）

#### 73. delete_range

删除 Range 的内容（保留段落标记）。

```python
{"action": "delete_range", "params": {"rng": "[0, 10]"}, "description": "删除范围内容"}
```

#### 74. delete_selection

删除当前 Selection 的内容。

```python
{"action": "delete_selection", "params": {}, "description": "删除选中内容"}
```

#### 75. clear_range

清空 Range 内容（等价于设为空字符串）。

```python
{"action": "clear_range", "params": {"rng": "[0, 10]"}, "description": "清空范围"}
```

---

### 大小写转换（op 调用）

#### 76. to_uppercase

```python
{"action": "to_uppercase", "params": {"rng": "[0, 10]"}, "description": "转为全大写"}
```

#### 77. to_lowercase

```python
{"action": "to_lowercase", "params": {"rng": "[0, 10]"}, "description": "转为全小写"}
```

#### 78. to_title_case

每个单词首字母大写。

```python
{"action": "to_title_case", "params": {"rng": "[0, 10]"}, "description": "转为标题格式"}
```

---

### 统计（op 或 op.text 调用）

#### 79. char_count

统计 Range 内的字符数。

```python
{"action": "char_count", "params": {"rng": "full_document"}, "description": "统计字符数"}
```

#### 80. word_count

统计 Range 内的单词数。

```python
{"action": "word_count", "params": {"rng": "full_document"}, "description": "统计单词数"}
```

#### 81. sentence_count

统计 Range 内的句子数。

```python
{"action": "sentence_count", "params": {"rng": "full_document"}, "description": "统计句子数"}
```

#### 82. paragraph_count

统计 Range 内的段落数。

```python
{"action": "paragraph_count", "params": {"rng": "full_document"}, "description": "统计段落数"}
```

---

### 文档操作

#### 83. new_document

新建空白文档。

```python
{"action": "new_document", "params": {}, "description": "新建文档"}
```

#### 84. save

保存文档。`path` 为空则覆盖原文件。

```python
{"action": "save", "params": {"path": null}, "description": "保存文档"}
```

---

## 执行流程

仔细阅读用户需求，决定需要调用哪些 action。只返回 JSON 数组：

```json
[
  {"action": "find", "params": {"text": "关键词"}, "description": "查找关键词"},
  {"action": "set_bold", "params": {"rng": "[10, 20]", "bold": true}, "description": "加粗"}
]
```

如果只需要一个 action，也返回只含一个元素的数组。

## 典型案例库

以下案例展示常见需求的正确 action 组合方式。遇到类似需求时，可参考对应案例的行动序列。

---

### 案例 A：批量查找 + 统一格式（先查后改）

**场景**：用户说"把所有出现的「甲方」加粗并设为红色"

**分析**：需要两步——第一步用 `find_all` 查到所有位置，第二步系统自动将操作类 action 展开到每个位置执行。

**正确 actions**：

```json
[
  {"action": "find_all", "params": {"text": "甲方"}, "description": "查找所有甲方出现位置"},
  {"action": "set_bold", "params": {"bold": true}, "description": "所有匹配处加粗"},
  {"action": "set_font_color", "params": {"color": "red"}, "description": "所有匹配处设红色"}
]
```

> **注意**：`find_all` 返回多个 `[start, end]`，后续 action 的 rng 不要填——系统会自动将操作类 action 按每个匹配位置展开执行。查询类 action（`find_all`、`find` 等）只执行一次，取当前位置。

---

### 案例 B：两步处理原则（整体 + 选中部分）

**场景**：用户选中了部分文字，说"把全文字体改成宋体，选中的文字额外加红色"

**分析**：用户请求中同时包含"整体目标"和"选中部分的额外要求"，必须分两步。先用 `get_full_text` 或 `rng: "full_document"` 操作整体，再用 `get_selection_info` 获取选中位置后单独处理。

**正确 actions**：

```json
[
  {"action": "get_full_text", "params": {}, "description": "获取全文范围"},
  {"action": "set_font_name", "params": {"font_name": "宋体"}, "description": "全文设为宋体"}
]
```

> **系统会自动注入选中位置**到 Step 2 的后续 action，用户只需省略 rng 或填 `"[start, end]"` 占位符，系统会替换为真实值。

---

### 案例 C：书签辅助定位 + 精准插入

**场景**：用户说"在『合同编号』书签处插入文本『2026-001』"

**分析**：先跳转到书签选中位置，再在该位置插入文本。

**正确 actions**：

```json
[
  {"action": "goto_bookmark", "params": {"name": "合同编号"}, "description": "跳转到书签位置"},
  {"action": "insert_text", "params": {"text": "2026-001", "before": false}, "description": "在书签位置插入文本"}
]
```

> 书签名称中禁止空格和特殊字符，请用下划线（如 `"合同编号"` → `"合同编号"`）。

---

## 注意事项

- `rng` 支持 `"full_document"`、`"[start, end]"`、省略/空（当前选区）三种
- 书签名称禁止包含空格，请用下划线（如 `"my_bookmark"`）
- 颜色名支持：`black` `white` `red` `green` `blue` `yellow` `cyan` `magenta` `gray` `dark_red` `dark_green` `dark_blue`
- 详细 API 参考见 [references/API_REFERENCE.md](references/API_REFERENCE.md)
