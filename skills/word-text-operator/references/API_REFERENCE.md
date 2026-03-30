# Word Text Operator - API Reference

## Overview

`WordTextOperator` 是 word-text-operator 技能的统一入口，整合了以下子模块：

| Module | File | Responsibility |
|---|---|---|
| `WordBase` | `word_base.py` | COM 连接生命周期管理 |
| `RangeNavigator` | `word_range_navigation.py` | Range 移动/扩展/比较/GoTo |
| `TextOperator` | `word_text_operations.py` | 文本读写/插入/删除/大小写/统计 |
| `TextFormatter` | `word_format.py` | 字体/段落/边框/底纹格式 |
| `FindReplace` | `word_find_replace.py` | 查找/替换/通配符/批量 |
| `BookmarkOperator` | `word_bookmark.py` | 书签 CRUD / 导入导出 |
| `SelectionOperator` | `word_selection.py` | 光标移动/选区扩展/Selection 专属操作 |

---

## Quick Start

```python
from scripts.word_text_operator import WordTextOperator

with WordTextOperator("document.docx") as op:
    # 查找并格式化
    rng = op.find("关键词")
    if rng:
        op.set_bold(rng, True)
        op.set_font_color(rng, "red")

    # 替换
    op.replace("旧文本", "新文本")

    # 创建书签
    rng = op.find("重要段落")
    op.bm.create(rng, "重要标记")
```

---

## WordTextOperator Main API

### Lifecycle

| Method | Description |
|---|---|
| `WordTextOperator(doc_path=None)` | 创建实例 |
| `connect(doc_path=None)` | 连接 Word |
| `disconnect(save_changes=False)` | 断开连接 |
| `with` context manager | 自动管理生命周期 |

### Text

| Method | Returns | Description |
|---|---|---|
| `get_text(start, end)` | `str` | 读取 [start, end) 范围的文本 |
| `get_full_text()` | `str` | 读取整个文档 |
| `get_selection_text()` | `str` | 读取当前选中内容 |
| `to_uppercase(rng)` | - | 转为全大写 |
| `to_lowercase(rng)` | - | 转为全小写 |
| `to_title_case(rng)` | - | 转为标题格式 |

### Find & Replace

| Method | Returns | Description |
|---|---|---|
| `find(text)` | `Range\|None` | 查找第一个匹配 |
| `find_all(text)` | `list[dict]` | 查找所有匹配的 {start,end,text} |
| `count_occurrences(text)` | `int` | 统计出现次数 |
| `replace(find, replace)` | `int` | 替换所有匹配 |
| `find_wildcards(pattern)` | `int` | 通配符查找/替换 |
| `find_with_format(text, bold=...)` | `bool` | 带格式查找 |
| `replace_with_format(find, replace, bold=...)` | `int` | 替换并应用格式 |
| `batch_replace({old:new, ...})` | `dict` | 批量替换 |

### Bookmark

| Method | Returns | Description |
|---|---|---|
| `create_bookmark(name, start, end)` | `bool` | 创建书签 |
| `go_to_bookmark(name)` | `Range\|None` | 跳转到书签 |
| `get_bookmarks()` | `list[dict]` | 列出所有书签 |
| `bookmark_text(name, text)` | `bool` | 查找文本并添加书签 |

### Format (Range-level)

| Method | Description |
|---|---|
| `set_bold(rng, True/False)` | 加粗/取消加粗 |
| `set_italic(rng, True/False)` | 斜体/取消斜体 |
| `set_underline(rng, "single"/"double"/...)` | 下划线 |
| `set_font_color(rng, "#FF0000"/"red"/0xFF0000)` | 字体颜色 |
| `set_font_name(rng, "宋体"/"Arial")` | 字体名称 |
| `set_font_size(rng, 14.0)` | 字号（磅） |
| `set_paragraph_alignment(rng, "center"/"justify"/...)` | 段落对齐 |
| `set_highlight(rng, "yellow")` | 高亮背景色 |

### Range Utilities

| Method | Returns | Description |
|---|---|---|
| `get_range(start, end)` | `Range` | 获取指定范围的 Range |
| `get_full_range()` | `Range` | 获取整个文档的 Range |
| `get_selection_range()` | `Range` | 获取当前 Selection 的 Range |
| `select(rng)` | - | 选中 Range |
| `expand_to_word(rng)` | `int` | 扩展到完整单词 |
| `expand_to_sentence(rng)` | `int` | 扩展到完整句子 |
| `expand_to_paragraph(rng)` | `int` | 扩展到完整段落 |
| `collapse(rng, "start"/"end")` | - | 折叠 Range |
| `move(rng, unit, count)` | `int` | 移动 Range |
| `compare_ranges(rng1, rng2)` | `int` | 比较位置（-1/0/1） |

### Insert / Delete

| Method | Description |
|---|---|
| `insert_text(rng, text, before=True)` | 插入文本 |
| `insert_page_break(rng)` | 插入分页符 |
| `delete_range(rng)` | 删除 Range 内容 |
| `delete_selection()` | 删除当前选中 |

### Statistics

| Method | Returns | Description |
|---|---|---|
| `char_count(rng)` | `int` | 字符数 |
| `word_count(rng)` | `int` | 单词数 |

---

## Submodule Reference

### RangeNavigator (`word_range_navigation.py`)

Range 导航与范围管理：

```python
# 基础
nav.get_range(0, 100)         # 按字符偏移创建 Range
nav.get_full_range()           # 整个文档
nav.get_selection_range()      # Selection -> Range
nav.set_range(rng, start, end) # 重设起止
nav.clone_range(rng)           # 复制 Range

# Expand
nav.expand_to_sentence(rng)    # 扩展到句子
nav.expand_to_paragraph(rng)   # 扩展到段落
nav.expand_to_line(rng)        # 扩展到行
nav.expand_to_word(rng)        # 扩展到单词
nav.expand_to_document(rng)     # 扩展到整篇文档
nav.select_range(rng)          # 选中（切换到 Selection）

# Move
nav.move(rng, unit=4, count=1)  # 整体移动（unit: 1=char, 2=word, 4=para...）
nav.move_start(rng, unit, count) # 移动起始位置
nav.move_end(rng, unit, count)   # 移动结束位置
nav.move_while(rng, " \t\n")    # 跳过空白字符
nav.move_until(rng, ".,!?")     # 移动直到遇到标点
nav.collapse(rng, "start")     # 折叠为空

# Compare
nav.in_range(rng, container)   # rng 是否在 container 内
nav.compare_location(rng1, rng2) # -1/0/1 比较
nav.is_equal(rng1, rng2)       # 是否完全相同
nav.is_inside(rng, container)  # 是否严格在内部

# GoTo
nav.go_to_bookmark("bm1")      # 跳转到书签
nav.go_to_comment(3)           # 跳转到第3条批注
nav.go_to_page(5)              # 跳转到第5页
nav.go_to_line(10)             # 跳转到第10行
nav.go_to_start()              # 文档开头
nav.go_to_end()                # 文档末尾
```

### TextOperator (`word_text_operations.py`)

文本内容操作：

```python
# 读写
text.get_text(rng)             # 纯文本
text.get_formatted_text(rng)   # 带格式文本
text.get_full_document_text()  # 全文
text.get_paragraph_text(0)    # 指定段落文本（索引从0）

# 写入
text.set_text(rng, "新文本")    # 替换内容
text.replace_text(rng, old, new) # 内部替换

# 插入
text.insert_before(rng, "前缀")  # Range 前插入
text.insert_after(rng, "后缀")   # Range 后插入
text.insert_file(rng, "file.docx") # 插入文件
text.insert_break(rng, 6)         # 插入分页符（Type=6）
text.insert_paragraph(rng)        # 插入段落标记
text.insert_symbol(rng, 9679, "Wingdings") # 插入符号

# 删除
text.delete(rng, unit=1, count=1)  # 删除内容
text.delete_all(rng)               # 删除所有
text.clear(rng)                    # 清空

# 大小写
text.to_lowercase(rng)
text.to_uppercase(rng)
text.to_title_case(rng)
text.to_toggle_case(rng)

# 统计
text.char_count(rng)
text.word_count(rng)
text.sentence_count(rng)
text.paragraph_count(rng)
text.line_count(rng)

# 复制/粘贴
text.copy(rng)
text.cut(rng)
text.paste(rng)
text.paste_formatted(rng)
```

### TextFormatter (`word_format.py`)

字符和段落格式：

```python
# 字体
fmt.set_font_name(rng, "宋体")
fmt.set_font_size(rng, 14.0)
fmt.set_bold(rng, True)
fmt.set_italic(rng, True)
fmt.set_underline(rng, "single")   # "none"/"single"/"double"/"words"...
fmt.set_underline_color(rng, "blue")
fmt.set_strike_through(rng, True)  # 删除线
fmt.set_superscript(rng, True)     # 上标
fmt.set_subscript(rng, True)       # 下标
fmt.set_small_caps(rng, True)      # 小型大写
fmt.set_all_caps(rng, True)        # 全大写
fmt.set_hidden(rng, True)          # 隐藏文字

# 颜色
fmt.set_font_color(rng, "#FF0000") # 字体颜色
fmt.set_highlight(rng, "yellow")   # 高亮背景
fmt.clear_highlight(rng)

# 间距
fmt.set_spacing(rng, 3.0)          # 字符间距（磅）
fmt.set_vertical_position(rng, 6)  # 上下偏移（磅）
fmt.set_kerning(rng, 10.0)         # 字距调整阈值
fmt.set_character_width(rng, 120)  # 宽度百分比

# 段落
fmt.set_alignment(rng, "center")   # "left"/"center"/"right"/"justify"
fmt.set_line_spacing(rng, 1.5, "1.5")  # 行距
fmt.set_space_before_para(rng, 12) # 段前
fmt.set_space_after_para(rng, 6)   # 段后
fmt.set_indent_left(rng, cm=2.0)   # 左缩进
fmt.set_indent_right(rng, cm=1.0)  # 右缩进
fmt.set_first_line_indent(rng, cm=0.74) # 首行缩进
fmt.set_outline_level(rng, 1)      # 大纲级别（1-9）
fmt.set_keep_together(rng, True)   # 段内不分页
fmt.set_keep_with_next(rng, True)  # 与下段同页

# 边框
fmt.set_border(rng, "bottom", line_style=1, color=0x000000)
fmt.set_box_border(rng)            # 四边框
fmt.clear_border(rng)

# 底纹
fmt.set_shading(rng, fill_color="yellow")
fmt.clear_shading(rng)

# 读取
fmt.get_font_info(rng)             # -> dict
fmt.get_paragraph_format_info(rng) # -> dict
fmt.get_format_summary(rng)        # 人类可读摘要
```

### FindReplace (`word_find_replace.py`)

高级查找与替换：

```python
# 基础
find.find_in_range(rng, "文本")       # 单次查找
find.find_next_in_range(rng, "文本")  # 查找并返回匹配 Range
find.replace_in_range(rng, old, new)  # 范围替换
find.replace_in_document(old, new)    # 全文档替换
find.replace_in_selection(old, new)   # Selection 中替换

# 遍历
for match in find.find_all_in_range(rng, "关键词"):
    print(match.Text)

positions = find.find_all_positions(rng, "关键词")
# [{'start': 10, 'end': 15, 'text': '关键词'}, ...]

# 通配符
find.find_wildcards_in_range(rng, "<[A-Z][a-z]+>", replace_text="X")
# <词开头, [A-Z]大写 [a-z]+多个小写 = 标题格式词

# 带格式查找
find.find_with_format_in_range(rng, "文本", bold=True, italic=True)
find.replace_with_format(rng, old, new, bold=True, font_size=14)

# 批量
find.batch_find(rng, ["词1", "词2", "词3"])  # 统计每个词出现次数
find.batch_replace(rng, {"A": "B", "C": "D"})  # 批量替换

# 快捷
find.highlight_all(rng, "关键词", highlight_color=7)  # 高亮
find.replace_paragraph_marks(rng, separator=" ")  # 段落标记替换
find.count_matches(rng, "关键词")  # 统计次数
```

### BookmarkOperator (`word_bookmark.py`)

书签管理：

```python
# 创建
bm.create(rng, "我的书签")
bm.create_at_selection("光标书签")
bm.create_quick_bookmark(rng, "快速标签")

# 读取
bm.list_all()      # [{'name', 'start', 'end', 'text'}, ...]
bm.get("书签名")    # Bookmark 对象
bm.exists("书签名") # bool
bm.get_range("书签名")  # Range
bm.get_text("书签名")  # 书签内容文本
bm.get_bookmark_info("书签名") # 详细信息

# 更新
bm.update_range("书签", new_start, new_end)
bm.rename("旧名", "新名")
bm.select("书签名")  # 跳转到书签

# 删除
bm.delete("书签名")
bm.delete_all()
bm.delete_in_range(rng)  # 删除 Range 内的所有书签

# 导入导出
bm.export_bookmarks("bm.json")
bm.import_bookmarks("bm.json")  # 返回导入数量
```

### SelectionOperator (`word_selection.py`)

Selection 专用操作（与 Range 互补）：

```python
# 基础
sel.has_selection      # 是否有选中
sel.is_collapsed       # 是否折叠（光标点）
sel.selection_text     # 选中文本
sel.selection_range    # 转为 Range
sel.get_selection_info() # 详细信息

# 折叠/展开
sel.collapse_to_start()
sel.collapse_to_end()
sel.expand_to_word()
sel.expand_to_paragraph()
sel.expand_to_line()

# 光标移动
sel.move(unit=2, count=3)     # 移动 3 个词
sel.move_left(unit=1, count=2)   # 左移 2 字符
sel.move_right(unit=2, count=1)  # 右移 1 词
sel.move_up(count=1)             # 上移一行
sel.move_down(count=1)           # 下移一行
sel.move_to_document_start()
sel.move_to_document_end()
sel.move_to_line_start()
sel.move_to_line_end()

# 扩展选区
sel.extend_left(count=3)    # 向左扩展 3 字符
sel.extend_right(count=1)   # 向右扩展 1 词
sel.extend_to_match()       # 扩展到配对字符

# 选中
sel.select_word()
sel.select_line()
sel.select_paragraph()
sel.select_all()

# 查找
sel.find_and_select("文本")
sel.find_next_and_select("文本")
sel.find_previous_and_select("文本")
sel.replace_selection(old, new, replace_all=True)

# 格式化
sel.set_bold(True)
sel.set_font_name("Arial")
sel.set_font_size(14)
sel.clear_formatting()

# 内容操作
sel.type_text("输入文本")
sel.delete_selection()
sel.insert_paragraph()
sel.cut_selection()
sel.copy_selection()
sel.paste_selection()
```

---

## wdUnit 常量速查

| 常量 | 值 | 说明 |
|---|---|---|
| `WD_UNIT_CHARACTER` | 1 | 字符 |
| `WD_UNIT_WORD` | 2 | 单词 |
| `WD_UNIT_SENTENCE` | 3 | 句子 |
| `WD_UNIT_PARAGRAPH` | 4 | 段落 |
| `WD_UNIT_LINE` | 5 | 行 |
| `WD_UNIT_STORY` | 6 | 整篇文档 |

## wdColor 常量速查（`word_format.py` 中已定义）

| 名称 | 值 |
|---|---|
| `black` | 0x000000 |
| `red` | 0xFF0000 |
| `blue` | 0x0000FF |
| `green` | 0x00FF00 |
| `yellow` | 0xFFFF00 |
| `white` | 0xFFFFFF |
| `dark_red` | 0x800000 |
| `dark_blue` | 0x000080 |
| `gray` | 0x808080 |
| `auto` | 0xFFFFFFFF |

## wdUnderline 常量速查

| 名称 | 值 |
|---|---|
| `none` | 0 |
| `single` | 1 |
| `words` | 2 |
| `double` | 3 |
| `thick` | 4 |
| `dotted` | 5 |
| `dash` | 7 |
| `wavy` | 15 |

## wdAlignment 常量

| 名称 | 值 | 说明 |
|---|---|---|
| `left` | 0 | 左对齐 |
| `center` | 1 | 居中 |
| `right` | 2 | 右对齐 |
| `justify` | 3 | 两端对齐 |
| `distribute` | 4 | 分散对齐 |
