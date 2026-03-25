---
name: word-paragraph-operation
description: Word 文档段落格式化操作技能。当用户要求修改段落的格式（如对齐方式、缩进、行距、段落间距）时触发。支持以下操作：
  - 设置对齐方式（左对齐、居中、右对齐、两端对齐）
  - 设置首行缩进或悬挂缩进
  - 设置行距（单倍、1.5倍、2倍等）
  - 设置段前/段后间距
---

# Word 段落操作技能

## 功能说明

本技能提供 Word 文档中段落的格式化操作能力。

## 使用前提

1. 用户已选中文档中的段落或将光标放在目标段落
2. 操作对象是当前选中段落或光标所在段落

## 可用操作

### 1. 设置对齐方式 (set_alignment)

设置段落的对齐方式。

**参数：**
- `alignment`: 对齐方式，字符串
  - "left": 左对齐
  - "center": 居中对齐
  - "right": 右对齐
  - "justify": 两端对齐

### 2. 设置首行缩进 (set_first_line_indent)

设置段落的首行缩进量。

**参数：**
- `first_line`: 首行缩进量，数值（磅）
  - 正值：首行缩进（如 21 磅 ≈ 0.74厘米 ≈ 2字符）
  - 负值：悬挂缩进
  - 0：取消首行缩进

### 3. 设置左边缩进 (set_left_indent)

设置段落的左边缩进。

**参数：**
- `left`: 左边缩进量，数值（磅）

### 4. 设置行距 (set_line_spacing)

设置段落内文字的行间距。

**参数：**
- `spacing`: 行距倍数，数值
  - 1.0: 单倍行距
  - 1.5: 1.5倍行距
  - 2.0: 2倍行距
  - 固定值如 12、14（单位：磅）

### 5. 设置段落间距 (set_paragraph_spacing)

设置段前和段后的间距。

**参数：**
- `space_before`: 段前间距，数值（磅）
- `space_after`: 段后间距，数值（磅）

## 执行流程

1. 确认用户选中了需要修改的段落
2. 理解用户的格式化需求
3. 调用相应的 API 方法执行操作
4. 返回执行结果

## API 接口

```python
# 设置对齐
word_connector.set_paragraph_alignment("justify")  # left/center/right/justify

# 设置首行缩进（21磅 ≈ 2字符）
word_connector.set_indent(first_line=21, indent_type="first_line")

# 设置左边缩进
word_connector.set_indent(left_indent=42, indent_type="left")

# 设置行距
word_connector.set_line_spacing(1.5)  # 1.0/1.5/2.0

# 设置段落间距
word_connector.set_paragraph_spacing(before=0, after=8)
```

## 常用格式标准

- 中文文档正文：首行缩进 2 字符，行距 1.5 倍
- 英文文档：行距 1.0 倍
- 标题：段前 12-18 磅，段后 6-12 磅

## 注意事项

- 所有操作仅影响当前选中的段落
- 如果没有选中任何内容，操作将应用于光标所在段落
- 缩进和间距的单位为磅（1 磅 ≈ 0.035 厘米）
