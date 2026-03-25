---
name: word-font-operation
description: Word 文档字体格式化操作技能。当用户要求修改文字的字体、字号、加粗、斜体、下划线、颜色等字体属性时触发。支持以下操作：
  - 修改字体名称（如：宋体、黑体、楷体、仿宋、微软雅黑）
  - 修改字号大小（5-72磅）
  - 设置加粗/取消加粗
  - 设置斜体/取消斜体
  - 添加/取消下划线
  - 修改文字颜色
---

# Word 字体操作技能

## 功能说明

本技能提供 Word 文档中文字的字体格式化操作能力。

## 使用前提

1. 用户已选中文档中的文字
2. 操作对象是当前选中的文字区域

## 可用操作

### 1. 设置字体 (set_font)

修改选中文字的字体类型。

**参数：**
- `font_name`: 字体名称，字符串

**常用字体：**
- 宋体、黑体、楷体、仿宋、微软雅黑
- Arial、Times New Roman、Calibri

### 2. 设置字号 (set_font_size)

修改选中文字的字号大小。

**参数：**
- `size`: 字号，数值类型（5-72磅）

**常用字号：**
- 小五(9)、五号(10.5)、小四(12)、四号(14)、小三(15)、三号(16)、小二(18)、二号(22)、小一(24)、一号(26)

### 3. 设置加粗 (set_bold)

将选中文字设置为加粗或取消加粗。

**参数：**
- `bold`: true（加粗）或 false（取消加粗）

### 4. 设置斜体 (set_italic)

将选中文字设置为斜体或取消斜体。

**参数：**
- `italic`: true（斜体）或 false（取消斜体）

### 5. 设置下划线 (set_underline)

为选中文字添加下划线或取消下划线。

**参数：**
- `underline`: true（添加）或 false（取消）

### 6. 设置文字颜色 (set_font_color)

修改选中文字的颜色。

**参数：**
- `color`: 颜色值，支持颜色名称（黑色、红色、蓝色、绿色）或十六进制（#000000、#FF0000）

## 执行流程

1. 确认用户选中了需要修改的文字
2. 理解用户的格式化需求
3. 调用相应的 API 方法执行操作
4. 返回执行结果

## API 接口

```python
# 设置字体
word_connector.set_font(font_name="宋体")

# 设置字号
word_connector.set_font(size=14)

# 设置加粗
word_connector.set_font(bold=True)

# 设置斜体
word_connector.set_font(italic=True)

# 设置下划线
word_connector.set_font(underline=True)

# 设置颜色
word_connector.set_font_color("#FF0000")
```

## 注意事项

- 所有操作仅影响当前选中的文字
- 如果没有选中文本，操作将应用于光标所在位置的新输入
- 字号单位为磅（pt），与 Word 中的字号一致
