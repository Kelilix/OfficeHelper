---
name: word-page-operation
description: Word 文档页面格式化操作技能。当用户要求修改页面设置（如页边距、纸张大小、页面方向）时触发。支持以下操作：
  - 设置页边距（上、下、左、右）
  - 设置纸张大小（A4、A3、A5、Letter）
  - 设置页面方向（纵向、横向）
---

# Word 页面操作技能

## 功能说明

本技能提供 Word 文档的页面格式化操作能力。

## 使用前提

1. 文档已打开
2. 操作应用于整个文档（或选中的节）

## 可用操作

### 1. 设置页边距 (set_page_margins)

设置页面的上、下、左、右边距。

**参数：**
- `top`: 上边距，数值（厘米），推荐范围 1.27-5
- `bottom`: 下边距，数值（厘米），推荐范围 1.27-5
- `left`: 左边距，数值（厘米），推荐范围 1.27-5
- `right`: 右边距，数值（厘米），推荐范围 1.27-5

### 2. 设置纸张大小 (set_paper_size)

设置文档的纸张大小。

**参数：**
- `paper_size`: 纸张大小，字符串
  - "A4": A4 纸 (210×297mm)
  - "A3": A3 纸 (297×420mm)
  - "A5": A5 纸 (148×210mm)
  - "Letter": Letter 纸 (8.5×11in / 216×279mm)

### 3. 设置页面方向 (set_page_orientation)

设置页面的方向。

**参数：**
- `orientation`: 方向，字符串
  - "portrait": 纵向
  - "landscape": 横向

## 执行流程

1. 确认用户的页面设置需求
2. 调用相应的 API 方法执行操作
3. 返回执行结果

## API 接口

```python
# 设置页边距
word_connector.set_page_margins(top=2.54, bottom=2.54, left=3.17, right=3.17)

# 设置纸张大小
word_connector.set_paper_size("A4")  # A4/A3/A5/Letter

# 设置页面方向
word_connector.set_page_orientation("landscape")  # portrait/landscape
```

## 常用标准

- 中文文档标准（GB/T 9704）：上 3.7cm，下 3.5cm，左 2.8cm，右 2.6cm
- 通用文档：A4，纵向，边距 2.54cm

## 注意事项

- 页面设置应用于整个文档（除非文档有分节）
- 修改页边距可能影响分页和排版
- 纸张大小的改变可能导致内容重新分页
