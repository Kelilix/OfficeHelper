# -*- coding: utf-8 -*-
"""
Word 段落操作模块

提供基于 Paragraphs / Paragraph / ParagraphFormat 的高级段落操作能力：
- 段落 CRUD：增、删、合并、拆分
- 段落属性读取与写入（对齐、行距、缩进、间距等）
- 批量操作：遍历、反序、重排
- 特殊段落识别：空段落、标题、列表项、表格内段落
- 编号列表（ListParagraphs）操作
- 样式应用与读取
- 段落边框与底纹
- 与 Range / Selection 的互操作

本模块是 word-text-operator 的互补模块：
- word-text-operator 的段落格式通过 Range/Selection 操作（影响光标所在段落的格式）
- 本模块直接操作 Paragraphs 集合，支持精确索引/条件查询、批量处理、CRUD 等高级功能
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Optional, Union, List

if TYPE_CHECKING:
    from win32com.client import CDispatch


# ── wdParagraphAlignment ───────────────────────────────────────────────────────
WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_ALIGN_PARAGRAPH_RIGHT = 2
WD_ALIGN_PARAGRAPH_JUSTIFY = 3
WD_ALIGN_PARAGRAPH_DISTRIBUTE = 4

# ── wdLineSpacingType ─────────────────────────────────────────────────────────
WD_LINESPACING_SINGLE = 0
WD_LINESPACING_ONE_POINT_FIVE = 1
WD_LINESPACING_DOUBLE = 2
WD_LINESPACING_AT_LEAST = 3
WD_LINESPACING_EXACT = 4
WD_LINESPACING_MULTIPLE = 5  # 多倍行距：LineSpacing 为磅，12 磅 = 单倍

# ── wdListType ────────────────────────────────────────────────────────────────
WD_LIST_NONE = 0
WD_LIST_BULLET = 1
WD_LIST_NUMBER = 2
WD_LIST_PICTURE_BULLET = 3
WD_LIST_MIXED = -1

# ── wdNumberStyle ────────────────────────────────────────────────────────────
WD_STYLE_NORMAL = -1
WD_STYLE_HEADING_1 = -2
WD_STYLE_HEADING_2 = -3
WD_STYLE_HEADING_3 = -4

# ── wdParagraphControl ───────────────────────────────────────────────────────
WD_TERMINATE_COLUMN = 1
WD_TERMINATE_HARD_RULE = 4
WD_TERMINATE_INLINE_SHAPE = 5
WD_TERMINATE_LIST = 2
WD_TERMINATE_LIST_2 = 6

# ── wdRelocate ───────────────────────────────────────────────────────────────
WD_RELATIVE_VERTICAL_SIZE = 0
WD_RELATIVE_HORIZONTAL_SIZE = 1
WD_MOVE = 0
WD_EXTEND = 1


_ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH_LEFT,
    "center": WD_ALIGN_PARAGRAPH_CENTER,
    "right": WD_ALIGN_PARAGRAPH_RIGHT,
    "justify": WD_ALIGN_PARAGRAPH_JUSTIFY,
    "distribute": WD_ALIGN_PARAGRAPH_DISTRIBUTE,
}

_ALIGN_REVERSE = {v: k for k, v in _ALIGN_MAP.items()}

_LINESPACING_MAP = {
    "single": WD_LINESPACING_SINGLE,
    "one_point_five": WD_LINESPACING_ONE_POINT_FIVE,
    "1.5": WD_LINESPACING_ONE_POINT_FIVE,
    "double": WD_LINESPACING_DOUBLE,
    "at_least": WD_LINESPACING_AT_LEAST,
    "exact": WD_LINESPACING_EXACT,
    "multiple": WD_LINESPACING_MULTIPLE,
}


class ParagraphOperator:
    """
    段落操作封装。

    本类封装了对 Word Paragraphs 集合的直接操作，提供：
    - 按索引访问（支持负数）、按条件查询
    - 段落属性读写（对齐、行距、缩进、间距、样式）
    - 段落增删合并拆分
    - 边框底纹
    - 编号列表操作
    - 与 Range/Selection 互操作
    """

    def __init__(self, word_base):
        """
        Args:
            word_base: WordBase 实例
        """
        self._wb = word_base

    # =========================================================================
    # 基础访问
    # =========================================================================

    def _paragraphs(self) -> "CDispatch":
        """返回文档的 Paragraphs 集合。"""
        return self._wb.document.Paragraphs

    def _list_paras(self) -> "CDispatch":
        """返回文档的 ListParagraphs 集合（带编号的段落）。"""
        return self._wb.document.ListParagraphs

    def count(self) -> int:
        """返回文档中的段落总数。"""
        return self._paragraphs().Count

    def get(self, index: int) -> "CDispatch":
        """
        按索引获取段落（1-based，Word 原生索引）。

        支持负数索引（-1 = 最后一段）。

        Args:
            index: 段落编号（从 1 开始），或负数（-1 为最后一段）

        Returns:
            Paragraph COM 对象
        """
        paras = self._paragraphs()
        total = paras.Count
        if index < 0:
            index = total + index + 1
        if index < 1 or index > total:
            raise IndexError(f"段落索引 {index} 超出范围（1~{total}）")
        return paras(index)

    def all(self) -> List["CDispatch"]:
        """返回所有段落的列表。"""
        paras = self._paragraphs()
        return [paras(i) for i in range(1, paras.Count + 1)]

    def first(self) -> "CDispatch":
        """返回第一段。"""
        return self.get(1)

    def last(self) -> "CDispatch":
        """返回最后一段。"""
        return self.get(-1)

    def range(self, start: int, end: int) -> List["CDispatch"]:
        """
        获取 [start, end] 范围内的段落（含首尾）。

        Args:
            start: 起始索引（1-based）
            end: 结束索引（1-based，负数支持）
        """
        paras = self._paragraphs()
        total = paras.Count
        if start < 0:
            start = total + start + 1
        if end < 0:
            end = total + end + 1
        return [paras(i) for i in range(max(1, start), min(total, end) + 1)]

    def at_range(self, rng: "CDispatch") -> List["CDispatch"]:
        """
        获取 Range 范围内包含的所有段落。

        Args:
            rng: Word Range 对象

        Returns:
            段落对象列表
        """
        return [rng.Paragraphs(i) for i in range(1, rng.Paragraphs.Count + 1)]

    # =========================================================================
    # 段落属性读取
    # =========================================================================

    def get_text(self, para: "CDispatch") -> str:
        """读取段落的纯文本内容。"""
        return para.Range.Text.rstrip("\r\x07")

    def get_length(self, para: "CDispatch") -> int:
        """返回段落文本的字符数。"""
        return para.Range.End - para.Range.Start

    def get_index(self, para: "CDispatch") -> int:
        """返回段落在文档中的编号（从 1 开始）。"""
        try:
            idx = int(para.Index)
            if idx >= 1:
                return idx
        except Exception:
            pass
        # Selection.Paragraphs(1) 等部分 Paragraph 在 pywin32 下无 Index，用 Range.Start 匹配
        start = int(para.Range.Start)
        paras = self._paragraphs()
        total = int(paras.Count)
        for i in range(1, total + 1):
            if int(paras(i).Range.Start) == start:
                return i
        raise ValueError(f"无法在文档中定位该段落（Range.Start={start}）")

    def get_style_name(self, para: "CDispatch") -> str:
        """返回段落使用的样式名称。"""
        return para.Style.NameLocal

    def get_style_wd_name(self, para: "CDispatch") -> int:
        """返回段落样式的 Word 内置编号。"""
        return para.Style.Type

    def get_outline_level(self, para: "CDispatch") -> int:
        """返回段落的大纲级别（0=正文，1~9=标题级别）。"""
        return para.OutlineLevel

    def is_heading(self, para: "CDispatch") -> bool:
        """判断段落是否为标题样式。"""
        style = para.Style.NameLocal.lower()
        return "heading" in style or "标题" in style

    def is_empty(self, para: "CDispatch") -> bool:
        """判断段落是否为空（仅含段落标记）。"""
        text = para.Range.Text
        return text == "\r" or text == "\x07"

    def is_list_item(self, para: "CDispatch") -> bool:
        """判断段落是否属于编号列表。"""
        return para.Range.ListFormat.ListType != WD_LIST_NONE

    def is_in_table(self, para: "CDispatch") -> bool:
        """判断段落是否在表格内。"""
        return para.Range.Information(12)  # wdWithInTable = 12

    def get_format_info(self, para: "CDispatch") -> dict:
        """
        读取段落的完整格式属性。

        Returns:
            包含对齐、行距、缩进、间距等属性的字典
        """
        pf = para.Range.ParagraphFormat
        return {
            "alignment": pf.Alignment,
            "alignment_name": _ALIGN_REVERSE.get(pf.Alignment, "unknown"),
            "line_spacing": pf.LineSpacing,
            "line_spacing_rule": pf.LineSpacingRule,
            "space_before": pf.SpaceBefore,
            "space_after": pf.SpaceAfter,
            "left_indent": pf.LeftIndent,
            "right_indent": pf.RightIndent,
            "first_line_indent": pf.FirstLineIndent,
            "outline_level": pf.OutlineLevel,
            "widow_control": pf.WidowControl,
            "keep_together": pf.KeepTogether,
            "keep_with_next": pf.KeepWithNext,
            "page_break_before": pf.PageBreakBefore,
            "style_name": para.Style.NameLocal,
        }

    # =========================================================================
    # 段落属性写入
    # =========================================================================

    def set_alignment(self, para: "CDispatch", align: Union[str, int]):
        """
        设置段落对齐方式。

        Args:
            para: 段落对象
            align: "left" | "center" | "right" | "justify" | "distribute" 或常量值
        """
        if isinstance(align, str):
            align = _ALIGN_MAP.get(align.lower(), WD_ALIGN_PARAGRAPH_LEFT)
        para.Range.ParagraphFormat.Alignment = align

    def set_line_spacing(
        self,
        para: "CDispatch",
        value: Optional[float] = None,
        rule: Union[str, int, None] = None,
    ):
        """
        设置行间距。

        Args:
            para: 段落对象
            value: 行距值
                - 1.0/single: 固定值，rule=single
                - 1.5: 固定值，rule=1.5
                - 2.0/double: 固定值，rule=double
                - 数值（磅值）: 配合 rule="at_least" 或 "exact"
            rule: "single" | "1.5" | "double" | "at_least" | "exact" 或常量
        """
        pf = para.Range.ParagraphFormat
        if rule is None:
            rule = "single"
        if isinstance(rule, str):
            rule = _LINESPACING_MAP.get(rule.lower().strip(), WD_LINESPACING_SINGLE)
        else:
            rule = int(rule)

        if value is not None and rule == WD_LINESPACING_SINGLE:
            v0 = float(value)
            if abs(v0 - 1.5) < 0.01:
                rule = WD_LINESPACING_ONE_POINT_FIVE
            elif abs(v0 - 2.0) < 0.01:
                rule = WD_LINESPACING_DOUBLE

        pf.LineSpacingRule = rule
        if value is None:
            return

        v = float(value)
        # 多倍行距 (wdLineSpaceMultiple=5)：Word 用「磅」表示，12 磅 = 单倍；直接写 1.5 会变成 1.5/12≈0.13 倍
        if rule == WD_LINESPACING_MULTIPLE:
            if 0.25 <= v <= 10.0:
                pf.LineSpacing = v * 12.0
            else:
                pf.LineSpacing = v
            return

        # 单倍 / 1.5 倍 / 双倍：由 LineSpacingRule 决定，勿再把 1.5、2 当作磅写入（否则会异常）
        if rule in (
            WD_LINESPACING_SINGLE,
            WD_LINESPACING_ONE_POINT_FIVE,
            WD_LINESPACING_DOUBLE,
        ):
            if 0.5 <= v <= 3.0:
                return
            pf.LineSpacing = v
            return

        if rule in (WD_LINESPACING_AT_LEAST, WD_LINESPACING_EXACT):
            pf.LineSpacing = v
            return

        pf.LineSpacing = v

    def set_space_before(self, para: "CDispatch", points: float):
        """设置段前间距（磅值）。"""
        para.Range.ParagraphFormat.SpaceBefore = points

    def set_space_after(self, para: "CDispatch", points: float):
        """设置段后间距（磅值）。"""
        para.Range.ParagraphFormat.SpaceAfter = points

    def set_indent_left(
        self,
        para: "CDispatch",
        characters: Optional[float] = None,
        cm: Optional[float] = None,
    ):
        """
        设置段落左缩进。

        Args:
            para: 段落对象
            characters: 缩进字符数（使用 Word 字符单位，非磅值）
            cm: 厘米数（优先级高于 characters）
        """
        pf = para.Range.ParagraphFormat
        app = para.Range.Application
        if cm is not None:
            pf.LeftIndent = app.CentimetersToPoints(float(cm))
        elif characters is not None:
            n = float(characters)
            try:
                pf.CharacterUnitLeftIndent = int(round(n))
            except Exception:
                # 旧版或异常：约 0.37cm/字符（与 Word「字符」标尺一致）
                pf.LeftIndent = app.CentimetersToPoints(0.37 * n)

    def set_indent_right(
        self,
        para: "CDispatch",
        characters: Optional[float] = None,
        cm: Optional[float] = None,
    ):
        """设置段落右缩进。"""
        pf = para.Range.ParagraphFormat
        app = para.Range.Application
        if cm is not None:
            pf.RightIndent = app.CentimetersToPoints(float(cm))
        elif characters is not None:
            n = float(characters)
            try:
                pf.CharacterUnitRightIndent = int(round(n))
            except Exception:
                pf.RightIndent = app.CentimetersToPoints(0.37 * n)

    def set_first_line_indent(
        self,
        para: "CDispatch",
        characters: Optional[float] = None,
        cm: Optional[float] = None,
    ):
        """
        设置首行缩进。传负值向左缩进（悬挂缩进）。

        Args:
            para: 段落对象
            characters: 缩进字符数（使用 Word 字符单位；勿直接写磅，否则会几乎看不见）
            cm: 厘米数（优先级高于 characters）
        """
        pf = para.Range.ParagraphFormat
        app = para.Range.Application
        if cm is not None:
            pf.FirstLineIndent = app.CentimetersToPoints(float(cm))
        elif characters is not None:
            n = float(characters)
            try:
                pf.CharacterUnitFirstLineIndent = int(round(n))
            except Exception:
                # 常见「2 字符 ≈ 0.74cm」
                pf.FirstLineIndent = app.CentimetersToPoints(0.37 * n)

    def set_hanging_indent(self, para: "CDispatch", characters: float):
        """
        设置悬挂缩进（首行外的其他行左缩进）。

        Args:
            para: 段落对象
            characters: 悬挂缩进的字符数
        """
        pf = para.Range.ParagraphFormat
        app = para.Range.Application
        n = float(characters)
        try:
            # Word：负的 CharacterUnitFirstLineIndent 表示悬挂缩进（字符单位）
            pf.CharacterUnitFirstLineIndent = -int(round(abs(n)))
        except Exception:
            pts = app.CentimetersToPoints(0.37 * abs(n))
            pf.FirstLineIndent = -abs(pts)

    def set_outline_level(self, para: "CDispatch", level: int):
        """
        设置大纲级别。

        Args:
            para: 段落对象
            level: 0=正文，1~9=标题级别
        """
        para.Range.ParagraphFormat.OutlineLevel = level

    def set_keep_together(self, para: "CDispatch", on: bool = True):
        """设置段内不分页。"""
        para.Range.ParagraphFormat.KeepTogether = -1 if on else 0

    def set_keep_with_next(self, para: "CDispatch", on: bool = True):
        """设置与下段同页。"""
        para.Range.ParagraphFormat.KeepWithNext = -1 if on else 0

    def set_page_break_before(self, para: "CDispatch", on: bool = True):
        """设置段前分页。"""
        para.Range.ParagraphFormat.PageBreakBefore = -1 if on else 0

    def set_widow_control(self, para: "CDispatch", on: bool = True):
        """设置孤行控制。"""
        para.Range.ParagraphFormat.WidowControl = -1 if on else 0

    def set_style(self, para: "CDispatch", style_name: str):
        """
        应用样式到段落。

        Args:
            para: 段落对象
            style_name: 样式名称（如 "正文"、"标题 1"）
        """
        para.Style = self._wb.document.Styles(style_name)

    def reset_format(self, para: "CDispatch"):
        """将段落格式恢复为默认样式（Normal）的格式。"""
        para.Format = ""

    # =========================================================================
    # 边框与底纹（段落级）
    # =========================================================================

    def set_border(
        self,
        para: "CDispatch",
        side: str = "bottom",
        line_style: int = 1,
        line_width: int = 4,
        color: Union[str, int] = 0x000000,
        space: float = 6.0,
    ):
        """
        给段落指定边添加边框。

        Args:
            para: 段落对象
            side: "top" | "bottom" | "left" | "right"
            line_style: 线条样式（0=无, 1=单线, 2=双线...）
            line_width: 线条宽度（4=0.75磅, 6=1磅, 8=1.5磅...）
            color: 边框颜色（整数或颜色名）
            space: 边框与文字间距（磅值）
        """
        if isinstance(color, str):
            color = self._resolve_color(color)
        rng = para.Range
        rng.Borders.Enable = 1
        side_cap = side.capitalize()
        border = getattr(rng.Borders, f"wdBorder{'{'}{side_cap}{'}'}")
        border.LineStyle = line_style
        border.LineWidth = line_width
        border.Color = color
        border.Space = space

    def clear_border(self, para: "CDispatch"):
        """清除段落的所有边框。"""
        para.Range.Borders.Enable = 0

    def set_shading(
        self,
        para: "CDispatch",
        fill_color: Union[str, int] = 0xFFFF00,
        texture: int = 0,
    ):
        """
        设置段落底纹（背景填充色）。

        Args:
            para: 段落对象
            fill_color: 填充背景色
            texture: 底纹纹理（0=无, 1=横线, 2=竖线...）
        """
        if isinstance(fill_color, str):
            fill_color = self._resolve_color(fill_color)
        rng = para.Range
        rng.Shading.Texture = texture
        rng.Shading.BackgroundPatternColor = fill_color

    def clear_shading(self, para: "CDispatch"):
        """清除段落的底纹。"""
        rng = para.Range
        rng.Shading.Texture = 0
        rng.Shading.BackgroundPatternColor = 0xFFFFFFFF

    def _resolve_color(self, color: Union[str, int]) -> int:
        """将颜色名或整数转为 Word 颜色值。"""
        color_map = {
            "black": 0x000000,
            "white": 0xFFFFFF,
            "red": 0xFF0000,
            "green": 0x00FF00,
            "blue": 0x0000FF,
            "yellow": 0xFFFF00,
            "cyan": 0x00FFFF,
            "magenta": 0xFF00FF,
            "gray": 0x808080,
        }
        if isinstance(color, str):
            if color.startswith("#"):
                return int(color[1:], 16)
            return color_map.get(color.lower(), 0x000000)
        return color

    # =========================================================================
    # 编号列表操作
    # =========================================================================

    def list_count(self) -> int:
        """返回文档中编号段落的数量。"""
        return self._list_paras().Count

    def list_paragraphs(self) -> List["CDispatch"]:
        """返回所有编号段落对象列表。"""
        lps = self._list_paras()
        return [lps(i) for i in range(1, lps.Count + 1)]

    def is_list_paragraph(self, para: "CDispatch") -> bool:
        """判断段落是否在 ListParagraphs 集合中。"""
        lp = self._list_paras()
        if lp.Count == 0:
            return False
        for i in range(1, lp.Count + 1):
            if lp(i).Range.Start == para.Range.Start:
                return True
        return False

    def apply_bullet(self, para: "CDispatch", bullet_type: str = "bullet"):
        """
        将段落转换为项目符号列表项。

        Args:
            para: 段落对象
            bullet_type: 列表类型标识（传给 ListFormat.ListTemplate）
        """
        para.Range.ListFormat.ApplyBullet(NumberStyle=65280)

    def apply_numbering(
        self, para: "CDispatch", number_format: str = "decimal", start_at: int = 1
    ):
        """
        将段落转换为编号列表项。

        Args:
            para: 段落对象
            number_format: 编号格式（"decimal" | "lowerLetter" 等）
            start_at: 起始编号
        """
        lf = para.Range.ListFormat
        try:
            lf.ApplyNumbering(NumberStyle=0)
        except Exception:
            lf.ApplyBullet(NumberStyle=65280)

    def remove_list_format(self, para: "CDispatch"):
        """移除段落的列表格式（还原为普通段落）。"""
        para.Range.ListFormat.ListType = WD_LIST_NONE

    def get_list_level(self, para: "CDispatch") -> int:
        """返回段落在列表中的级别（1-based）。"""
        return para.Range.ListFormat.ListLevelNumber + 1

    def set_list_level(self, para: "CDispatch", level: int):
        """
        设置段落在列表中的级别。

        Args:
            para: 段落对象
            level: 级别（1-based）
        """
        para.Range.ListFormat.ListLevelNumber = max(1, level) - 1

    def get_list_number(self, para: "CDispatch") -> Optional[int]:
        """
        获取段落的当前编号值（如 "3." 中的 3）。

        Returns:
            编号数字，未找到返回 None
        """
        try:
            return para.Range.ListFormat.ListValue + 1
        except Exception:
            return None

    # =========================================================================
    # 段落内容操作（CRUD）
    # =========================================================================

    def set_text(self, para: "CDispatch", text: str):
        """
        替换段落内容（保留段落格式）。

        Args:
            para: 段落对象
            text: 新文本
        """
        para.Range.Text = text

    def insert_text_before(self, para: "CDispatch", text: str) -> "CDispatch":
        """
        在段落开头插入文本。

        Returns:
            段落 Range（插入后内容不变）
        """
        rng = para.Range
        rng.InsertBefore(text)
        return rng

    def insert_text_after(self, para: "CDispatch", text: str) -> "CDispatch":
        """
        在段落末尾（段落标记之前）插入文本。

        Returns:
            段落 Range
        """
        rng = para.Range
        rng.Collapse(Direction=1)  # wdCollapseEnd
        rng.InsertBefore(text)
        rng.MoveEnd(Unit=4, Count=-1)  # wdParagraph，往回合并
        return rng

    def delete_paragraph(self, para: "CDispatch"):
        """
        删除整个段落（内容 + 段落标记）。

        慎用：会合并相邻段落。
        """
        para.Range.Delete()

    def clear_paragraph(self, para: "CDispatch"):
        """清空段落内容（保留段落标记）。"""
        text = para.Range.Text
        if len(text) > 0:
            para.Range.MoveEnd(Unit=1, Count=-1)  # 不含段落标记
            para.Range.Text = ""

    def add_paragraph_after(self, para: "CDispatch") -> "CDispatch":
        """
        在当前段落之后插入新段落。

        Returns:
            新段落对象
        """
        rng = para.Range
        rng.Collapse(Direction=1)  # wdCollapseEnd
        rng.InsertParagraph()
        new_rng = rng.Duplicate
        new_rng.MoveStart(Unit=1, Count=1)  # 移到新段落开头
        return self._wb.document.Paragraphs(new_rng.Start)

    def add_paragraph_before(self, para: "CDispatch") -> "CDispatch":
        """
        在当前段落之前插入新段落。

        Returns:
            新段落对象
        """
        rng = para.Range
        rng.Collapse(Direction=0)  # wdCollapseStart
        rng.InsertParagraph()
        rng.MoveEnd(Unit=4, Count=-1)  # wdParagraph，往回合并
        new_rng = rng.Duplicate
        new_rng.MoveStart(Unit=4, Count=-1)
        return self._wb.document.Paragraphs(new_rng.Start)

    def add_empty_paragraph_after(self, para: "CDispatch") -> "CDispatch":
        """在当前段落之后插入空段落。"""
        rng = para.Range
        rng.Collapse(Direction=1)
        rng.InsertParagraph()
        return para

    def merge_with_next(self, para: "CDispatch") -> bool:
        """
        将当前段落与下一段合并。

        Returns:
            是否成功（最后一段无法与下一段合并）
        """
        try:
            next_para = para.Next()
            if next_para is None:
                return False
            next_para.Range.MoveStart(Unit=1, Count=1)  # 跳过下一段的段落标记
            para.Range.End = next_para.Range.End
            next_para.Range.Delete()
            return True
        except Exception:
            return False

    def merge_with_previous(self, para: "CDispatch") -> bool:
        """
        将当前段落与上一段合并。

        Returns:
            是否成功（第一段无法与上一段合并）
        """
        try:
            prev_para = para.Previous()
            if prev_para is None:
                return False
            prev_para.Range.End = para.Range.End
            para.Range.Delete()
            return True
        except Exception:
            return False

    def split_paragraph(self, para: "CDispatch", separator: str = "\t") -> List["CDispatch"]:
        """
        按分隔符将一个段落拆分为多个段落。

        Args:
            para: 段落对象
            separator: 分隔符（默认 Tab），也可以是其他字符

        Returns:
            拆分后的段落列表
        """
        text = self.get_text(para)
        if separator not in text:
            return [para]

        parts = text.split(separator)
        results = []
        self.set_text(para, parts[0])
        results.append(para)

        for part in parts[1:]:
            new_para = self.add_paragraph_after(para)
            self.set_text(new_para, part)
            results.append(new_para)
            para = new_para

        return results

    # =========================================================================
    # 批量操作
    # =========================================================================

    def find_by_text(
        self, text: str, whole_word: bool = False, match_case: bool = False
    ) -> List["CDispatch"]:
        """
        查找包含指定文本的所有段落。

        Args:
            text: 要查找的文本
            whole_word: 全字匹配
            match_case: 区分大小写

        Returns:
            匹配段落列表
        """
        results = []
        for para in self.all():
            para_text = self.get_text(para)
            if match_case:
                found = text in para_text
            else:
                found = text.lower() in para_text.lower()
            if found and whole_word:
                import re
                pattern = r"\b" + re.escape(text) + r"\b"
                found = bool(re.search(pattern, para_text if match_case else para_text.lower()))
            if found:
                results.append(para)
        return results

    def find_empty_paragraphs(self) -> List["CDispatch"]:
        """返回文档中所有空段落。"""
        return [para for para in self.all() if self.is_empty(para)]

    def find_headings(self) -> List["CDispatch"]:
        """返回文档中所有标题段落。"""
        return [para for para in self.all() if self.is_heading(para)]

    def find_headings_by_level(self, level: int) -> List["CDispatch"]:
        """
        返回指定大纲级别的标题段落。

        Args:
            level: 1~9，对应 Word 的大纲级别
        """
        return [para for para in self.all() if self.get_outline_level(para) == level]

    def find_list_paragraphs(self) -> List["CDispatch"]:
        """返回所有编号/项目符号列表段落。"""
        return [para for para in self.all() if self.is_list_item(para)]

    def apply_format_to_all(
        self,
        align: Union[str, int, None] = None,
        line_spacing: Optional[float] = None,
        line_spacing_rule: Union[str, int, None] = None,
        space_before: Optional[float] = None,
        space_after: Optional[float] = None,
        indent_left: Optional[float] = None,
        indent_right: Optional[float] = None,
        first_line_indent: Optional[float] = None,
    ) -> int:
        """
        批量设置所有段落的格式属性。

        Returns:
            实际修改的段落数量
        """
        count = 0
        for para in self.all():
            pf = para.Range.ParagraphFormat
            if align is not None:
                self.set_alignment(para, align)
            if line_spacing is not None:
                self.set_line_spacing(para, line_spacing, line_spacing_rule)
            if space_before is not None:
                pf.SpaceBefore = space_before
            if space_after is not None:
                pf.SpaceAfter = space_after
            if indent_left is not None:
                self.set_indent_left(para, characters=indent_left)
            if indent_right is not None:
                self.set_indent_right(para, characters=indent_right)
            if first_line_indent is not None:
                self.set_first_line_indent(para, characters=first_line_indent)
            count += 1
        return count

    def reverse_order(self, start: int = 1, end: Optional[int] = None) -> List["CDispatch"]:
        """
        将指定范围内的段落顺序反转（原地反序排列文本内容）。

        通过移动段落 Range 实现，不改变段落数量。

        Args:
            start: 起始索引（1-based）
            end: 结束索引（None=到末尾）

        Returns:
            反转后的段落列表
        """
        paras = self.range(start, end) if end else self.range(start, self.count())
        if len(paras) <= 1:
            return paras

        texts = [self.get_text(p) for p in paras]
        texts.reverse()

        for para, text in zip(paras, texts):
            self.set_text(para, text)

        return paras

    # =========================================================================
    # Range / Selection 互操作
    # =========================================================================

    def get_paragraph_at_selection(self) -> "CDispatch":
        """获取当前 Selection 所在段落的段落对象。"""
        sel = self._wb.selection
        return sel.Paragraphs(1)

    def get_paragraph_at_range(self, rng: "CDispatch") -> "CDispatch":
        """
        获取指定 Range 所在段落的段落对象。

        Args:
            rng: Word Range 对象

        Returns:
            段落对象
        """
        return rng.Paragraphs(1)

    def select_paragraph(self, para: "CDispatch"):
        """选中整个段落（影响界面 Selection）。"""
        para.Range.Select()

    def select_range_of_paragraphs(self, start: int, end: int):
        """
        选中指定范围内的所有段落（影响界面 Selection）。

        Args:
            start: 起始段落索引（1-based）
            end: 结束段落索引（1-based）
        """
        p1 = self.get(start)
        p2 = self.get(end)
        rng = self._wb.document.Range(p1.Range.Start, p2.Range.End)
        rng.Select()

    def collapse_all(self):
        """
        将文档中所有段落合并为一个（保留内容，去除多余段落标记）。
        慎用：会破坏文档结构。
        """
        doc = self._wb.document
        doc.Content.Paragraphs(1).Range.End = doc.Content.End

    # =========================================================================
    # 读取文档结构
    # =========================================================================

    def get_outline_summary(self) -> List[dict]:
        """
        返回文档的大纲结构摘要（用于生成目录）。

        Returns:
            包含 {level, text, index} 的字典列表
        """
        summary = []
        for para in self.all():
            if self.get_outline_level(para) <= 9:
                summary.append({
                    "level": self.get_outline_level(para),
                    "text": self.get_text(para).strip(),
                    "index": self.get_index(para),
                })
        return summary

    def get_document_structure(self) -> List[dict]:
        """
        返回文档的段落结构信息。

        Returns:
            每段的结构信息列表
        """
        return [
            {
                "index": self.get_index(para),
                "text": self.get_text(para)[:60],
                "style": self.get_style_name(para),
                "level": self.get_outline_level(para),
                "is_heading": self.is_heading(para),
                "is_empty": self.is_empty(para),
                "is_list": self.is_list_item(para),
            }
            for para in self.all()
        ]
