# -*- coding: utf-8 -*-
"""
Word 文本格式模块

提供 Range/Selection 的格式操作能力（Font / ParagraphFormat / Borders / Shading）。
由于很多操作 Range 和 Selection 共用一套接口，本模块尽量统一封装。
"""

from __future__ import annotations
from typing import TYPE_CHECKING, Optional, Union, List

if TYPE_CHECKING:
    from win32com.client import CDispatch


WD_COLOR = {
    "black": 0x000000,
    "white": 0xFFFFFF,
    "red": 0xFF0000,
    "green": 0x00FF00,
    "blue": 0x0000FF,
    "yellow": 0xFFFF00,
    "cyan": 0x00FFFF,
    "magenta": 0xFF00FF,
    "gray": 0x808080,
    "dark_red": 0x800000,
    "dark_green": 0x008000,
    "dark_blue": 0x000080,
    "dark_yellow": 0x808000,
    "auto": 0xFFFFFFFF,
}

WD_UNDERLINE_TYPES = {
    "none": 0,
    "single": 1,
    "words": 2,
    "double": 3,
    "thick": 4,
    "dotted": 5,
    "dotted_heavy": 6,
    "dash": 7,
    "dash_heavy": 8,
    "long_dash": 9,
    "long_dash_heavy": 10,
    "dash_dot": 11,
    "dash_dot_heavy": 12,
    "dash_dot_dot": 13,
    "dash_dot_dot_heavy": 14,
    "wavy": 15,
    "wavy_heavy": 16,
    "wavy_dotted": 17,
}

WD_HIGHLIGHT_COLOR = {
    "none": 0,
    "black": 1,
    "blue": 2,
    "cyan": 3,
    "green": 4,
    "magenta": 5,
    "red": 6,
    "yellow": 7,
    "white": 8,
    "dark_blue": 9,
    "dark_cyan": 10,
    "dark_green": 11,
    "dark_magenta": 12,
    "dark_red": 13,
    "dark_yellow": 14,
    "dark_gray": 15,
    "light_gray": 16,
}

WD_ALIGN_PARAGRAPH = {
    "left": 0,
    "center": 1,
    "right": 2,
    "justify": 3,
    "distribute": 4,
}

WD_LINE_SPACING_RULE = {
    "single": 0,
    "1.5": 1,
    "double": 2,
    "at_least": 3,
    "exact": 4,
}

WD_PANOSE_LETTERFORM = {
    "normal": 0,
    "old_style": 2,
    "modern": 4,
    "modern_width": 6,
    "traditional": 8,
    "modern_weight": 10,
    "letterpress": 12,
}


class TextFormatter:
    """字符 + 段落格式操作封装。"""

    def __init__(self, word_base):
        self._wb = word_base
        self._color_map = WD_COLOR
        self._underline_map = WD_UNDERLINE_TYPES
        self._highlight_map = WD_HIGHLIGHT_COLOR

    # ========================================================================
    # 公共工具
    # ========================================================================

    def _resolve_color(self, color: Union[str, int]) -> int:
        """将颜色名或整数转为 Word 颜色值。"""
        if isinstance(color, str):
            return self._color_map.get(color.lower(), int(color, 16) if color.startswith("#") else int(color))
        return color

    def _resolve_underline(self, underline: Union[str, int]) -> int:
        """将下划线名称转为常量值。"""
        if isinstance(underline, str):
            return self._underline_map.get(underline.lower(), 1)
        return underline

    def _resolve_paragraph_align(self, align: Union[str, int]) -> int:
        """将段落对齐名称转为常量值。"""
        if isinstance(align, str):
            return self._align_map.get(align.lower(), 0)
        return align

    def _font(self, rng: "CDispatch"):
        """获取 Range 的 Font 对象。"""
        return rng.Font

    def _para_fmt(self, rng: "CDispatch"):
        """获取 Range 的 ParagraphFormat 对象。"""
        return rng.ParagraphFormat

    # ========================================================================
    # 字体：名称 / 大小 / 样式
    # ========================================================================

    def set_font_name(self, rng: "CDispatch", name: str):
        """设置字体名称（如 "宋体"、"Arial"）。"""
        rng.Font.Name = name

    def set_font_name_ascii(self, rng: "CDispatch", name: str):
        """设置 ASCII 字符字体。"""
        rng.Font.NameAscii = name

    def set_font_name_east_asia(self, rng: "CDispatch", name: str):
        """设置东亚字体（中日韩）。"""
        rng.Font.NameEastAsia = name

    def set_font_name_bi(self, rng: "CDispatch", name: str):
        """设置复杂文字字体。"""
        rng.Font.NameBi = name

    def get_font_name(self, rng: "CDispatch") -> str:
        """获取字体名称。"""
        return rng.Font.Name

    def set_font_size(self, rng: "CDispatch", size: float):
        """设置字号（磅值）。"""
        rng.Font.Size = size

    def set_font_size_half(self, rng: "CDispatch", half_points: int):
        """设置半磅字号（Word 内部存储方式）。"""
        rng.Font.Size = half_points / 2.0

    def get_font_size(self, rng: "CDispatch") -> float:
        """获取字号。"""
        return rng.Font.Size

    def set_bold(self, rng: "CDispatch", bold: bool = True):
        """设置加粗。"""
        rng.Font.Bold = -1 if bold else 0

    def set_italic(self, rng: "CDispatch", italic: bool = True):
        """设置斜体。"""
        rng.Font.Italic = -1 if italic else 0

    def set_underline(self, rng: "CDispatch", underline: Union[str, int] = "single"):
        """
        设置下划线。

        Args:
            underline: 下划线类型，如 "single"、"double"、"words"，
                       或 WD_UNDERLINE_TYPES 中的常量值。
        """
        rng.Font.Underline = self._resolve_underline(underline)

    def set_underline_color(self, rng: "CDispatch", color: Union[str, int]):
        """设置下划线颜色。"""
        rng.Font.UnderlineColor = self._resolve_color(color)

    def set_strike_through(self, rng: "CDispatch", strike: bool = True):
        """设置删除线。"""
        rng.Font.StrikeThrough = -1 if strike else 0

    def set_double_strike_through(self, rng: "CDispatch", strike: bool = True):
        """设置双删除线。"""
        rng.Font.DoubleStrikeThrough = -1 if strike else 0

    def set_superscript(self, rng: "CDispatch", on: bool = True):
        """设置上标。"""
        rng.Font.Superscript = -1 if on else 0

    def set_subscript(self, rng: "CDispatch", on: bool = True):
        """设置下标。"""
        rng.Font.Subscript = -1 if on else 0

    def set_small_caps(self, rng: "CDispatch", on: bool = True):
        """设置小型大写字母。"""
        rng.Font.SmallCaps = -1 if on else 0

    def set_all_caps(self, rng: "CDispatch", on: bool = True):
        """设置全大写（忽略实际大小写）。"""
        rng.Font.AllCaps = -1 if on else 0

    def set_hidden(self, rng: "CDispatch", hidden: bool = True):
        """设置隐藏文字。"""
        rng.Font.Hidden = -1 if hidden else 0

    def set_emphasis_mark(self, rng: "CDispatch", mark: int = 1):
        """
        设置着重号样式。

        Args:
            mark: 0=无, 1=点号, 2=圆圈
        """
        rng.Font.EmphasisMark = mark

    # ========================================================================
    # 字体：颜色 / 底纹 / 边框
    # ========================================================================

    def set_font_color(self, rng: "CDispatch", color: Union[str, int]):
        """设置字体颜色。"""
        rng.Font.Color = self._resolve_color(color)

    def get_font_color(self, rng: "CDispatch") -> int:
        """获取字体颜色值。"""
        return rng.Font.Color

    def set_highlight(self, rng: "CDispatch", color: Union[str, int] = "yellow"):
        """
        设置文字高亮（背景色）。

        Args:
            color: 高亮颜色名称或 wdHighlightColor 常量
        """
        if isinstance(color, str):
            color = self._highlight_map.get(color.lower(), 7)
        rng.Font.Highlight = color

    def clear_highlight(self, rng: "CDispatch"):
        """清除文字高亮。"""
        rng.Font.Highlight = 0

    def set_backstyle_italic(self, rng: "CDispatch", italic: bool = True):
        """设置 BackslashItalic 样式（视觉上略有倾斜）。"""
        rng.Font.BackSlashItalic = -1 if italic else 0

    def set_character_width(self, rng: "CDispatch", width: int):
        """
        设置字符宽度百分比（100 = 正常）。

        Args:
            width: 1-600 之间的整数
        """
        rng.Font.Width = width

    def set_spacing_scale(self, rng: "CDispatch", scale: int):
        """
        设置字符间距百分比（100 = 正常）。

        Args:
            scale: 1-600 之间的整数
        """
        rng.Font.Scaling = scale

    # ========================================================================
    # 字符间距 (Spacing)
    # ========================================================================

    def set_spacing_before(self, rng: "CDispatch", points: float):
        """
        设置字符前间距（磅值）。
        与字符间距不同：Spacing 在字符之间，Position 垂直偏移。
        """
        rng.Font.Spacing = points

    def get_spacing(self, rng: "CDispatch") -> float:
        """获取字符间距（磅值）。"""
        return rng.Font.Spacing

    def set_expansion(self, rng: "CDispatch", points: float):
        """设置字符宽度扩展（磅值，>0 展开，<0 压缩）。"""
        rng.Font.Expansion = points

    def set_vertical_position(self, rng: "CDispatch", points: float):
        """
        设置字符垂直位置偏移（磅值）。
        正值上移，负值下移。
        """
        rng.Font.Position = points

    def set_kerning(self, rng: "CDispatch", min_points: float):
        """
        设置最小字距调整阈值（磅值）。
        字号大于此值时自动调整字符间距。
        """
        rng.Font.Kerning = min_points

    # ========================================================================
    # 段落格式
    # ========================================================================

    def set_alignment(self, rng: "CDispatch", align: Union[str, int]):
        """
        设置段落对齐方式。

        Args:
            align: "left" | "center" | "right" | "justify" | "distribute"
                  或 WD_ALIGN_PARAGRAPH 常量值
        """
        if isinstance(align, str):
            align = WD_ALIGN_PARAGRAPH.get(align.lower(), 0)
        rng.ParagraphFormat.Alignment = align

    def get_alignment(self, rng: "CDispatch") -> int:
        """获取段落对齐方式常量值。"""
        return rng.ParagraphFormat.Alignment

    def set_line_spacing(
        self,
        rng: "CDispatch",
        value: float,
        rule: Union[str, int] = "single",
    ):
        """
        设置行间距。

        Args:
            value: 行距数值
                - "single"/0/None: 使用 LineSpacingRule = single
                - 数字（如 1.5、2）配合 rule="1.5"/"double"
                - 磅值数字（如 20）配合 rule="at_least"/"exact"
            rule: 行距规则
        """
        if isinstance(rule, str):
            rule = WD_LINE_SPACING_RULE.get(rule.lower(), 0)
        rng.ParagraphFormat.LineSpacingRule = rule
        rng.ParagraphFormat.LineSpacing = value

    def set_line_spacing_rule(
        self, rng: "CDispatch", rule: Union[str, int], value: Optional[float] = None
    ):
        """
        单独设置行距规则。

        Args:
            rule: WD_LINE_SPACING_RULE 常量
            value: 如果 rule=at_least/exact，需要指定具体磅值
        """
        if isinstance(rule, str):
            rule = WD_LINE_SPACING_RULE.get(rule.lower(), 0)
        rng.ParagraphFormat.LineSpacingRule = rule
        if value is not None:
            rng.ParagraphFormat.LineSpacing = value

    def set_space_before_para(self, rng: "CDispatch", points: float):
        """设置段前间距（磅值）。"""
        rng.ParagraphFormat.SpaceBefore = points

    def set_space_after_para(self, rng: "CDispatch", points: float):
        """设置段后间距（磅值）。"""
        rng.ParagraphFormat.SpaceAfter = points

    def get_space_before_para(self, rng: "CDispatch") -> float:
        """获取段前间距。"""
        return rng.ParagraphFormat.SpaceBefore

    def get_space_after_para(self, rng: "CDispatch") -> float:
        """获取段后间距。"""
        return rng.ParagraphFormat.SpaceAfter

    def set_indent_left(
        self, rng: "CDispatch", characters: Optional[float] = None, cm: Optional[float] = None
    ):
        """
        设置段落左缩进。

        Args:
            characters: 缩进字符数（Word 默认单位）
            cm: 厘米数（优先级高于 characters）
        """
        if cm is not None:
            rng.ParagraphFormat.LeftIndent = rng.Application.CentimetersToPoints(cm)
        elif characters is not None:
            rng.ParagraphFormat.LeftIndent = characters

    def set_indent_right(self, rng: "CDispatch", characters: Optional[float] = None, cm: Optional[float] = None):
        """设置段落右缩进。"""
        if cm is not None:
            rng.ParagraphFormat.RightIndent = rng.Application.CentimetersToPoints(cm)
        elif characters is not None:
            rng.ParagraphFormat.RightIndent = characters

    def set_first_line_indent(
        self, rng: "CDispatch", characters: Optional[float] = None, cm: Optional[float] = None
    ):
        """设置首行缩进。传负值向左缩进（悬挂缩进）。"""
        if cm is not None:
            rng.ParagraphFormat.FirstLineIndent = rng.Application.CentimetersToPoints(cm)
        elif characters is not None:
            rng.ParagraphFormat.FirstLineIndent = characters

    def set_hanging_indent(self, rng: "CDispatch", characters: float):
        """设置悬挂缩进（首行外的其他行左缩进）。"""
        rng.ParagraphFormat.FirstLineIndent = rng.ParagraphFormat.LeftIndent - characters

    def set_outline_level(self, rng: "CDispatch", level: int):
        """
        设置大纲级别（用于导航窗格和生成目录）。

        Args:
            level: 0-8，0=正文，1=1级标题，2=2级标题...
        """
        rng.ParagraphFormat.OutlineLevel = level

    def set_keep_together(self, rng: "CDispatch", on: bool = True):
        """设置段内不分页（保持段落完整）。"""
        rng.ParagraphFormat.KeepTogether = -1 if on else 0

    def set_keep_with_next(self, rng: "CDispatch", on: bool = True):
        """设置与下段同页。"""
        rng.ParagraphFormat.KeepWithNext = -1 if on else 0

    def set_page_break_before(self, rng: "CDispatch", on: bool = True):
        """设置段前分页。"""
        rng.ParagraphFormat.PageBreakBefore = -1 if on else 0

    def set_widow_control(self, rng: "CDispatch", on: bool = True):
        """设置孤行控制。"""
        rng.ParagraphFormat.WidowControl = -1 if on else 0

    def set_reading_order(self, rng: "CDispatch", ltr: bool = True):
        """
        设置阅读顺序。

        Args:
            ltr: True=从左到右（中文/英文），False=从右到左（阿拉伯文等）
        """
        rng.ParagraphFormat.ReadingOrder = 0 if ltr else 1

    # ========================================================================
    # Borders（边框）
    # ========================================================================

    def set_border(
        self,
        rng: "CDispatch",
        side: str = "bottom",
        line_style: int = 1,
        line_width: int = 4,
        color: Union[str, int] = 0x000000,
        space: float = 6.0,
    ):
        """
        给 Range 的指定边添加边框。

        Args:
            side: "top" | "bottom" | "left" | "right" | "inside" | "outside"
            line_style: 线条样式（0=无, 1=单线, 2=双线, 3=点线...）
            line_width: 线条宽度（1=0.5磅, 4=0.75磅, 6=1磅, 8=1.5磅, 18=2.25磅）
            color: 边框颜色
            space: 边框与文字间距（磅值）
        """
        color_val = self._resolve_color(color)
        if side in ("top", "bottom", "left", "right", "inside", "outside"):
            border = getattr(rng.Borders, f"wdBorder{'{'}{side.capitalize()}{'}'}")
        else:
            border = rng.Borders(1)  # 默认 bottom

        border.LineStyle = line_style
        border.LineWidth = line_width
        border.Color = color_val
        border.Space = space

    def clear_border(self, rng: "CDispatch", side: str = "all"):
        """清除边框。"""
        if side == "all":
            rng.Borders.Enable = 0
        else:
            border = getattr(rng.Borders, f"wdBorder{'{'}{side.capitalize()}{'}'}", None)
            if border:
                border.LineStyle = 0

    def set_box_border(
        self,
        rng: "CDispatch",
        line_style: int = 1,
        line_width: int = 4,
        color: Union[str, int] = 0x000000,
    ):
        """给 Range 添加四周方框。"""
        rng.Borders.Enable = 1
        for side in ("Top", "Left", "Bottom", "Right"):
            border = getattr(rng.Borders, f"wdBorder{'{'}{side}{'}'}")
            border.LineStyle = line_style
            border.LineWidth = line_width
            border.Color = color_val

    # ========================================================================
    # Shading（底纹 / 背景色）
    # ========================================================================

    def set_shading(
        self,
        rng: "CDispatch",
        fill_color: Union[str, int] = 0xFFFF00,
        texture: int = 0,
        foreground: Union[str, int] = 0x000000,
    ):
        """
        设置文字底纹（背景填充色）。

        Args:
            fill_color: 填充背景色
            texture: 底纹纹理（0=无, 1=横线, 2=竖线...）
            foreground: 纹理前景色
        """
        rng.Shading.Texture = texture
        rng.Shading.ForegroundPatternColor = self._resolve_color(foreground)
        rng.Shading.BackgroundPatternColor = self._resolve_color(fill_color)

    def clear_shading(self, rng: "CDispatch"):
        """清除底纹（背景色）。"""
        rng.Shading.Texture = 0
        rng.Shading.ForegroundPatternColor = 0xFFFFFFFF
        rng.Shading.BackgroundPatternColor = 0xFFFFFFFF

    def set_highlight_colored(self, rng: "CDispatch", fill: Union[str, int], fg: Union[str, int] = 0):
        """设置高亮（类似荧光笔效果）。"""
        rng.Shading.Texture = 100  # wdTextureInverse
        rng.Shading.ForegroundPatternColor = self._resolve_color(fg)
        rng.Shading.BackgroundPatternColor = self._resolve_color(fill)

    # ========================================================================
    # Tab / 制表符
    # ========================================================================

    def add_tab(
        self, rng: "CDispatch", position: float, align: str = "left", leader: str = "none"
    ):
        """
        在 Range 段落末尾添加制表位。

        Args:
            position: 制表位位置（磅值）
            align: "left" | "center" | "right" | "decimal"
            leader: "none" | "dot" | "dash" | "underline" | "heavy"
        """
        align_map = {"left": 0, "center": 1, "right": 2, "decimal": 3}
        leader_map = {"none": 0, "dot": 1, "dash": 2, "underline": 3, "heavy": 4, "middle_dot": 5}
        tab = rng.ParagraphFormat.Tabs.Add(
            Position=position,
            Alignment=align_map.get(align.lower(), 0),
            Leader=leader_map.get(leader.lower(), 0),
        )
        return tab

    def clear_tabs(self, rng: "CDispatch"):
        """清除所有制表位。"""
        rng.ParagraphFormat.Tabs.ClearAll()

    # ========================================================================
    # 格式刷
    # ========================================================================

    def copy_format(self, rng: "CDispatch"):
        """复制 Range 的格式到格式刷。"""
        rng.FormatterContainer = rng.Font

    def paste_format(self, rng: "CDispatch"):
        """将格式刷的格式粘贴到 Range。"""
        rng.FormatterContainer = rng

    # ========================================================================
    # 格式信息读取
    # ========================================================================

    def get_font_info(self, rng: "CDispatch") -> dict:
        """读取 Range 的字体信息。"""
        f = rng.Font
        return {
            "name": f.Name,
            "size": f.Size,
            "bold": f.Bold,
            "italic": f.Italic,
            "underline": f.Underline,
            "color": f.Color,
            "highlight": f.Highlight,
        }

    def get_paragraph_format_info(self, rng: "CDispatch") -> dict:
        """读取 Range 的段落格式信息。"""
        pf = rng.ParagraphFormat
        return {
            "alignment": pf.Alignment,
            "line_spacing": pf.LineSpacing,
            "line_spacing_rule": pf.LineSpacingRule,
            "space_before": pf.SpaceBefore,
            "space_after": pf.SpaceAfter,
            "left_indent": pf.LeftIndent,
            "right_indent": pf.RightIndent,
            "first_line_indent": pf.FirstLineIndent,
            "outline_level": pf.OutlineLevel,
        }

    def get_format_summary(self, rng: "CDispatch") -> str:
        """生成格式信息的人类可读摘要。"""
        fi = self.get_font_info(rng)
        pf = self.get_paragraph_format_info(rng)
        parts = [
            f"字体: {fi['name']}, {fi['size']}pt",
            f"样式: {'粗体' if fi['bold'] else ''}{'斜体' if fi['italic'] else ''}{'下划线' if fi['underline'] else ''}".strip() or "常规",
            f"对齐: {pf['alignment']}",
            f"行距: {pf['line_spacing']}",
        ]
        return " | ".join(p for p in parts if p)
