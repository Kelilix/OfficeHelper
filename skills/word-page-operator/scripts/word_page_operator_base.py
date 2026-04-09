# -*- coding: utf-8 -*-
"""
PageSetup 页面设置操作模块

提供基于 Document.PageSetup / SectionFormat 的页面级操作能力：
- 页边距读取与设置（厘米/英寸）
- 纸张大小读取与设置（预设/自定义）
- 纸张方向（纵向/横向）
- 分栏操作（栏数、栏宽、栏间距）
- 页面背景与边框
- 页面垂直对齐
- 水印操作
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Optional, Union, List

if TYPE_CHECKING:
    from win32com.client import CDispatch


# ── wdPageOrientation ───────────────────────────────────────────────────────────
WD_ORIENT_PORTRAIT = 0
WD_ORIENT_LANDSCAPE = 1

# ── wdTextAlignment ────────────────────────────────────────────────────────────
WD_ALIGN_PAGE_VERTICAL_TOP = 0
WD_ALIGN_PAGE_VERTICAL_CENTER = 1
WD_ALIGN_PAGE_VERTICAL_JUSTIFY = 2
WD_ALIGN_PAGE_VERTICAL_BOTTOM = 3

# ── wdSectionStart ─────────────────────────────────────────────────────────────
WD_SECTION_START_CONTINUOUS = 3
WD_SECTION_START_NEW_PAGE = 0
WD_SECTION_START_EVEN_PAGE = 2
WD_SECTION_START_ODD_PAGE = 1

# ── wdHeaderFooterIndex ────────────────────────────────────────────────────────
WD_HEADER_PRIMARY = -1
WD_HEADER_FIRST = 0
WD_HEADER_EVEN_PAGES = 1

# ── wdBorderType ──────────────────────────────────────────────────────────────
WD_BORDER_TOP = 1
WD_BORDER_BOTTOM = 2
WD_BORDER_LEFT = 3
WD_BORDER_RIGHT = 4
WD_BORDER_HORIZONTAL = 5
WD_BORDER_VERTICAL = 6
WD_BORDER_DIAGONAL_DOWN = 7
WD_BORDER_DIAGONAL_UP = 8
WD_BORDER_ALL = 9

# ── 预设纸张大小（厘米）─────────────────────────────────────────────────────────
PAPER_PRESETS_CM = {
    "A0":     (84.1,  118.9),
    "A1":     (59.4,  84.1),
    "A2":     (42.0,  59.4),
    "A3":     (29.7,  42.0),
    "A4":     (21.0,  29.7),
    "A5":     (14.8,  21.0),
    "A6":     (10.5,  14.8),
    "B4":     (25.0,  35.3),
    "B5":     (17.6,  25.0),
    "Letter": (21.59, 27.94),
    "Legal":  (21.59, 35.56),
    "Tabloid": (27.94, 43.18),
    "Executive": (18.41, 26.67),
}

# ── 预设页边距（厘米）───────────────────────────────────────────────────────────
MARGIN_PRESETS_CM = {
    "normal":  {"top": 2.54, "bottom": 2.54, "left": 3.17, "right": 3.17},
    "narrow":  {"top": 1.27, "bottom": 1.27, "left": 1.27, "right": 1.27},
    "wide":    {"top": 2.54, "bottom": 2.54, "left": 5.08, "right": 5.08},
    "mirrored": {"top": 2.54, "bottom": 2.54, "left": 3.17, "right": 3.17},
}

_ORIENT_MAP = {
    "portrait": WD_ORIENT_PORTRAIT,
    "landscape": WD_ORIENT_LANDSCAPE,
}
_ORIENT_REVERSE = {v: k for k, v in _ORIENT_MAP.items()}

_VERTICAL_ALIGN_MAP = {
    "top": WD_ALIGN_PAGE_VERTICAL_TOP,
    "center": WD_ALIGN_PAGE_VERTICAL_CENTER,
    "justify": WD_ALIGN_PAGE_VERTICAL_JUSTIFY,
    "bottom": WD_ALIGN_PAGE_VERTICAL_BOTTOM,
}

_SECTION_START_MAP = {
    "continuous": WD_SECTION_START_CONTINUOUS,
    "new_page": WD_SECTION_START_NEW_PAGE,
    "even_page": WD_SECTION_START_EVEN_PAGE,
    "odd_page": WD_SECTION_START_ODD_PAGE,
}
_SECTION_START_REVERSE = {v: k for k, v in _SECTION_START_MAP.items()}


class PageSetupOperator:
    """
    页面设置操作封装。

    封装了 PageSetup 对象的各种属性和方法，提供：
    - 页边距的读取与设置（支持厘米/英寸/预设）
    - 纸张大小的读取与设置（支持预设/自定义）
    - 纸张方向（纵向/横向）
    - 分栏设置（栏数、栏宽、栏间距）
    - 页面边框与背景
    - 页面垂直对齐
    - 起始页码设置
    """

    def __init__(self, word_base):
        """
        Args:
            word_base: WordBase 实例（包含 _document, _word_app 等）
        """
        self._wb = word_base

    # =========================================================================
    # 辅助方法
    # =========================================================================

    def _cm_to_points(self, cm: float) -> float:
        """厘米转磅值（1cm ≈ 28.35pt）。"""
        return float(cm) * 28.35

    def _points_to_cm(self, points: float) -> float:
        """磅值转厘米。"""
        return float(points) / 28.35

    def _inch_to_points(self, inch: float) -> float:
        """英寸转磅值（1 inch = 72pt）。"""
        return float(inch) * 72.0

    def _resolve_section(self, index: int) -> "CDispatch":
        """
        根据节索引获取 Section 对象（1-based，-1=最后节）。

        Args:
            index: 节编号（从 1 开始），或负数（-1 为最后一节）

        Returns:
            Section COM 对象
        """
        secs = self._wb.document.Sections
        total = secs.Count
        if index < 0:
            index = total + index + 1
        if index < 1 or index > total:
            raise IndexError(f"节索引 {index} 超出范围（1~{total}）")
        return secs(index)

    def _page_setup(self, section: "CDispatch") -> "CDispatch":
        """获取指定节的 PageSetup 对象。"""
        return section.PageSetup

    # =========================================================================
    # 页面设置信息读取
    # =========================================================================

    def get_page_setup_info(self, index: int) -> dict:
        """
        读取指定节的完整页面设置信息。

        Args:
            index: 节编号（1-based）

        Returns:
            包含所有页面设置属性的字典
        """
        sec = self._resolve_section(index)
        ps = sec.PageSetup

        app = self._wb._word_app

        # 页边距（转厘米）
        top    = self._points_to_cm(ps.TopMargin)
        bottom = self._points_to_cm(ps.BottomMargin)
        left   = self._points_to_cm(ps.LeftMargin)
        right  = self._points_to_cm(ps.RightMargin)
        header = self._points_to_cm(ps.HeaderDistance)
        footer = self._points_to_cm(ps.FooterDistance)

        # 纸张
        width  = self._points_to_cm(ps.PageWidth)
        height = self._points_to_cm(ps.PageHeight)

        # 方向
        orient_raw = ps.Orientation
        orientation = _ORIENT_REVERSE.get(orient_raw, "portrait")

        # 分栏
        cols = ps.TextColumns
        col_count = cols.Count if cols.Count > 0 else 1
        try:
            col_width = self._points_to_cm(cols.Width)
        except Exception:
            col_width = 0.0
        try:
            col_spacing = self._points_to_cm(cols.Spacing)
        except Exception:
            col_spacing = 0.0
        try:
            equal_width = bool(cols.EvenlySpaced)
        except Exception:
            equal_width = True

        # 节属性
        sec_start_raw = sec.Properties("SectionType").Value
        try:
            sec_start = _SECTION_START_REVERSE.get(sec_start_raw, "continuous")
        except Exception:
            sec_start = "continuous"

        # 页眉页脚
        first_diff = bool(ps.DifferentFirstPageHeaderFooter)
        odd_even = bool(ps.OddAndEvenPagesHeaderFooter)

        # 垂直对齐
        v_align_raw = ps.VerticalAlignment
        v_align_map = {
            WD_ALIGN_PAGE_VERTICAL_TOP: "top",
            WD_ALIGN_PAGE_VERTICAL_CENTER: "center",
            WD_ALIGN_PAGE_VERTICAL_JUSTIFY: "justify",
            WD_ALIGN_PAGE_VERTICAL_BOTTOM: "bottom",
        }
        v_align = v_align_map.get(v_align_raw, "top")

        # 起始页码
        start_page = ps.StartingPageNumber

        return {
            "section_index": index,
            "margins": {
                "top":    round(top, 2),
                "bottom": round(bottom, 2),
                "left":   round(left, 2),
                "right":  round(right, 2),
                "header": round(header, 2),
                "footer": round(footer, 2),
            },
            "paper": {
                "width":  round(width, 2),
                "height": round(height, 2),
                "orientation": orientation,
            },
            "columns": {
                "count":       col_count,
                "width":       round(col_width, 2),
                "spacing":     round(col_spacing, 2),
                "equal_width": equal_width,
            },
            "section_start": sec_start,
            "first_page_different": first_diff,
            "odd_and_even_pages":   odd_even,
            "vertical_alignment":   v_align,
            "starting_page_number": start_page,
        }

    def get_page_margins(self, index: int) -> dict:
        """读取页边距（厘米）。"""
        sec = self._resolve_section(index)
        ps = sec.PageSetup
        return {
            "top":    round(self._points_to_cm(ps.TopMargin), 2),
            "bottom": round(self._points_to_cm(ps.BottomMargin), 2),
            "left":   round(self._points_to_cm(ps.LeftMargin), 2),
            "right":  round(self._points_to_cm(ps.RightMargin), 2),
        }

    def get_paper_size(self, index: int) -> dict:
        """读取纸张大小（厘米）。"""
        sec = self._resolve_section(index)
        ps = sec.PageSetup
        return {
            "width":  round(self._points_to_cm(ps.PageWidth), 2),
            "height": round(self._points_to_cm(ps.PageHeight), 2),
        }

    def get_orientation(self, index: int) -> str:
        """读取纸张方向：portrait 或 landscape。"""
        sec = self._resolve_section(index)
        ps = sec.PageSetup
        orient = ps.Orientation
        return _ORIENT_REVERSE.get(orient, "portrait")

    def get_column_count(self, index: int) -> int:
        """读取分栏数。"""
        sec = self._resolve_section(index)
        cols = sec.PageSetup.TextColumns
        return max(1, cols.Count if cols.Count > 0 else 1)

    def get_column_info(self, index: int) -> dict:
        """读取分栏详细信息。"""
        sec = self._resolve_section(index)
        cols = sec.PageSetup.TextColumns
        col_count = max(1, cols.Count if cols.Count > 0 else 1)
        try:
            width = self._points_to_cm(cols.Width)
        except Exception:
            width = 0.0
        try:
            spacing = self._points_to_cm(cols.Spacing)
        except Exception:
            spacing = 0.0
        try:
            equal = bool(cols.EvenlySpaced)
        except Exception:
            equal = True
        return {
            "count":       col_count,
            "width":       round(width, 2),
            "spacing":     round(spacing, 2),
            "equal_width": equal,
        }

    def get_section_start_type(self, index: int) -> str:
        """读取节起始类型。"""
        sec = self._resolve_section(index)
        try:
            st = sec.Properties("SectionType").Value
            return _SECTION_START_REVERSE.get(st, "continuous")
        except Exception:
            return "continuous"

    # =========================================================================
    # 页边距设置
    # =========================================================================

    def set_page_margins(
        self,
        index: int,
        top: Optional[float] = None,
        bottom: Optional[float] = None,
        left: Optional[float] = None,
        right: Optional[float] = None,
    ):
        """
        设置页边距（厘米）。

        Args:
            index: 节编号（1-based）
            top/bottom/left/right: 边距值（厘米），None 表示保持不变
        """
        sec = self._resolve_section(index)
        ps = sec.PageSetup

        if top is not None:
            ps.TopMargin = self._cm_to_points(top)
        if bottom is not None:
            ps.BottomMargin = self._cm_to_points(bottom)
        if left is not None:
            ps.LeftMargin = self._cm_to_points(left)
        if right is not None:
            ps.RightMargin = self._cm_to_points(right)

    def set_page_margins_by_inch(
        self,
        index: int,
        top: Optional[float] = None,
        bottom: Optional[float] = None,
        left: Optional[float] = None,
        right: Optional[float] = None,
    ):
        """
        设置页边距（英寸）。

        Args:
            index: 节编号（1-based）
            top/bottom/left/right: 边距值（英寸），None 表示保持不变
        """
        sec = self._resolve_section(index)
        ps = sec.PageSetup

        if top is not None:
            ps.TopMargin = self._inch_to_points(top)
        if bottom is not None:
            ps.BottomMargin = self._inch_to_points(bottom)
        if left is not None:
            ps.LeftMargin = self._inch_to_points(left)
        if right is not None:
            ps.RightMargin = self._inch_to_points(right)

    def set_page_margins_preset(self, index: int, preset: str):
        """
        使用预设方案设置页边距。

        Args:
            index: 节编号（1-based）
            preset: 预设名称（normal/narrow/wide/mirrored）
        """
        values = MARGIN_PRESETS_CM.get(preset.lower())
        if values is None:
            raise ValueError(f"未知的页边距预设: {preset}，可用：{list(MARGIN_PRESETS_CM.keys())}")
        self.set_page_margins(index, **values)

    # =========================================================================
    # 纸张设置
    # =========================================================================

    def set_paper_size(self, index: int, width: float, height: float):
        """
        设置纸张大小（厘米）。

        Args:
            index: 节编号（1-based）
            width: 纸张宽度（厘米）
            height: 纸张高度（厘米）
        """
        sec = self._resolve_section(index)
        ps = sec.PageSetup
        ps.PageWidth = self._cm_to_points(width)
        ps.PageHeight = self._cm_to_points(height)

    def set_paper_size_preset(self, index: int, preset: str):
        """
        使用预设类型设置纸张大小。

        Args:
            index: 节编号（1-based）
            preset: 纸张类型（A0~A6, B4, B5, Letter, Legal, Tabloid, Executive）
        """
        key = preset.upper()
        sizes = PAPER_PRESETS_CM.get(key)
        if sizes is None:
            raise ValueError(
                f"未知的纸张预设: {preset}，可用：{list(PAPER_PRESETS_CM.keys())}"
            )
        self.set_paper_size(index, *sizes)

    def set_orientation(self, index: int, orientation: str):
        """
        设置纸张方向。

        Args:
            index: 节编号（1-based）
            orientation: "portrait"（纵向）或 "landscape"（横向）
        """
        sec = self._resolve_section(index)
        ps = sec.PageSetup
        orient = _ORIENT_MAP.get(orientation.lower())
        if orient is None:
            raise ValueError(
                f"未知的纸张方向: {orientation}，可用：portrait, landscape"
            )
        ps.Orientation = orient

    # =========================================================================
    # 分栏操作
    # =========================================================================

    def set_columns(self, index: int, count: int, equal_width: bool = True):
        """
        设置分栏数。

        Args:
            index: 节编号（1-based）
            count: 栏数
            equal_width: 是否等宽栏
        """
        sec = self._resolve_section(index)
        ps = sec.PageSetup
        cols = ps.TextColumns
        cols.SetCount(count)
        if equal_width:
            cols.DistributeEvenly()

    def set_columns_with_gutter(
        self, index: int, count: int, spacing: float, equal_width: bool = True
    ):
        """
        设置分栏数并指定栏间距。

        Args:
            index: 节编号（1-based）
            count: 栏数
            spacing: 栏间距（厘米）
            equal_width: 是否等宽栏
        """
        sec = self._resolve_section(index)
        ps = sec.PageSetup
        cols = ps.TextColumns
        cols.SetCount(count)
        cols.Spacing = self._cm_to_points(spacing)
        if equal_width:
            cols.DistributeEvenly()

    def set_columns_equal_width(self, index: int):
        """将分栏设为等宽栏。"""
        sec = self._resolve_section(index)
        cols = sec.PageSetup.TextColumns
        cols.DistributeEvenly()

    def set_column_width(self, index: int, column: int, width: float):
        """
        设置指定栏的宽度（厘米）。

        Args:
            index: 节编号（1-based）
            column: 栏编号（从 1 开始）
            width: 栏宽（厘米）
        """
        sec = self._resolve_section(index)
        cols = sec.PageSetup.TextColumns
        n = cols.Count
        if column < 1 or column > n:
            raise IndexError(f"栏编号 {column} 超出范围（1~{n}）")
        col = cols(column)
        col.Width = self._cm_to_points(width)

    def apply_two_column_layout(self, index: int, with_line: bool = True):
        """
        应用两栏布局。

        Args:
            index: 节编号（1-based）
            with_line: 是否添加栏间分隔线
        """
        sec = self._resolve_section(index)
        ps = sec.PageSetup
        cols = ps.TextColumns
        cols.SetCount(2)
        cols.DistributeEvenly()
        if with_line:
            sec.PageSetup.LineBetween = True
        else:
            sec.PageSetup.LineBetween = False

    # =========================================================================
    # 页面边框与背景
    # =========================================================================

    def set_page_border(
        self,
        index: int,
        side: str = "all",
        line_style: int = 1,
        line_width: int = 6,
        color: Union[str, int] = 0x000000,
    ):
        """
        设置页面边框。

        Args:
            index: 节编号（1-based）
            side: "all" | "top" | "bottom" | "left" | "right"
            line_style: 线条样式（0=无, 1=单线, 2=双线...）
            line_width: 线条宽度
            color: 颜色（整数或颜色名）
        """
        if isinstance(color, str):
            color = self._resolve_color(color)
        sec = self._resolve_section(index)
        bds = sec.PageSetup.Borders

        side_map = {
            "all":     WD_BORDER_ALL,
            "top":     WD_BORDER_TOP,
            "bottom":  WD_BORDER_BOTTOM,
            "left":    WD_BORDER_LEFT,
            "right":   WD_BORDER_RIGHT,
        }
        border_side = side_map.get(side.lower(), WD_BORDER_ALL)

        # wdBorderAll = 9，表示同时设置四条边
        if border_side == WD_BORDER_ALL:
            for border in (bds.Top, bds.Left, bds.Right, bds.Bottom):
                border.LineStyle = line_style
                border.LineWidth = line_width
                border.Color = color
                border.Space = 6.0
        else:
            side_cap = side.lower().capitalize()
            border_attr = f"wdBorder{'{'}{side_cap}{'}'}"
            border = getattr(bds, border_attr)
            border.LineStyle = line_style
            border.LineWidth = line_width
            border.Color = color
            border.Space = 6.0

    def clear_page_border(self, index: int):
        """清除页面边框。"""
        sec = self._resolve_section(index)
        sec.PageSetup.Borders.Enable = 0

    def set_page_shading(
        self, index: int, fill_color: Union[str, int] = 0xCCE8FF
    ):
        """
        设置整页背景填充色。

        Args:
            index: 节编号（1-based）
            fill_color: 填充背景色（整数或颜色名）
        """
        if isinstance(fill_color, str):
            fill_color = self._resolve_color(fill_color)
        sec = self._resolve_section(index)
        shd = sec.PageSetup.Shading
        shd.Texture = 0
        shd.BackgroundPatternColor = fill_color

    def clear_page_shading(self, index: int):
        """清除页面背景色。"""
        sec = self._resolve_section(index)
        shd = sec.PageSetup.Shading
        shd.Texture = 0
        shd.BackgroundPatternColor = 0xFFFFFFFF

    # =========================================================================
    # 页面级格式
    # =========================================================================

    def set_vertical_alignment(self, index: int, align: str):
        """
        设置页面内容的垂直对齐方式。

        Args:
            index: 节编号（1-based）
            align: "top" | "center" | "justify" | "bottom"
        """
        sec = self._resolve_section(index)
        ps = sec.PageSetup
        v = _VERTICAL_ALIGN_MAP.get(align.lower())
        if v is None:
            raise ValueError(
                f"未知的垂直对齐方式: {align}，可用：top, center, justify, bottom"
            )
        ps.VerticalAlignment = v

    def set_starting_page_number(self, index: int, number: int):
        """设置起始页码。"""
        sec = self._resolve_section(index)
        sec.PageSetup.StartingPageNumber = max(1, int(number))

    def set_header_footer_distance(
        self,
        index: int,
        header: Optional[float] = None,
        footer: Optional[float] = None,
    ):
        """设置页眉/页脚与页面边缘的距离（厘米）。"""
        sec = self._resolve_section(index)
        ps = sec.PageSetup
        if header is not None:
            ps.HeaderDistance = self._cm_to_points(header)
        if footer is not None:
            ps.FooterDistance = self._cm_to_points(footer)

    # =========================================================================
    # 辅助方法
    # =========================================================================

    def _resolve_color(self, color: Union[str, int]) -> int:
        """将颜色名或整数转为 Word 颜色值。"""
        color_map = {
            "black":  0x000000,
            "white":  0xFFFFFF,
            "red":    0xFF0000,
            "green":  0x00FF00,
            "blue":   0x0000FF,
            "yellow": 0xFFFF00,
            "cyan":   0x00FFFF,
            "magenta":0xFF00FF,
            "gray":   0x808080,
            "darkBlue": 0x000080,
            "lightBlue": 0xCCE8FF,
            "orange": 0xFFA500,
            "purple": 0x800080,
        }
        if isinstance(color, str):
            if color.startswith("#"):
                return int(color[1:], 16)
            return color_map.get(color.lower(), 0x000000)
        return int(color)

    def apply_page_setup_to_all(self, index: int) -> int:
        """
        将当前节的页面设置（页边距、纸张、方向、分栏）复制到所有节。

        Returns:
            应用的节数量
        """
        sec = self._resolve_section(index)
        ps_src = sec.PageSetup
        secs = self._wb.document.Sections
        total = secs.Count
        for i in range(1, total + 1):
            if i == index:
                continue
            ps_dst = secs(i).PageSetup
            try:
                ps_dst.TopMargin    = ps_src.TopMargin
                ps_dst.BottomMargin = ps_src.BottomMargin
                ps_dst.LeftMargin   = ps_src.LeftMargin
                ps_dst.RightMargin  = ps_src.RightMargin
                ps_dst.PageWidth    = ps_src.PageWidth
                ps_dst.PageHeight   = ps_src.PageHeight
                ps_dst.Orientation  = ps_src.Orientation
                ps_dst.VerticalAlignment = ps_src.VerticalAlignment
            except Exception:
                pass
        return total

    def copy_page_setup(self, from_index: int, to_index: int) -> bool:
        """
        将源节的页面设置复制到目标节。

        Returns:
            是否成功
        """
        try:
            src = self._resolve_section(from_index)
            dst = self._resolve_section(to_index)
            ps_src = src.PageSetup
            ps_dst = dst.PageSetup

            ps_dst.TopMargin           = ps_src.TopMargin
            ps_dst.BottomMargin        = ps_src.BottomMargin
            ps_dst.LeftMargin          = ps_src.LeftMargin
            ps_dst.RightMargin         = ps_src.RightMargin
            ps_dst.PageWidth           = ps_src.PageWidth
            ps_dst.PageHeight          = ps_src.PageHeight
            ps_dst.Orientation         = ps_src.Orientation
            ps_dst.VerticalAlignment   = ps_src.VerticalAlignment
            ps_dst.HeaderDistance       = ps_src.HeaderDistance
            ps_dst.FooterDistance      = ps_src.FooterDistance
            ps_dst.StartingPageNumber  = ps_src.StartingPageNumber
            ps_dst.DifferentFirstPageHeaderFooter = ps_src.DifferentFirstPageHeaderFooter
            ps_dst.OddAndEvenPagesHeaderFooter    = ps_src.OddAndEvenPagesHeaderFooter
            ps_dst.LineBetween          = ps_src.LineBetween

            # 复制分栏
            try:
                cols_src = ps_src.TextColumns
                cols_dst = ps_dst.TextColumns
                cols_dst.SetCount(cols_src.Count)
                cols_dst.Spacing = cols_src.Spacing
                cols_dst.Width   = cols_src.Width
                cols_dst.EvenlySpaced = cols_src.EvenlySpaced
            except Exception:
                pass

            return True
        except Exception:
            return False

    def reset_page_setup(self, index: int):
        """
        重置指定节的页面设置为 Word 默认值。

        实际上通过 Normal.dotm 模板默认值覆盖实现。
        """
        sec = self._resolve_section(index)
        ps = sec.PageSetup
        app = self._wb._word_app

        # Word 默认值：A4 纵向
        ps.PageWidth  = self._cm_to_points(21.0)
        ps.PageHeight = self._cm_to_points(29.7)
        ps.Orientation = WD_ORIENT_PORTRAIT
        ps.TopMargin    = self._cm_to_points(2.54)
        ps.BottomMargin = self._cm_to_points(2.54)
        ps.LeftMargin   = self._cm_to_points(3.17)
        ps.RightMargin  = self._cm_to_points(3.17)
        ps.VerticalAlignment = WD_ALIGN_PAGE_VERTICAL_TOP
        ps.HeaderDistance = self._cm_to_points(1.5)
        ps.FooterDistance = self._cm_to_points(1.5)
        ps.StartingPageNumber = 1
        ps.DifferentFirstPageHeaderFooter = False
        ps.OddAndEvenPagesHeaderFooter    = False
        ps.LineBetween = False

        # 重置分栏为单栏
        ps.TextColumns.SetCount(1)
