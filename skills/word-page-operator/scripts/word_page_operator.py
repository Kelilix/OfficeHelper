# -*- coding: utf-8 -*-
"""
Word 页面操作主模块

整合 PageSetupOperator 和 SectionOperator，提供统一的页面设置操作能力。

本模块是 word-text-operator 和 word-paragraph-operator 的互补模块：
- word-text-operator：文本内容、字符格式
- word-paragraph-operator：段落格式
- word-page-operator：页面布局（页边距、纸张、分栏、分节、页眉页脚）
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from win32com.client import CDispatch

from scripts.word_page_operator_base import PageSetupOperator
from scripts.word_section_operator import SectionOperator


class PageOperator:
    """
    页面操作整合类。

    通过组合 PageSetupOperator（页面设置）和 SectionOperator（节操作），
    提供完整的页面级操作能力。
    """

    def __init__(self, word_base):
        """
        Args:
            word_base: WordBase 实例
        """
        self._wb = word_base
        self.page = PageSetupOperator(word_base)
        self.section = SectionOperator(word_base)

    # ── 节基础访问 ─────────────────────────────────────────────────────────

    def count(self) -> int:
        """返回文档中的节总数。"""
        return self.section.count()

    def get(self, index: int) -> "CDispatch":
        """
        按索引获取节（1-based，-1=最后节）。

        Args:
            index: 节编号（从 1 开始），或负数（-1 为最后一节）

        Returns:
            Section COM 对象
        """
        return self.section.get(index)

    def all(self) -> list:
        """返回所有节对象的列表。"""
        return self.section.all()

    def get_index(self, section: "CDispatch") -> int:
        """返回指定节对象的索引（1-based）。"""
        return self.section.get_index_of_section(section)

    # ── 快捷方法（委托给 page / section 子模块） ───────────────────────────

    def get_page_setup_info(self, index: int) -> dict:
        """读取完整页面设置信息。"""
        return self.page.get_page_setup_info(index)

    def get_page_margins(self, index: int) -> dict:
        """读取页边距。"""
        return self.page.get_page_margins(index)

    def set_page_margins(self, index: int, **kwargs):
        """设置页边距（厘米）。"""
        self.page.set_page_margins(index, **kwargs)

    def set_page_margins_by_inch(self, index: int, **kwargs):
        """设置页边距（英寸）。"""
        self.page.set_page_margins_by_inch(index, **kwargs)

    def set_page_margins_preset(self, index: int, preset: str):
        """使用预设方案设置页边距。"""
        self.page.set_page_margins_preset(index, preset)

    def get_paper_size(self, index: int) -> dict:
        """读取纸张大小。"""
        return self.page.get_paper_size(index)

    def set_paper_size(self, index: int, width: float, height: float):
        """设置纸张大小（厘米）。"""
        self.page.set_paper_size(index, width, height)

    def set_paper_size_preset(self, index: int, preset: str):
        """使用预设纸张类型。"""
        self.page.set_paper_size_preset(index, preset)

    def get_orientation(self, index: int) -> str:
        """读取纸张方向。"""
        return self.page.get_orientation(index)

    def set_orientation(self, index: int, orientation: str):
        """设置纸张方向。"""
        self.page.set_orientation(index, orientation)

    def get_column_count(self, index: int) -> int:
        """读取分栏数。"""
        return self.page.get_column_count(index)

    def get_column_info(self, index: int) -> dict:
        """读取分栏详情。"""
        return self.page.get_column_info(index)

    def set_columns(self, index: int, count: int, **kwargs):
        """设置分栏数。"""
        self.page.set_columns(index, count, **kwargs)

    def set_columns_with_gutter(self, index: int, count: int, spacing: float, **kwargs):
        """设置分栏数并指定栏间距。"""
        self.page.set_columns_with_gutter(index, count, spacing, **kwargs)

    def set_columns_equal_width(self, index: int):
        """设为等宽栏。"""
        self.page.set_columns_equal_width(index)

    def apply_two_column_layout(self, index: int, with_line: bool = True):
        """应用两栏布局。"""
        self.page.apply_two_column_layout(index, with_line)

    def set_vertical_alignment(self, index: int, align: str):
        """设置页面垂直对齐。"""
        self.page.set_vertical_alignment(index, align)

    def set_page_border(self, index: int, **kwargs):
        """设置页面边框。"""
        self.page.set_page_border(index, **kwargs)

    def clear_page_border(self, index: int):
        """清除页面边框。"""
        self.page.clear_page_border(index)

    def set_page_shading(self, index: int, **kwargs):
        """设置页面背景色。"""
        self.page.set_page_shading(index, **kwargs)

    def clear_page_shading(self, index: int):
        """清除页面背景色。"""
        self.page.clear_page_shading(index)

    def apply_page_setup_to_all(self, index: int) -> int:
        """将当前节设置应用到所有节。"""
        return self.page.apply_page_setup_to_all(index)

    def copy_page_setup(self, from_index: int, to_index: int) -> bool:
        """复制页面设置。"""
        return self.page.copy_page_setup(from_index, to_index)

    def reset_page_setup(self, index: int):
        """重置页面设置为默认值。"""
        self.page.reset_page_setup(index)

    # ── 分节操作 ──────────────────────────────────────────────────────────

    def insert_section_break(self, rng, break_type: str = "new_page"):
        """插入分节符。"""
        return self.section.insert_section_break(rng, break_type)

    def delete_section_break(self, index: int):
        """删除指定节前面的分节符。"""
        return self.section.delete_section_break(index)

    def set_section_start_type(self, index: int, type_str: str):
        """设置节起始类型。"""
        self.section.set_section_start_type(index, type_str)

    def set_section_start_new_page(self, index: int):
        """节从新页开始。"""
        self.section.set_section_start_new_page(index)

    def set_section_start_continuous(self, index: int):
        """节设为连续。"""
        self.section.set_section_start_continuous(index)

    def set_section_start_even_page(self, index: int):
        """节从偶数页开始。"""
        self.section.set_section_start_even_page(index)

    def set_section_start_odd_page(self, index: int):
        """节从奇数页开始。"""
        self.section.set_section_start_odd_page(index)

    def set_first_page_different(self, index: int, on: bool = True):
        """设置首页不同。"""
        self.section.set_first_page_different(index, on)

    def set_odd_and_even_pages(self, index: int, on: bool = True):
        """设置奇偶页不同。"""
        self.section.set_odd_and_even_pages(index, on)

    # ── 页眉页脚 ─────────────────────────────────────────────────────────

    def set_header(self, index: int, **kwargs):
        """设置页眉。"""
        self.section.set_header(index, **kwargs)

    def get_header(self, index: int, position: str = "primary") -> str:
        """读取页眉。"""
        return self.section.get_header(index, position)

    def clear_header(self, index: int, position: str = "primary"):
        """清除页眉。"""
        self.section.clear_header(index, position)

    def set_footer(self, index: int, **kwargs):
        """设置页脚。"""
        self.section.set_footer(index, **kwargs)

    def get_footer(self, index: int, position: str = "primary") -> str:
        """读取页脚。"""
        return self.section.get_footer(index, position)

    def clear_footer(self, index: int, position: str = "primary"):
        """清除页脚。"""
        self.section.clear_footer(index, position)

    def insert_page_number_in_header(self, index: int, **kwargs):
        """在页眉中插入页码。"""
        self.section.insert_page_number_in_header(index, **kwargs)

    def insert_page_number_in_footer(self, index: int, **kwargs):
        """在页脚中插入页码。"""
        self.section.insert_page_number_in_footer(index, **kwargs)

    # ── 分页控制 ──────────────────────────────────────────────────────────

    def get_page_count(self) -> int:
        """获取文档总页数。"""
        doc = self._wb.document
        return doc.ComputeStatistics(2)  # wdStatisticPages = 2

    def get_section_count(self) -> int:
        """获取文档总节数。"""
        return self.count()

    def get_page_of_range(self, rng) -> int:
        """
        获取指定 Range 所在页的页码（从 1 开始）。

        Args:
            rng: Word Range 对象

        Returns:
            页码
        """
        try:
            return rng.Information(3)  # wdActiveEndPageNumber = 3
        except Exception:
            return 1
