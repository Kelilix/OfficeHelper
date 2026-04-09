# -*- coding: utf-8 -*-
"""
Section 节操作模块

提供基于 Document.Sections / Section 的节级操作能力：
- 节的新增、删除、属性读取
- 分节符插入
- 节起始类型（连续/下一页/偶数页/奇数页）
- 页眉页脚操作（主/首页/奇偶页）
- 奇偶页不同、首页不同设置
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Optional, Union

if TYPE_CHECKING:
    from win32com.client import CDispatch

from scripts.word_page_operator_base import (
    WD_SECTION_START_CONTINUOUS,
    WD_SECTION_START_NEW_PAGE,
    WD_SECTION_START_EVEN_PAGE,
    WD_SECTION_START_ODD_PAGE,
    WD_HEADER_PRIMARY,
    WD_HEADER_FIRST,
    WD_HEADER_EVEN_PAGES,
    _SECTION_START_MAP,
    _SECTION_START_REVERSE,
)


class SectionOperator:
    """
    节操作封装。

    封装了 Sections 集合和 Section 对象的各种属性和方法，提供：
    - 节的基础访问（数量、索引、遍历）
    - 分节符插入
    - 节属性设置（起始类型、首页不同、奇偶页不同）
    - 页眉页脚操作
    """

    def __init__(self, word_base):
        """
        Args:
            word_base: WordBase 实例
        """
        self._wb = word_base

    # =========================================================================
    # 节基础访问
    # =========================================================================

    def _sections(self) -> "CDispatch":
        """返回文档的 Sections 集合。"""
        return self._wb.document.Sections

    def count(self) -> int:
        """返回文档中的节总数。"""
        return self._sections().Count

    def get(self, index: int) -> "CDispatch":
        """
        按索引获取节（1-based，-1=最后节）。

        Args:
            index: 节编号（从 1 开始），或负数（-1 为最后一节）

        Returns:
            Section COM 对象
        """
        secs = self._sections()
        total = secs.Count
        if index < 0:
            index = total + index + 1
        if index < 1 or index > total:
            raise IndexError(f"节索引 {index} 超出范围（1~{total}）")
        return secs(index)

    def all(self) -> list:
        """返回所有节对象的列表。"""
        secs = self._sections()
        return [secs(i) for i in range(1, secs.Count + 1)]

    def get_index_of_section(self, section: "CDispatch") -> int:
        """
        获取指定节对象的索引（1-based）。

        Args:
            section: Section COM 对象

        Returns:
            节编号（从 1 开始）
        """
        secs = self._sections()
        total = secs.Count
        for i in range(1, total + 1):
            if secs(i).Range.Start == section.Range.Start:
                return i
        raise ValueError("无法在文档中定位该节")

    def delete_section_break(self, index: int):
        """
        删除指定节前面的分节符，或删除所有分节符（index=0）。

        注意：
        - 不能删除文档的第一个节（它前面没有分节符）
        - 删除后，该节的页面设置随之丢失，内容并入前一节

        Args:
            index: 节编号（1-based），或 0（删除所有分节符，从后往前删）
        """
        total = self.count()
        if index == 0:
            # 删除所有（从后往前避免索引偏移）
            for i in range(total, 1, -1):
                self._delete_single(i)
            return
        if index < 1 or index > total:
            raise IndexError(f"节索引 {index} 超出范围（1~{total}）")
        if index == 1:
            raise ValueError("无法删除文档首节前的分节符（首节前不存在分节符）")
        self._delete_single(index)

    def _delete_single(self, index: int):
        """内部方法：删除指定节前面的分节符，不做边界检查。"""
        sec = self.get(index)
        start = sec.Range.Start - 1
        if start < 0:
            return
        # 向前扩展一个段落，确保覆盖分节符字符（chr(12) / chr(13)）
        rng = self._wb.document.Range(start, sec.Range.Start)
        rng.Expand(Unit=4)  # wdParagraph = 4
        # 用 Find 查找 chr(12)（分节符内部存储），替换为空即删除
        rng.Find.ClearFormatting()
        # 用 Find 查找 chr(12)（分节符内部存储），Execute 返回 True 表示找到并替换了
        rng.Find.Text = "\x0c"          # chr(12) = Word 内部的分节符字符
        rng.Find.MatchCase = True
        rng.Find.Forward = False
        rng.Find.Wrap = 0
        if rng.Find.Execute(Replace=2):  # wdReplaceOne，替换为空即删除
            return
        # 回退：直接删除 chr(13)（段落末尾的分节符标记）
        rng2 = self._wb.document.Range(start, start + 1)
        if rng2.Text and ord(rng2.Text) <= 31:
            rng2.Delete()

    def get_current_section_index(self) -> int:
        """
        获取当前 Selection（光标）所在节的索引。

        Returns:
            节编号（1-based）
        """
        sel = self._wb.selection
        try:
            sec = sel.Sections(1)
            return self.get_index_of_section(sec)
        except Exception:
            return 1

    # =========================================================================
    # 分节符操作
    # =========================================================================

    def insert_section_break(
        self, rng, break_type: str = "new_page"
    ) -> "CDispatch":
        """
        在指定 Range 处插入分节符。

        插入分节符后，光标自动移到新节开头。

        Args:
            rng: Word Range 对象（当前位置）
            break_type: 分节符类型
                - "continuous"：连续（无分页，用于杂志排版等）
                - "new_page"：下一页（从新页面开始）
                - "even_page"：偶数页（从下一个偶数页开始）
                - "odd_page"：奇数页（从下一个奇数页开始）

        Returns:
            新插入的 Section 对象
        """
        break_val = _SECTION_START_MAP.get(break_type.lower(), WD_SECTION_START_NEW_PAGE)

        # 在 rng 位置插入分节符（用 InsertBreak 插入 SectionBreak）
        # Word 中 wdSectionBreak = 4
        rng.InsertBreak(Type=4)  # wdSectionBreak

        # 设置当前节（刚分出来的节）的起始类型
        # 插入分节符后，光标会在新节中，Sections.Last 指向新节
        new_sec = self._sections().Last
        new_sec.Properties("SectionType").Value = break_val

        return new_sec

    def set_section_start_type(self, index: int, type_str: str):
        """
        设置指定节的起始类型。

        Args:
            index: 节编号（1-based）
            type_str: 起始类型
                - "continuous" / "new_page" / "even_page" / "odd_page"
        """
        sec = self.get(index)
        val = _SECTION_START_MAP.get(type_str.lower())
        if val is None:
            raise ValueError(
                f"未知的节起始类型: {type_str}，可用：continuous, new_page, even_page, odd_page"
            )
        sec.Properties("SectionType").Value = val

    def set_section_start_new_page(self, index: int):
        """快捷方式：将节设为从新页开始。"""
        self.set_section_start_type(index, "new_page")

    def set_section_start_continuous(self, index: int):
        """快捷方式：将节设为连续（无分页）。"""
        self.set_section_start_type(index, "continuous")

    def set_section_start_even_page(self, index: int):
        """快捷方式：将节设为从偶数页开始。"""
        self.set_section_start_type(index, "even_page")

    def set_section_start_odd_page(self, index: int):
        """快捷方式：将节设为从奇数页开始。"""
        self.set_section_start_type(index, "odd_page")

    def get_section_start_type(self, index: int) -> str:
        """
        读取节的起始类型。

        Args:
            index: 节编号（1-based）

        Returns:
            "continuous" | "new_page" | "even_page" | "odd_page"
        """
        sec = self.get(index)
        try:
            st = sec.Properties("SectionType").Value
            return _SECTION_START_REVERSE.get(st, "continuous")
        except Exception:
            return "continuous"

    # =========================================================================
    # 首页不同 / 奇偶页不同
    # =========================================================================

    def set_first_page_different(self, index: int, on: bool = True):
        """
        设置首页不同（首页使用不同的页眉页脚）。

        Args:
            index: 节编号（1-based）
            on: True=开启首页不同，False=关闭
        """
        sec = self.get(index)
        sec.PageSetup.DifferentFirstPageHeaderFooter = -1 if on else 0

    def set_odd_and_even_pages(self, index: int, on: bool = True):
        """
        设置奇偶页不同（奇偶页使用不同的页眉页脚）。

        Args:
            index: 节编号（1-based）
            on: True=开启奇偶页不同，False=关闭
        """
        sec = self.get(index)
        sec.PageSetup.OddAndEvenPagesHeaderFooter = -1 if on else 0

    def is_first_page_different(self, index: int) -> bool:
        """判断首页不同是否开启。"""
        sec = self.get(index)
        return bool(sec.PageSetup.DifferentFirstPageHeaderFooter)

    def is_odd_and_even_pages(self, index: int) -> bool:
        """判断奇偶页不同是否开启。"""
        sec = self.get(index)
        return bool(sec.PageSetup.OddAndEvenPagesHeaderFooter)

    # =========================================================================
    # 页眉页脚操作
    # =========================================================================

    def _resolve_header_footer_index(self, position: str) -> int:
        """
        将位置名称转为 wdHeaderFooterIndex 常量。

        Args:
            position: "primary" | "first" | "even_odd"

        Returns:
            wdHeaderFooterIndex 值
        """
        pos_map = {
            "primary":  WD_HEADER_PRIMARY,
            "first":    WD_HEADER_FIRST,
            "even_odd": WD_HEADER_EVEN_PAGES,
        }
        val = pos_map.get(position.lower())
        if val is None:
            raise ValueError(
                f"未知的页眉页脚位置: {position}，可用：primary, first, even_odd"
            )
        return val

    def set_header(
        self,
        index: int,
        position: str = "primary",
        text: str = "",
        alignment: str = "left",
    ):
        """
        设置页眉内容。

        Args:
            index: 节编号（1-based）
            position: "primary"（主页眉）| "first"（首页眉）| "even_odd"（偶数页眉）
            text: 页眉文本
            alignment: "left" | "center" | "right"
        """
        sec = self.get(index)
        hdr_idx = self._resolve_header_footer_index(position)
        hdr_range = sec.Headers(hdr_idx).Range
        hdr_range.Text = text

        # 设置对齐
        align_map = {"left": 0, "center": 1, "right": 2}
        hdr_range.ParagraphFormat.Alignment = align_map.get(alignment.lower(), 0)

    def get_header(
        self, index: int, position: str = "primary"
    ) -> str:
        """
        读取页眉内容。

        Args:
            index: 节编号（1-based）
            position: "primary" | "first" | "even_odd"

        Returns:
            页眉文本
        """
        sec = self.get(index)
        hdr_idx = self._resolve_header_footer_index(position)
        return sec.Headers(hdr_idx).Range.Text.rstrip("\r")

    def clear_header(self, index: int, position: str = "primary"):
        """
        清除页眉内容。

        Args:
            index: 节编号（1-based）
            position: "primary" | "first" | "even_odd"
        """
        self.set_header(index, position=position, text="")

    def set_footer(
        self,
        index: int,
        position: str = "primary",
        text: str = "",
        alignment: str = "left",
    ):
        """
        设置页脚内容。

        Args:
            index: 节编号（1-based）
            position: "primary"（主页脚）| "first"（首页脚）| "even_odd"（偶数页脚）
            text: 页脚文本
            alignment: "left" | "center" | "right"
        """
        sec = self.get(index)
        hdr_idx = self._resolve_header_footer_index(position)
        ftr_range = sec.Footers(hdr_idx).Range
        ftr_range.Text = text

        align_map = {"left": 0, "center": 1, "right": 2}
        ftr_range.ParagraphFormat.Alignment = align_map.get(alignment.lower(), 0)

    def get_footer(
        self, index: int, position: str = "primary"
    ) -> str:
        """
        读取页脚内容。

        Args:
            index: 节编号（1-based）
            position: "primary" | "first" | "even_odd"

        Returns:
            页脚文本
        """
        sec = self.get(index)
        hdr_idx = self._resolve_header_footer_index(position)
        return sec.Footers(hdr_idx).Range.Text.rstrip("\r")

    def clear_footer(self, index: int, position: str = "primary"):
        """
        清除页脚内容。

        Args:
            index: 节编号（1-based）
            position: "primary" | "first" | "even_odd"
        """
        self.set_footer(index, position=position, text="")

    def insert_page_number_in_header(
        self,
        index: int,
        position: str = "primary",
        alignment: str = "right",
    ):
        """
        在页眉中插入页码字段。

        Args:
            index: 节编号（1-based）
            position: "primary" | "first" | "even_odd"
            alignment: "left" | "center" | "right"
        """
        sec = self.get(index)
        hdr_idx = self._resolve_header_footer_index(position)
        hdr_range = sec.Headers(hdr_idx).Range

        # 插入 PAGE 域
        hdr_range.Collapse(Direction=1)  # wdCollapseEnd
        hdr_range.InsertAfter("\t")
        # 设置对齐
        align_map = {"left": 0, "center": 1, "right": 2}
        hdr_range.ParagraphFormat.Alignment = align_map.get(alignment.lower(), 2)

        # 将文本转为域
        hdr_range.Fields.Add(
            Range=hdr_range,
            Type=33,  # wdFieldPage
        )

    def insert_page_number_in_footer(
        self,
        index: int,
        position: str = "primary",
        alignment: str = "center",
    ):
        """
        在页脚中插入页码字段。

        Args:
            index: 节编号（1-based）
            position: "primary" | "first" | "even_odd"
            alignment: "left" | "center" | "right"
        """
        sec = self.get(index)
        ftr_idx = self._resolve_header_footer_index(position)
        ftr_range = sec.Footers(ftr_idx).Range

        # 在开头插入页码域
        ftr_range.Collapse(Direction=0)  # wdCollapseStart
        ftr_range.Fields.Add(
            Range=ftr_range,
            Type=33,  # wdFieldPage
        )
        # 追加页数
        ftr_range.InsertAfter(" / ")
        rng2 = sec.Footers(ftr_idx).Range.Duplicate
        rng2.Collapse(Direction=1)
        rng2.InsertAfter("")
        ftr_range2 = sec.Footers(ftr_idx).Range.Duplicate
        rng_end = ftr_range2.End - 1
        from_range = self._wb.document.Range(rng_end, rng_end)
        from_range.Fields.Add(
            Range=from_range,
            Type=85,  # wdFieldNumPages（总页数）
        )

        # 设置对齐
        align_map = {"left": 0, "center": 1, "right": 2}
        ftr_range.ParagraphFormat.Alignment = align_map.get(alignment.lower(), 1)

    def set_header_link(
        self, index: int, position: str = "primary", link_to_previous: bool = True
    ):
        """
        设置页眉是否链接到前一节。

        Args:
            index: 节编号（1-based）
            position: "primary" | "first" | "even_odd"
            link_to_previous: True=链接到前一节，False=断开链接
        """
        sec = self.get(index)
        hdr_idx = self._resolve_header_footer_index(position)
        hdr = sec.Headers(hdr_idx)
        hdr.LinkToPrevious = -1 if link_to_previous else 0

    def set_footer_link(
        self, index: int, position: str = "primary", link_to_previous: bool = True
    ):
        """
        设置页脚是否链接到前一节。

        Args:
            index: 节编号（1-based）
            position: "primary" | "first" | "even_odd"
            link_to_previous: True=链接到前一节，False=断开链接
        """
        sec = self.get(index)
        hdr_idx = self._resolve_header_footer_index(position)
        ftr = sec.Footers(hdr_idx)
        ftr.LinkToPrevious = -1 if link_to_previous else 0
