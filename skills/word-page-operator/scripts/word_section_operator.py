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

        核心逻辑：
        - Word 中，段落文本末尾存储顺序为：内容 + chr(12)(分节符) + chr(13)(段落标记)
        - 分节符 chr(12) 紧邻段落末尾 chr(13) 之前
        - 操作：展开到整段落末尾 → 向左收缩 1（Word 跳过 chr(13)）→ 落点恰好是 chr(12)
        - 删除 chr(12) 即可只移除分节符，保留 chr(13) 和段落内容不变
        - 从后往前遍历段落，避免删除后索引偏移导致漏删或越界

        注意：
        - 第一个节前面没有分节符，不可删除
        - 删除分节符不影响节的内容，只是把两个节的页面格式合并到前一节

        Args:
            index: 节编号（1-based），或 0（删除所有分节符）
        """
        total = self.count()
        if index == 0:
            self._delete_all_section_breaks()
            return
        if index < 1 or index > total:
            raise IndexError(f"节索引 {index} 超出范围（1~{total}）")
        if index == 1:
            raise ValueError("无法删除文档首节前的分节符（首节前不存在分节符）")
        self._delete_section_break_before(index)

    def _delete_section_break_before(self, section_index: int):
        """
        删除第 section_index 节前面的分节符（section_index >= 2）。

        定位方式：section_index 节的 Range.Start 紧邻在分节符 chr(12) 之后。
        因此向前取 target_start - 1 的位置，就是分节符 chr(12) 所在的位置。
        """
        if section_index < 2:
            return

        secs = self._sections()
        # section_index 节范围起点 = 紧邻分节符之后的位置
        target_start = secs(section_index).Range.Start
        if target_start < 2:
            return

        doc = self._wb.document
        # target_start - 1 就是分节符 chr(12) 所在的位置
        chr_12_pos = target_start - 1
        rng = doc.Range(chr_12_pos, chr_12_pos + 1)
        if rng.Text and ord(rng.Text) == 12:
            rng.Delete()

    def _delete_all_section_breaks(self):
        """
        删除文档中所有的分节符。

        遍历所有节（从后往前），对每个节（index >= 2）：
        - 该节的 Range.Start 紧跟在分节符 chr(12) 之后
        - 向左取 1 个字符，若为 chr(12) 则删除
        从后往前遍历可避免索引偏移导致的漏删或越界。
        """
        doc = self._wb.document
        secs = self._sections()
        total = secs.Count

        for i in range(total, 1, -1):
            target_start = secs(i).Range.Start
            if target_start < 2:
                continue
            chr_12_pos = target_start - 1
            rng = doc.Range(chr_12_pos, chr_12_pos + 1)
            if rng.Text and ord(rng.Text) == 12:
                rng.Delete()

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
        rng2 = sec.Footers(ftr_idx).Range.Duplicate
        rng2.Collapse(Direction=1)
        rng2.InsertAfter(" / ")
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
