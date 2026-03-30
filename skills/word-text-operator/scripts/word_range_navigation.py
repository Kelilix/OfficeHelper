# -*- coding: utf-8 -*-
"""
Word Range 导航模块

提供 Range 对象的各种导航能力：
- 移动指针（Start/End 属性）
- Move / MoveStart / MoveEnd / MoveWhile 族
- Expand 扩展范围
- URange / UBound 定位
- CompareLocation / InRange 范围比较
- GoTo 系列（定位到书签、域、表格等）

优先使用 Range，Selection 仅在用户交互场景下使用。
"""

from __future__ import annotations
from typing import TYPE_CHECKING, Optional, Union

if TYPE_CHECKING:
    from win32com.client import CDispatch


class RangeNavigator:
    """Range 导航操作封装。"""

    # wdGoToDirection 常量
    WD_GO_TO_PREVIOUS = -1
    WD_GO_TO_NEXT = 1
    WD_GO_TO_FIRST = 1
    WD_GO_TO_LAST = 2

    # wdGoToItem 常量
    WD_GO_TO_PAGE = 1
    WD_GO_TO_LINE = 5
    WD_GO_TO_SECTION = 8
    WD_GO_TO_BOOKMARK = -1
    WD_GO_TO_TABLE = 12
    WD_GO_TO_COMMENT = 6
    WD_GO_TO_ENDNOTE = 7
    WD_GO_TO_FOOTNOTE = 5
    WD_GO_TO_FIELD = 9
    WD_GO_TO_PROOFMARK = 10

    def __init__(self, word_base):
        """
        Args:
            word_base: WordBase 实例
        """
        self._wb = word_base

    # ========================================================================
    # 基础定位
    # ========================================================================

    def get_range(self, start: int, end: int) -> "CDispatch":
        """根据字符偏移创建 Range。"""
        return self._wb.get_range(start, end)

    def get_full_range(self) -> "CDispatch":
        """获取整个文档的 Range。"""
        return self._wb.document.Range(
            Start=0, End=self._wb.document.Characters.Count
        )

    def get_selection_range(self) -> "CDispatch":
        """将当前 Selection 转为 Range 对象。"""
        return self._wb.selection.Range

    def set_range(self, rng: "CDispatch", start: int, end: int):
        """设置 Range 的起止位置。"""
        rng.Start = start
        rng.End = end

    def clone_range(self, rng: "CDispatch") -> "CDispatch":
        """复制一个 Range（独立对象）。"""
        return rng.Duplicate

    # ========================================================================
    # Expand / UBounds
    # ========================================================================

    def expand_to_sentence(self, rng: "CDispatch") -> int:
        """将 Range 扩展到完整句子。返回扩展后的类型常量。"""
        return rng.Expand(Unit=3)  # wdSentence = 3

    def expand_to_paragraph(self, rng: "CDispatch") -> int:
        """将 Range 扩展到完整段落。返回扩展后的字符数。"""
        return rng.Expand(Unit=4)  # wdParagraph = 4

    def expand_to_line(self, rng: "CDispatch") -> int:
        """将 Range 扩展到整行。返回扩展后的字符数。"""
        return rng.Expand(Unit=5)  # wdLine = 5

    def expand_to_word(self, rng: "CDispatch") -> int:
        """将 Range 扩展到完整单词。返回扩展后的字符数。"""
        return rng.Expand(Unit=2)  # wdWord = 2

    def expand_to_document(self, rng: "CDispatch") -> int:
        """将 Range 扩展到整个文档。返回扩展后的字符数。"""
        return rng.Expand(Unit=0)  # wdCell = 0 即整个文档

    def select_range(self, rng: "CDispatch"):
        """选中（高亮）指定的 Range（切换到 Selection 模式）。"""
        rng.Select()

    # ========================================================================
    # Move 系列
    # ========================================================================

    def move(self, rng: "CDispatch", unit: int = 4, count: int = 1) -> int:
        """
        移动 Range 的起止位置（整体移动）。

        Args:
            rng: Range 对象
            unit: 移动单位（1=Character, 2=Word, 3=Sentence, 4=Paragraph, 5=Line, 6=Story）
            count: 移动次数（正数前移，负数后移）

        Returns:
            实际移动的次数（可能小于请求数）
        """
        return rng.Move(Unit=unit, Count=count)

    def move_start(self, rng: "CDispatch", unit: int = 1, count: int = 1) -> int:
        """将 Range 的起始位置前移。返回实际移动次数。"""
        return rng.MoveStart(Unit=unit, Count=count)

    def move_end(self, rng: "CDispatch", unit: int = 1, count: int = 1) -> int:
        """将 Range 的结束位置前移。返回实际移动次数。"""
        return rng.MoveEnd(Unit=unit, Count=count)

    def move_while(
        self, rng: "CDispatch", characters: str, count: int = 100
    ) -> int:
        """
        当遇到 characters 中的字符时，持续前移 Start 指针。

        Args:
            characters: 要排除的字符集（如 " \t\r\n"）
            count: 最大移动字符数

        Returns:
            实际移动的字符数
        """
        return rng.MoveWhile(Count=count, Cset=characters)

    def move_until(
        self, rng: "CDispatch", characters: str, count: int = 100
    ) -> int:
        """
        持续前移 Start 指针，直到遇到 characters 中的字符。

        Args:
            characters: 停止字符集（如 " \t\r\n"）
            count: 最大移动字符数

        Returns:
            实际移动的字符数
        """
        return rng.MoveUntil(Count=count, Cset=characters)

    def move_while_and(self, rng: "CDispatch", characters: str, count: int = 100):
        """
        同时前移 Start 和 End 指针，直到遇到非 characters 中的字符。
        与 MoveWhile 不同：MoveWhile 仅移动 Start，此方法同时收缩两端。
        """
        return rng.MoveWhile(Count=count, Cset=characters, Matchaval=1)

    def move_end_until(
        self, rng: "CDispatch", characters: str, count: int = 100
    ) -> int:
        """将 Range.End 向后移动，直到遇到 characters 中的字符。"""
        return rng.MoveEndUntil(Count=count, Cset=characters)

    def move_end_while(
        self, rng: "CDispatch", characters: str, count: int = 100
    ) -> int:
        """将 Range.End 向后移动，直到遇到非 characters 中的字符。"""
        return rng.MoveEndWhile(Count=count, Cset=characters)

    def move_start_unit(self, rng: "CDispatch", unit: int = 4, count: int = 1) -> int:
        """按单位前移 Start。wdUnit 同上。"""
        return rng.MoveStart(Unit=unit, Count=count)

    def move_end_unit(self, rng: "CDispatch", unit: int = 4, count: int = 1) -> int:
        """按单位前移 End。wdUnit 同上。"""
        return rng.MoveEnd(Unit=unit, Count=count)

    def collapse(self, rng: "CDispatch", direction: str = "start"):
        """
        折叠 Range 为空（Start == End），即光标位置。

        Args:
            direction: "start"（向前折叠）或 "end"（向后折叠）
        """
        rng.Collapse(Direction=0 if direction == "start" else 1)

    # ========================================================================
    # 范围比较
    # ========================================================================

    def in_range(self, rng: "CDispatch", container: "CDispatch") -> bool:
        """
        判断 rng 是否完全包含在 container 内部。

        注意：Word 的 InRange 是 rng 的视角——rng 是否在 container 内。
        """
        return rng.InRange(container)

    def compare_location(
        self, rng1: "CDispatch", rng2: "CDispatch"
    ) -> int:
        """
        比较两个 Range 的相对位置。

        Returns:
            -1: rng1 在 rng2 之前
             0: 重叠或相同
             1: rng1 在 rng2 之后
        """
        if rng1.End <= rng2.Start:
            return -1
        elif rng1.Start >= rng2.End:
            return 1
        else:
            return 0

    def is_equal(self, rng1: "CDispatch", rng2: "CDispatch") -> bool:
        """判断两个 Range 是否完全相同（起止位置一致）。"""
        return rng1.Start == rng2.Start and rng1.End == rng2.End

    def is_inside(self, rng: "CDispatch", container: "CDispatch") -> bool:
        """判断 rng 是否严格在 container 内部（不含边界）。"""
        return rng.Start > container.Start and rng.End < container.End

    # ========================================================================
    # GoTo 系列
    # ========================================================================

    def go_to_bookmark(self, name: str) -> Optional["CDispatch"]:
        """
        跳转到指定书签，返回其 Range。

        Args:
            name: 书签名

        Returns:
            书签的 Range，未找到返回 None
        """
        try:
            bm = self._wb.document.Bookmarks(name)
            return bm.Range
        except Exception:
            return None

    def go_to_comment(self, index: int) -> Optional["CDispatch"]:
        """
        跳转到第 N 个批注。

        Args:
            index: 批注编号（从 1 开始）

        Returns:
            批注的 Range
        """
        try:
            comment = self._wb.document.Comments(index)
            return comment.Reference
        except Exception:
            return None

    def go_to_page(self, page_num: int) -> "CDispatch":
        """
        跳转到指定页码开头。

        Args:
            page_num: 页码（从 1 开始）

        Returns:
            该页开头的 Range
        """
        rng = self._wb.document.GoTo(
            What=1,  # wdGoToPage
            Which=1,  # wdGoToFirst
            Count=page_num - 1,
        )
        return rng

    def go_to_line(self, line_num: int) -> "CDispatch":
        """跳转到指定行号。"""
        return self._wb.document.GoTo(
            What=5,  # wdGoToLine
            Which=1,  # wdGoToFirst
            Count=line_num,
        )

    def go_to_end(self) -> "CDispatch":
        """跳转到文档末尾。"""
        return self._wb.document.GoTo(
            What=5, Which=2  # wdGoToLast
        )

    def go_to_start(self) -> "CDispatch":
        """跳转到文档开头。"""
        return self._wb.document.Range(Start=0, End=0)

    # ========================================================================
    # 获取范围信息
    # ========================================================================

    def get_characters(self, rng: "CDispatch") -> "CDispatch":
        """获取 Range 内的所有字符集合（Characters 对象）。"""
        return rng.Characters

    def get_words(self, rng: "CDispatch") -> "CDispatch":
        """获取 Range 内的所有单词集合。"""
        return rng.Words

    def get_sentences(self, rng: "CDispatch") -> "CDispatch":
        """获取 Range 内的所有句子集合。"""
        return rng.Sentences

    def get_paragraphs(self, rng: "CDispatch") -> "CDispatch":
        """获取 Range 内的所有段落集合。"""
        return rng.Paragraphs

    def get_bookmarks_in_range(self, rng: "CDispatch") -> list:
        """获取 Range 内的所有书签名称列表。"""
        return [
            bm.Name
            for bm in self._wb.document.Bookmarks
            if rng.InRange(bm.Range)
        ]

    def get_story_type(self, rng: "CDispatch") -> int:
        """返回 Range 所在 story 的类型。wdStoryType 常量。"""
        return rng.StoryType

    def get_length(self, rng: "CDispatch") -> int:
        """返回 Range 的字符长度。"""
        return rng.End - rng.Start

    def get_text(self, rng: "CDispatch") -> str:
        """返回 Range 的文本内容。"""
        return rng.Text
