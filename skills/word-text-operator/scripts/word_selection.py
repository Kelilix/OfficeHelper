# -*- coding: utf-8 -*-
"""
Word Selection 操作模块

封装 Selection 对象的各种操作能力。
Selection 表示当前用户选中的区域（或光标点），与 Range 的区别在于：
- Selection 是"可见的"，总是有且只有一个
- 可以直接与用户交互（键盘/鼠标选中）
- 移动和扩展操作比 Range 更直观
- 适合需要即时反馈的场景

本模块在 Range 可用时优先使用 Range；Selection 作为 Range 的补充。
"""

from __future__ import annotations
from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from win32com.client import CDispatch


# wdUnits
WD_UNIT_CHARACTER = 1
WD_UNIT_WORD = 2
WD_UNIT_SENTENCE = 3
WD_UNIT_PARAGRAPH = 4
WD_UNIT_LINE = 5
WD_UNIT_STORY = 6
WD_UNIT_SCREEN = 7
WD_UNIT_SECTION = 8
WD_UNIT_COLUMN = 9
WD_UNIT_ROW = 10
WD_UNIT_WINDOW = 11

# wdMovementType
WD_MOVE = 0
WD_EXTEND = 1

# wdSelectionType
WD_NO_SELECTION = 0
WD_SELECTION = 1
WD_LINE_SELECTION = 2
WD_CHARACTER_SELECTION = 3
WD_WORD_SELECTION = 4
WD_SENTENCE_SELECTION = 5
WD_PARAGRAPH_SELECTION = 6
WD_BLOCK_SELECTION = 7
WD_INLINE_SHAPE_SELECTION = 8
WD_SHAPE_SELECTION = 9


class SelectionOperator:
    """Word Selection 操作封装。"""

    def __init__(self, word_base):
        self._wb = word_base

    # ========================================================================
    # 基础属性
    # ========================================================================

    @property
    def sel(self) -> "CDispatch":
        """返回当前 Selection 对象。"""
        return self._wb.selection

    @property
    def has_selection(self) -> bool:
        """检查是否有文本被选中（非折叠状态）。"""
        return self.sel.Exists

    @property
    def is_collapsed(self) -> bool:
        """检查 Selection 是否折叠（光标点，无选中内容）。"""
        return self.sel.Start == self.sel.End

    @property
    def selection_text(self) -> str:
        """返回当前 Selection 的文本。"""
        return self.sel.Text

    @property
    def selection_range(self) -> "CDispatch":
        """将 Selection 转为 Range 对象。"""
        return self.sel.Range

    @property
    def selection_start(self) -> int:
        """返回 Selection 起始位置。"""
        return self.sel.Start

    @property
    def selection_end(self) -> int:
        """返回 Selection 结束位置。"""
        return self.sel.End

    @property
    def selection_type(self) -> int:
        """返回当前 Selection 的类型常量。"""
        return self.sel.Type

    def get_selection_info(self) -> dict:
        """返回 Selection 的详细信息。"""
        return {
            "text": self.sel.Text,
            "start": self.sel.Start,
            "end": self.sel.End,
            "length": self.sel.End - self.sel.Start,
            "type": self.sel.Type,
            "type_name": self._get_type_name(self.sel.Type),
            "is_collapsed": self.sel.Start == self.sel.End,
        }

    # ========================================================================
    # 折叠 / 展开
    # ========================================================================

    def collapse_to_start(self):
        """将 Selection 折叠到起始点（向左收缩为光标）。"""
        self.sel.Collapse(Direction=0)  # wdCollapseStart

    def collapse_to_end(self):
        """将 Selection 折叠到结束点（向右收缩为光标）。"""
        self.sel.Collapse(Direction=1)  # wdCollapseEnd

    def expand_to_word(self) -> int:
        """将 Selection 扩展到包含当前单词。"""
        return self.sel.Expand(Unit=WD_UNIT_WORD)

    def expand_to_sentence(self) -> int:
        """将 Selection 扩展到包含当前句子。"""
        return self.sel.Expand(Unit=WD_UNIT_SENTENCE)

    def expand_to_paragraph(self) -> int:
        """将 Selection 扩展到包含当前段落。"""
        # sel.Expand(WD_UNIT_PARAGRAPH) 在 pywin32 下对部分 story 类型
        # 扩展不完整，改为 MoveEnd + MoveStart 强制扩到段首段尾。
        moved = self.sel.MoveEnd(Unit=WD_UNIT_PARAGRAPH, Count=1)
        self.sel.MoveStart(Unit=WD_UNIT_PARAGRAPH, Count=-1)
        return moved

    def expand_to_line(self) -> int:
        """将 Selection 扩展到整行。"""
        return self.sel.Expand(Unit=WD_UNIT_LINE)

    def expand_to_sentence_full(self) -> int:
        """将 Selection 扩展到完整句子（包含句号后的空格）。"""
        return self.sel.Expand(Unit=WD_UNIT_SENTENCE)

    # ========================================================================
    # 移动光标（Collapse + Move）
    # ========================================================================

    def move(self, unit: int = WD_UNIT_CHARACTER, count: int = 1) -> int:
        """
        移动光标（Selection）。

        Args:
            unit: 移动单位（WD_UNIT_* 常量）
            count: 移动次数（正=向前，负=向后）

        Returns:
            实际移动的次数
        """
        return self.sel.Move(Unit=unit, Count=count)

    def move_left(
        self, unit: int = WD_UNIT_CHARACTER, count: int = 1, extend: bool = False
    ):
        """
        向左移动光标。

        Args:
            unit: WD_UNIT_CHARACTER 或 WD_UNIT_WORD
            count: 移动字符/词数
            extend: True=选区扩展模式，False=折叠移动
        """
        self.sel.MoveLeft(Unit=unit, Count=count, Extend=WD_EXTEND if extend else WD_MOVE)

    def move_right(
        self, unit: int = WD_UNIT_CHARACTER, count: int = 1, extend: bool = False
    ):
        """向右移动光标（同 move_left）。"""
        self.sel.MoveRight(Unit=unit, Count=count, Extend=WD_EXTEND if extend else WD_MOVE)

    def move_up(
        self, unit: int = WD_UNIT_LINE, count: int = 1, extend: bool = False
    ):
        """向上移动光标。"""
        self.sel.MoveUp(Unit=unit, Count=count, Extend=WD_EXTEND if extend else WD_MOVE)

    def move_down(
        self, unit: int = WD_UNIT_LINE, count: int = 1, extend: bool = False
    ):
        """向下移动光标。"""
        self.sel.MoveDown(Unit=unit, Count=count, Extend=WD_EXTEND if extend else WD_MOVE)

    def move_to_line_start(self):
        """移动到当前行开头。"""
        self.sel.HomeKey(Unit=WD_UNIT_LINE)

    def move_to_line_end(self):
        """移动到当前行末尾。"""
        self.sel.EndKey(Unit=WD_UNIT_LINE)

    def move_to_document_start(self):
        """移动到文档开头。"""
        self.sel.HomeKey(Unit=WD_UNIT_STORY)

    def move_to_document_end(self):
        """移动到文档末尾。"""
        self.sel.EndKey(Unit=WD_UNIT_STORY)

    def move_to_paragraph_start(self):
        """移动到段落开头。"""
        self.sel.HomeKey(Unit=WD_UNIT_PARAGRAPH)

    def move_to_paragraph_end(self):
        """移动到段落末尾。"""
        self.sel.EndKey(Unit=WD_UNIT_PARAGRAPH)

    # ========================================================================
    # 扩展选区（Extend）
    # ========================================================================

    def extend_to_word(self):
        """扩展选区到当前单词末尾。"""
        self.sel.MoveRight(Unit=WD_UNIT_WORD, Count=1, Extend=WD_EXTEND)

    def extend_to_sentence(self):
        """扩展选区到当前句子。"""
        self.sel.Expand(Unit=WD_UNIT_SENTENCE)

    def extend_to_paragraph(self):
        """扩展选区到当前段落。"""
        self.sel.Expand(Unit=WD_UNIT_PARAGRAPH)

    def extend_to_line(self):
        """扩展选区到整行。"""
        self.sel.Expand(Unit=WD_UNIT_LINE)

    def extend_left(self, unit: int = WD_UNIT_CHARACTER, count: int = 1):
        """向左扩展选区。"""
        self.sel.MoveLeft(Unit=unit, Count=count, Extend=WD_EXTEND)

    def extend_right(self, unit: int = WD_UNIT_CHARACTER, count: int = 1):
        """向右扩展选区。"""
        self.sel.MoveRight(Unit=unit, Count=count, Extend=WD_EXTEND)

    def extend_up(self, count: int = 1):
        """向上扩展选区。"""
        self.sel.MoveUp(Unit=WD_UNIT_LINE, Count=count, Extend=WD_EXTEND)

    def extend_down(self, count: int = 1):
        """向下扩展选区。"""
        self.sel.MoveDown(Unit=WD_UNIT_LINE, Count=count, Extend=WD_EXTEND)

    def extend_to_match(self) -> bool:
        """
        扩展 Selection 直到遇到与开头相同字符为止。
        常用于选中配对引号、括号等。
        """
        return bool(self.sel.MoveWhile(Count=1, Cset=self.sel.Characters(1).Text, Extend=WD_EXTEND))

    # ========================================================================
    # 选中文本
    # ========================================================================

    def select_word(self):
        """选中光标所在单词。"""
        self.sel.Expand(Unit=WD_UNIT_WORD)

    def select_line(self):
        """选中当前行。"""
        self.sel.Expand(Unit=WD_UNIT_LINE)

    def select_paragraph(self):
        """选中当前段落。"""
        self.sel.Expand(Unit=WD_UNIT_PARAGRAPH)

    def select_sentence(self):
        """选中当前句子。"""
        self.sel.Expand(Unit=WD_UNIT_SENTENCE)

    def select_all(self):
        """选中整个文档。"""
        self.sel.WholeStory()

    def select_range(self, start: int, end: int):
        """选中指定字符范围的区域。"""
        rng = self._wb.get_range(start, end)
        rng.Select()

    # ========================================================================
    # 查找与替换（Selection 模式）
    # ========================================================================

    def find_and_select(
        self,
        text: str,
        whole_word: bool = False,
        match_case: bool = False,
        forward: bool = True,
        wrap: int = 0,
    ) -> bool:
        """
        在文档中查找并选中匹配项。

        Returns:
            是否找到
        """
        find = self.sel.Find
        find.ClearFormatting()
        find.Text = text
        find.WholeWord = whole_word
        find.MatchCase = match_case
        find.Forward = forward
        find.Wrap = wrap
        return bool(find.Execute())

    def find_next_and_select(
        self,
        text: str,
        whole_word: bool = False,
        match_case: bool = False,
    ) -> bool:
        """
        从当前位置向后查找并选中下一个匹配项。

        Returns:
            是否找到下一个
        """
        return self.find_and_select(text, whole_word, match_case, forward=True, wrap=0)

    def find_previous_and_select(
        self,
        text: str,
        whole_word: bool = False,
        match_case: bool = False,
    ) -> bool:
        """从当前位置向前查找并选中上一个匹配项。"""
        return self.find_and_select(text, whole_word, match_case, forward=False, wrap=0)

    def replace_selection(
        self,
        find_text: str,
        replace_text: str,
        whole_word: bool = False,
        match_case: bool = False,
        replace_all: bool = False,
    ) -> int:
        """
        在当前 Selection 中替换文本。

        Returns:
            替换次数
        """
        find = self.sel.Find
        find.ClearFormatting()
        find.Text = find_text
        find.WholeWord = whole_word
        find.MatchCase = match_case
        find.Forward = True
        find.Wrap = 0
        find.Replacement.ClearFormatting()
        find.Replacement.Text = replace_text

        replace_mode = 2 if replace_all else 1  # wdReplaceAll / wdReplaceOne
        return find.Execute(Replace=replace_mode)

    # ========================================================================
    # 格式化
    # ========================================================================

    def set_bold(self, bold: bool = True):
        """设置当前 Selection 的加粗。"""
        self.sel.Font.Bold = -1 if bold else 0

    def set_italic(self, italic: bool = True):
        """设置当前 Selection 的斜体。"""
        self.sel.Font.Italic = -1 if italic else 0

    def set_underline(self, underline: int = 1):
        """设置当前 Selection 的下划线。"""
        self.sel.Font.Underline = underline

    def set_font_name(self, name: str):
        """设置当前 Selection 的字体。"""
        self.sel.Font.Name = name

    def set_font_size(self, size: float):
        """设置当前 Selection 的字号。"""
        self.sel.Font.Size = size

    def set_font_color(self, color: int):
        """设置当前 Selection 的字体颜色。"""
        self.sel.Font.Color = color

    def set_highlight(self, highlight: int):
        """设置当前 Selection 的高亮。"""
        self.sel.Font.Highlight = highlight

    def set_alignment(self, align: int):
        """设置当前 Selection 段落的水平对齐。"""
        self.sel.ParagraphFormat.Alignment = align

    def clear_formatting(self):
        """清除当前 Selection 的格式（还原为默认样式）。"""
        self.sel.Font.Reset()
        self.sel.ParagraphFormat.Reset()

    # ========================================================================
    # 内容操作
    # ========================================================================

    def type_text(self, text: str):
        """
        在当前光标位置输入文本（替代已选中内容）。

        注意：如果有文本被选中，会先删除选中内容再输入。
        """
        self.sel.TypeText(Text=text)

    def delete_selection(self, unit: int = WD_UNIT_CHARACTER, count: int = 1) -> int:
        """
        删除当前 Selection。

        Returns:
            实际删除的字符数
        """
        return self.sel.Delete(Unit=unit, Count=count)

    def insert_paragraph(self):
        """插入段落（回车）。"""
        self.sel.TypeParagraph()

    def insert_page_break(self):
        """插入分页符。"""
        self.sel.InsertBreak(Type=6)

    def cut_selection(self):
        """剪切当前 Selection 到剪贴板。"""
        self.sel.Cut()

    def copy_selection(self):
        """复制当前 Selection 到剪贴板。"""
        self.sel.Copy()

    def paste_selection(self):
        """粘贴剪贴板内容到当前 Selection 位置。"""
        self.sel.Paste()

    # ========================================================================
    # 辅助
    # ========================================================================

    def _get_type_name(self, type_code: int) -> str:
        """将 Selection 类型常量转为可读名称。"""
        names = {
            WD_NO_SELECTION: "no_selection",
            WD_SELECTION: "selection",
            WD_LINE_SELECTION: "line_selection",
            WD_CHARACTER_SELECTION: "character_selection",
            WD_WORD_SELECTION: "word_selection",
            WD_SENTENCE_SELECTION: "sentence_selection",
            WD_PARAGRAPH_SELECTION: "paragraph_selection",
            WD_BLOCK_SELECTION: "block_selection",
            WD_INLINE_SHAPE_SELECTION: "inline_shape_selection",
            WD_SHAPE_SELECTION: "shape_selection",
        }
        return names.get(type_code, f"unknown({type_code})")

    def set_range_from_selection(self):
        """将 Selection 转为 Range 对象（供其他模块使用）。"""
        return self.sel.Range

    def scroll_to_selection(self):
        """滚动视图使 Selection 可见。"""
        self.sel.Range.GoTo()
