# -*- coding: utf-8 -*-
"""
Word 文本操作模块

提供 Range/Selection 的文本内容操作能力：
- 读取 / 写入 / 追加文本
- 插入内容（InsertBefore / InsertAfter / InsertFile / InsertParagraph）
- 删除内容（Delete）
- 替换内容（Range.Text / Selection.Text）
- 大小写转换（Case）
- 统计字符 / 单词数
"""

from __future__ import annotations
from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from win32com.client import CDispatch


# wdRestoreLast / wdCharacter
WD_CHARACTER = 1
# wdCase 常量
WD_LOWER_CASE = 0
WD_UPPER_CASE = 1
WD_TITLE_CASE = 2
WD_TOGGLE_CASE = 3
# wdUnits
WD_WORD = 2
WD_SENTENCE = 3
WD_PARAGRAPH = 4


class TextOperator:
    """文本内容操作封装。"""

    def __init__(self, word_base):
        self._wb = word_base

    # ========================================================================
    # 读取
    # ========================================================================

    def get_text(self, rng: "CDispatch") -> str:
        """读取 Range 的纯文本内容。"""
        return rng.Text

    def get_formatted_text(self, rng: "CDispatch") -> str:
        """读取 Range 的带格式文本（包含富文本标记）。"""
        return rng.FormattedText

    def get_selection_text(self) -> str:
        """读取当前 Selection 的文本。"""
        return self._wb.selection.Text

    def get_paragraph_text(self, para_index: int) -> str:
        """读取指定段落的文本（索引从 0 开始）。"""
        return self._wb.document.Paragraphs(para_index + 1).Range.Text

    def get_full_document_text(self) -> str:
        """读取整个文档的文本。"""
        return self._wb.document.Content.Text

    # ========================================================================
    # 写入 / 替换
    # ========================================================================

    def set_text(self, rng: "CDispatch", text: str):
        """用新文本完全替换 Range 的内容。"""
        rng.Text = text

    def replace_text(self, rng: "CDispatch", old: str, new: str) -> int:
        """
        在 Range 内替换文本（单次替换第一个匹配）。

        Returns:
            替换次数
        """
        rng.Text = rng.Text.replace(old, new)
        return 1 if old in rng.Text else 0

    # ========================================================================
    # 插入
    # ========================================================================

    def insert_before(self, rng: "CDispatch", text: str) -> "CDispatch":
        """
        在 Range 之前插入文本，返回新 Range（包含插入的文本）。

        不会影响原 Range 的内容，但会改变文档字符偏移。
        """
        rng.InsertBefore(text)
        return rng

    def insert_after(self, rng: "CDispatch", text: str) -> "CDispatch":
        """
        在 Range 之后插入文本，返回新 Range（包含插入的文本）。
        """
        rng.InsertAfter(text)
        return rng

    def insert_file(
        self,
        rng: "CDispatch",
        file_path: str,
        confirm_conversions: bool = False,
        link: bool = False,
        attachment: bool = False,
    ) -> "CDispatch":
        """
        在 Range 处插入另一个文件的内容。

        Args:
            rng: 插入位置
            file_path: 要插入的文件路径
            confirm_conversions: 是否显示转换确认对话框
            link: 是否以链接方式插入
            attachment: 是否作为附件插入

        Returns:
            插入内容所在的新 Range
        """
        rng.InsertFile(
            file_path,
            ConfirmConversions=confirm_conversions,
            Link=link,
            Attachment=attachment,
        )
        return rng

    def insert_break(self, rng: "CDispatch", break_type: int = 7):
        """
        在 Range 处插入分隔符。

        Args:
            break_type: 分隔符类型
                0 = wdSectionBreakContinuous 连续分节符
                1 = wdSectionBreakNextPage 分节符（下一页）
                2 = wdSectionBreakNextPage 同上
                3 = wdSectionBreakEvenPage 偶数页分节符
                4 = wdSectionBreakOddPage 奇数页分节符
                5 = wdLineBreak 换行符（类似 Shift+Enter）
                6 = wdPageBreak 分页符
                7 = wdColumnBreak 分栏符
        """
        rng.InsertBreak(Type=break_type)

    def insert_page_break(self, rng: "CDispatch"):
        """插入分页符。"""
        rng.InsertBreak(Type=6)

    def insert_paragraph(self, rng: "CDispatch"):
        """插入段落标记（相当于回车）。"""
        rng.InsertParagraph()

    def insert_paragraph_after(self, rng: "CDispatch"):
        """在 Range 后插入段落标记。"""
        rng.InsertParagraphAfter()

    def insert_symbol(
        self,
        rng: "CDispatch",
        character_code: int,
        font_name: Optional[str] = None,
        unicode: bool = False,
    ):
        """
        在 Range 处插入符号。

        Args:
            character_code: 字符代码（ASCII 或 Unicode 码点）
            font_name: 符号字体名称（如 "Wingdings"）
            unicode: character_code 是否为 Unicode 码点
        """
        rng.InsertSymbol(
            CharacterNumber=character_code,
            Font=font_name,
            Unicode=unicode,
        )

    # ========================================================================
    # 删除
    # ========================================================================

    def delete(self, rng: "CDispatch", unit: int = 1, count: int = 1) -> int:
        """
        删除 Range 的内容。

        Args:
            rng: 要删除的 Range
            unit: 删除单位（1=Character, 2=Word, 3=Sentence, 4=Paragraph）
            count: 删除数量

        Returns:
            实际删除的数量
        """
        return rng.Delete(Unit=unit, Count=count)

    def delete_all(self, rng: "CDispatch"):
        """删除 Range 的全部内容（保留段落标记）。"""
        rng.Delete(Unit=WD_CHARACTER, Count=len(rng.Text))

    def clear(self, rng: "CDispatch"):
        """清空 Range 的内容，等价于 set_text(rng, "")。"""
        rng.Text = ""

    def delete_selection(self):
        """删除当前 Selection 的内容。"""
        self._wb.selection.Delete()

    # ========================================================================
    # 大小写转换
    # ========================================================================

    def to_lowercase(self, rng: "CDispatch"):
        """将 Range 内文本转为全小写。"""
        rng.Case = WD_LOWER_CASE

    def to_uppercase(self, rng: "CDispatch"):
        """将 Range 内文本转为全大写。"""
        rng.Case = WD_UPPER_CASE

    def to_title_case(self, rng: "CDispatch"):
        """将 Range 内每个单词首字母大写。"""
        rng.Case = WD_TITLE_CASE

    def to_toggle_case(self, rng: "CDispatch"):
        """大小写互相切换（Hello -> hELLO）。"""
        rng.Case = WD_TOGGLE_CASE

    # ========================================================================
    # 统计
    # ========================================================================

    def char_count(self, rng: "CDispatch", include_spaces: bool = False) -> int:
        """
        统计 Range 内的字符数。

        Args:
            include_spaces: 是否包含空格
        """
        return rng.ComputeStatistics(Stat=2, IncludeFootnotesAndEndnotes=include_spaces)
        # wdStatisticCharacters = 2

    def word_count(self, rng: "CDispatch") -> int:
        """统计 Range 内的单词数。"""
        return rng.ComputeStatistics(Stat=4)
        # wdStatisticWords = 4

    def sentence_count(self, rng: "CDispatch") -> int:
        """统计 Range 内的句子数。"""
        return rng.ComputeStatistics(Stat=3)
        # wdStatisticSentences = 3

    def paragraph_count(self, rng: "CDispatch") -> int:
        """统计 Range 内的段落数。"""
        return rng.ComputeStatistics(Stat=5)
        # wdStatisticParagraphs = 5

    def line_count(self, rng: "CDispatch") -> int:
        """统计 Range 内的行数。"""
        return rng.ComputeStatistics(Stat=1)
        # wdStatisticLines = 1

    # ========================================================================
    # 复制 / 剪切 / 粘贴
    # ========================================================================

    def copy(self, rng: "CDispatch"):
        """复制 Range 内容到剪贴板。"""
        rng.Copy()

    def cut(self, rng: "CDispatch"):
        """剪切 Range 内容到剪贴板。"""
        rng.Cut()

    def paste(self, rng: "CDispatch"):
        """在 Range 处粘贴剪贴板内容（替换 Range）。"""
        rng.Paste()

    def paste_formatted(self, rng: "CDispatch"):
        """粘贴并保留源格式。"""
        rng.PasteAndFormat(16)  # wdPasteDefault = 0; wdFormatPlainText = 22; wdUseDestinationStyles = 21

    # ========================================================================
    # 特殊操作
    # ========================================================================

    def split_range(self, rng: "CDispatch") -> list["CDispatch"]:
        """
        将 Range 按段落拆分为多个子 Range。

        Returns:
            子 Range 列表
        """
        paragraphs = rng.Paragraphs
        result = []
        for i in range(1, paragraphs.Count + 1):
            result.append(paragraphs(i).Range)
        return result

    def split_by_sentence(self, rng: "CDispatch") -> list["CDispatch"]:
        """将 Range 按句子拆分为多个子 Range。"""
        sentences = rng.Sentences
        result = []
        for i in range(1, sentences.Count + 1):
            result.append(sentences(i))
        return result

    def split_by_word(self, rng: "CDispatch") -> list["CDispatch"]:
        """将 Range 按单词拆分为多个子 Range。"""
        words = rng.Words
        result = []
        for i in range(1, words.Count + 1):
            result.append(words(i))
        return result

    def normalize_end_of_paragraph(self, rng: "CDispatch"):
        """
        规范化段落结束符：将换行符替换为段落标记。
        常用于从网页或其他文档粘贴后清理格式。
        """
        rng.Find.ClearFormatting()
        rng.Find.Text = "^l"  # 手动换行符 (Ctrl+Enter)
        rng.Find.Replacement.Text = "^p"
        rng.Find.Forward = True
        rng.Find.Wrap = 0
        rng.Find.Execute(Replace=2)  # wdReplaceAll

    def trim_spaces(self, rng: "CDispatch"):
        """去除 Range 首尾空白字符。"""
        rng.Text = rng.Text.strip()

    def normalize_spaces(self, rng: "CDispatch"):
        """将多个连续空格替换为单个空格。"""
        rng.Find.ClearFormatting()
        rng.Find.Text = "  "  # 两个空格
        rng.Find.Replacement.Text = " "
        rng.Find.Forward = True
        rng.Find.Wrap = 0
        while rng.Find.Execute(Replace=2):
            pass  # 持续替换直到没有匹配
