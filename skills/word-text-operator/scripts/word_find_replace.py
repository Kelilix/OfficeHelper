# -*- coding: utf-8 -*-
"""
Word 查找与替换模块

封装 Find / Replace 的全部能力：
- 纯文本查找与替换
- 带格式的查找与替换
- 通配符查找（Wildcards）
- 正则表达式查找
- 遍历所有匹配项
- 批量查找多个关键词

优先使用 Range.Find，Selection.Find 仅用于用户交互场景。
"""

from __future__ import annotations
from typing import TYPE_CHECKING, Iterator, Optional, Callable, List

if TYPE_CHECKING:
    from win32com.client import CDispatch


# wdFind 常量
WD_FIND_STOP = 0
WD_FIND_REPLACE_ONE = 1
WD_FIND_REPLACE_ALL = 2
WD_FIND_REPLACE_ASK = 3

# wdFindWrap 常量
WD_FIND_WRAP_NONE = 0      # 不循环，查到末尾停止
WD_FIND_WRAP_STORY = 1     # 循环到起点
WD_FIND_WRAP_CONTINUE = 2  # 显示提示

# wdSearchDirection
WD_SEARCH_BACKWARD = 0
WD_SEARCH_FORWARD = 1


class FindReplace:
    """Word 查找与替换操作封装。"""

    def __init__(self, word_base):
        self._wb = word_base

    # ========================================================================
    # 基础查找
    # ========================================================================

    def _get_find(self, rng_or_sel) -> "CDispatch":
        """获取 Find 对象。"""
        return rng_or_sel.Find

    def _setup_find(
        self,
        find_obj: "CDispatch",
        text: str,
        whole_word: bool = False,
        match_case: bool = False,
        forward: bool = True,
        wrap: int = WD_FIND_WRAP_NONE,
        find_again: bool = False,
    ):
        """配置 Find 对象的通用属性。"""
        if not find_again:
            find_obj.ClearFormatting()
        find_obj.Text = text
        find_obj.WholeWord = whole_word
        find_obj.MatchCase = match_case
        find_obj.Forward = forward
        find_obj.Wrap = wrap
        find_obj.Format = False

    def find_in_range(
        self,
        rng: "CDispatch",
        text: str,
        whole_word: bool = False,
        match_case: bool = False,
        forward: bool = True,
        wrap: int = WD_FIND_WRAP_NONE,
    ) -> bool:
        """
        在 Range 内查找文本（单次）。

        Returns:
            是否找到
        """
        find = rng.Find
        self._setup_find(find, text, whole_word, match_case, forward, wrap)
        return bool(find.Execute())

    def find_next_in_range(
        self,
        rng: "CDispatch",
        text: str,
        whole_word: bool = False,
        match_case: bool = False,
    ) -> Optional["CDispatch"]:
        """
        在 Range 内查找第一个匹配项，返回匹配内容的 Range。

        找到后 Range 自动变为匹配区域。
        """
        find = rng.Find
        self._setup_find(
            find, text, whole_word, match_case, forward=True, wrap=WD_FIND_WRAP_NONE
        )
        if find.Execute():
            return rng.Duplicate
        return None

    # ========================================================================
    # 替换
    # ========================================================================

    def replace_in_range(
        self,
        rng: "CDispatch",
        find_text: str,
        replace_text: str,
        whole_word: bool = False,
        match_case: bool = False,
        replace_all: bool = True,
    ) -> int:
        """
        在 Range 内替换文本。

        Args:
            rng: 要执行替换的 Range
            find_text: 要查找的文本
            replace_text: 替换为的文本
            whole_word: 全字匹配
            match_case: 区分大小写
            replace_all: True=全部替换，False=只替换第一个

        Returns:
            替换次数
        """
        find = rng.Find
        self._setup_find(find, find_text, whole_word, match_case)
        find.Replacement.ClearFormatting()
        find.Replacement.Text = replace_text

        count = 0
        if replace_all:
            while find.Execute(Replace=WD_FIND_REPLACE_ALL):
                count += 1
                if not find.Parent.Find.Wrap:
                    break
            return count
        else:
            if find.Execute(Replace=WD_FIND_REPLACE_ONE):
                return 1
            return 0

    def replace_in_document(
        self,
        find_text: str,
        replace_text: str,
        whole_word: bool = False,
        match_case: bool = False,
    ) -> int:
        """
        在整个文档中替换。

        Returns:
            替换次数
        """
        rng = self._wb.document.Range(
            Start=0, End=self._wb.document.Characters.Count
        )
        return self.replace_in_range(
            rng, find_text, replace_text, whole_word, match_case, replace_all=True
        )

    def replace_in_selection(
        self,
        find_text: str,
        replace_text: str,
        whole_word: bool = False,
        match_case: bool = False,
        replace_all: bool = True,
    ) -> int:
        """
        在当前 Selection 中替换。

        Returns:
            替换次数
        """
        sel = self._wb.selection
        find = sel.Find
        self._setup_find(find, find_text, whole_word, match_case)
        find.Replacement.ClearFormatting()
        find.Replacement.Text = replace_text

        if replace_all:
            return find.Execute(Replace=WD_FIND_REPLACE_ALL)
        else:
            return find.Execute(Replace=WD_FIND_REPLACE_ONE)

    # ========================================================================
    # 遍历所有匹配（Generator）
    # ========================================================================

    def find_all_in_range(
        self,
        rng: "CDispatch",
        text: str,
        whole_word: bool = False,
        match_case: bool = False,
        backward: bool = False,
    ) -> Iterator["CDispatch"]:
        """
        遍历 Range 内所有匹配项的 Range（Generator）。

        Usage:
            for match in find_all_in_range(rng, "关键词"):
                print(match.Text)

        注意：遍历过程中对文档的修改会导致索引失效，请先收集所有位置。
        """
        doc_len = self._wb.document.Characters.Count
        find = rng.Find
        self._setup_find(
            find,
            text,
            whole_word,
            match_case,
            forward=not backward,
            wrap=WD_FIND_WRAP_STORY,
        )

        found_count = 0
        while find.Execute():
            found_count += 1
            yield rng.Duplicate

        return found_count

    def find_all_positions(
        self,
        rng: "CDispatch",
        text: str,
        whole_word: bool = False,
        match_case: bool = False,
    ) -> List[dict]:
        """
        收集 Range 内所有匹配的 (start, end, text) 信息。

        适合在文档修改前保存位置快照。
        """
        positions = []
        for match in self.find_all_in_range(rng, text, whole_word, match_case):
            positions.append(
                {"start": match.Start, "end": match.End, "text": match.Text}
            )
        return positions

    # ========================================================================
    # 通配符 / 正则 / 格式查找
    # ========================================================================

    def find_wildcards_in_range(
        self, rng: "CDispatch", pattern: str, replace_text: Optional[str] = None
    ) -> int:
        """
        在 Range 内使用通配符查找和替换。

        常用通配符：
            ?   任意单个字符
            *   任意多个字符
            [abc]  括号内任意字符
            {n}   精确出现 n 次
            {n,m} 出现 n 到 m 次
            <    词开头
            >    词结尾

        Args:
            pattern: 通配符模式
            replace_text: 替换文本（可选）

        Returns:
            替换次数（如果提供了 replace_text），否则返回 0
        """
        find = rng.Find
        find.ClearFormatting()
        find.Text = pattern
        find.MatchWildcards = True
        find.Forward = True
        find.Wrap = WD_FIND_WRAP_NONE

        if replace_text is not None:
            find.Replacement.ClearFormatting()
            find.Replacement.Text = replace_text
            count = 0
            while find.Execute(Replace=WD_FIND_REPLACE_ALL):
                count += 1
            return count
        else:
            return 1 if find.Execute() else 0

    def find_with_format_in_range(
        self,
        rng: "CDispatch",
        text: str,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        underline: Optional[int] = None,
        font_name: Optional[str] = None,
        font_size: Optional[float] = None,
        font_color: Optional[int] = None,
        highlight: Optional[int] = None,
        case: Optional[int] = None,
    ) -> bool:
        """
        带格式约束的查找。

        Args:
            text: 要查找的文本
            bold: 是否加粗（None 表示不限制）
            italic: 是否斜体
            underline: 下划线类型
            font_name: 字体名称
            font_size: 字号
            font_color: 颜色
            highlight: 高亮色
            case: 大小写类型

        Returns:
            是否找到
        """
        find = rng.Find
        find.ClearFormatting()
        find.Text = text
        find.Forward = True
        find.Wrap = WD_FIND_WRAP_NONE

        if bold is not None:
            find.Font.Bold = -1 if bold else 0
        if italic is not None:
            find.Font.Italic = -1 if italic else 0
        if underline is not None:
            find.Font.Underline = underline
        if font_name is not None:
            find.Font.Name = font_name
        if font_size is not None:
            find.Font.Size = font_size
        if font_color is not None:
            find.Font.Color = font_color
        if highlight is not None:
            find.Font.Highlight = highlight

        return bool(find.Execute())

    def replace_with_format(
        self,
        rng: "CDispatch",
        find_text: str,
        replace_text: str,
        **font_kwargs,
    ) -> int:
        """
        替换文本，并将替换后文本设置为指定格式。

        font_kwargs 支持: bold, italic, underline, font_name, font_size,
                          font_color, highlight 等（同 find_with_format_in_range）。
        """
        find = rng.Find
        find.ClearFormatting()
        find.Text = find_text
        find.Forward = True
        find.Wrap = WD_FIND_WRAP_NONE

        find.Replacement.ClearFormatting()
        find.Replacement.Text = replace_text

        for attr, value in font_kwargs.items():
            if value is not None:
                setattr(find.Replacement.Font, attr.capitalize().replace("_", ""), value)

        count = 0
        while find.Execute(Replace=WD_FIND_REPLACE_ALL):
            count += 1
        return count

    # ========================================================================
    # 批量查找
    # ========================================================================

    def batch_find(
        self,
        rng: "CDispatch",
        queries: List[str],
        whole_word: bool = False,
        match_case: bool = False,
    ) -> dict:
        """
        批量查找多个关键词，返回每个词在 Range 内的出现次数。

        Args:
            queries: 关键词列表
            whole_word: 全字匹配
            match_case: 区分大小写

        Returns:
            { "关键词1": count1, "关键词2": count2, ... }
        """
        result = {}
        doc_len = self._wb.document.Characters.Count

        for query in queries:
            count = 0
            test_rng = self._wb.document.Range(Start=rng.Start, End=rng.End)
            find = test_rng.Find
            self._setup_find(find, query, whole_word, match_case)
            while find.Execute(Replace=0):
                count += 1

            result[query] = count

        return result

    def batch_replace(
        self,
        rng: "CDispatch",
        replacements: dict,
        whole_word: bool = False,
        match_case: bool = False,
    ) -> dict:
        """
        批量替换多个词对。

        Args:
            replacements: { "原词": "新词", ... }
            whole_word: 全字匹配
            match_case: 区分大小写

        Returns:
            每个词的替换次数
        """
        result = {}
        for old, new in replacements.items():
            result[old] = self.replace_in_range(
                rng, old, new, whole_word, match_case, replace_all=True
            )
        return result

    # ========================================================================
    # 常用快捷操作
    # ========================================================================

    def highlight_all(
        self, rng: "CDispatch", text: str, highlight_color: int = 7
    ) -> int:
        """
        将 Range 内所有匹配的文本高亮显示。

        Args:
            highlight_color: wdHighlightColor 常量，默认黄色(7)
        """
        find = rng.Find
        self._setup_find(find, text)
        find.Replacement.ClearFormatting()
        find.Replacement.Font.Highlight = highlight_color

        count = 0
        while find.Execute(Replace=WD_FIND_REPLACE_ALL):
            count += 1
        return count

    def replace_paragraph_marks(self, rng: "CDispatch", separator: str = " "):
        """
        将段落标记替换为指定分隔符（段落合并时的文本清理）。

        Args:
            separator: 替换段落标记的字符，默认空格
        """
        find = rng.Find
        find.ClearFormatting()
        find.Text = "^p"  # 段落标记
        find.Replacement.ClearFormatting()
        find.Replacement.Text = separator
        find.Forward = True
        find.Wrap = WD_FIND_WRAP_NONE
        find.Execute(Replace=WD_FIND_REPLACE_ALL)

    def wrap_matched_pairs(
        self,
        rng: "CDispatch",
        open_marker: str,
        close_marker: str,
        style_filter: Optional[str] = None,
    ) -> int:
        """
        在匹配文本两端加标记（如加粗括号、加引号）。

        Args:
            open_marker: 前置标记（如 "【"）
            close_marker: 后置标记（如 "】"）
            style_filter: 如果指定样式名，只处理该样式的文本

        Returns:
            处理次数
        """
        count = 0
        for match in list(self.find_all_in_range(rng, "*")):
            if style_filter and match.Style.NameLocal != style_filter:
                continue
            original = match.Text
            match.Text = open_marker + original + close_marker
            count += 1
        return count

    def count_matches(self, rng: "CDispatch", text: str) -> int:
        """统计 Range 内某文本出现的次数。"""
        return sum(1 for _ in self.find_all_in_range(rng, text))
