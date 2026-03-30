# -*- coding: utf-8 -*-
"""
WordTextOperator - Word 文本操作统一入口

整合所有子模块，提供面向 AI Agent 的统一 API。
同时兼容 context-manager 用法（with 语句）。

Usage:
    from word_text_operator import WordTextOperator

    with WordTextOperator("document.docx") as op:
        # 获取文档全文
        print(op.get_full_text())

        # 查找关键词
        rng = op.find("目标文本")
        if rng:
            op.set_bold(rng, True)
            op.set_font_color(rng, "red")

        # 替换
        op.replace("旧文本", "新文本")
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Optional, Union

if TYPE_CHECKING:
    from win32com.client import CDispatch

# Import submodules
from .word_base import WordBase
from .word_range_navigation import RangeNavigator
from .word_text_operations import TextOperator
from .word_format import TextFormatter
from .word_find_replace import FindReplace
from .word_bookmark import BookmarkOperator
from .word_selection import SelectionOperator


class WordTextOperator:
    """
    Word 文本操作统一封装。

    整合 Range、Selection、Find、Bookmark 四大操作域，
    提供简洁的 API 供 AI Agent 调用。

    使用方式：
        # 方式1：context manager（推荐）
        with WordTextOperator("doc.docx") as op:
            ...

        # 方式2：手动管理
        op = WordTextOperator()
        op.connect("doc.docx")
        ...
        op.disconnect()
    """

    def __init__(
        self,
        doc_path: Optional[str] = None,
        visible: bool = True,
        display_alerts: bool = False,
    ):
        """
        初始化操作器。

        Args:
            doc_path: 要打开的文档路径，None 表示不打开文档
            visible: 是否显示 Word 窗口
            display_alerts: 是否显示警告对话框
        """
        self._base = WordBase(visible=visible, display_alerts=display_alerts)
        self._nav: Optional[RangeNavigator] = None
        self._text: Optional[TextOperator] = None
        self._fmt: Optional[TextFormatter] = None
        self._find: Optional[FindReplace] = None
        self._bm: Optional[BookmarkOperator] = None
        self._sel: Optional[SelectionOperator] = None

        if doc_path:
            self.connect(doc_path)

    # ========================================================================
    # 生命周期管理
    # ========================================================================

    def connect(self, doc_path: Optional[str] = None) -> "WordTextOperator":
        """
        连接到 Word 应用程序。

        Args:
            doc_path: 要打开的文档路径

        Returns:
            self
        """
        self._base.connect(doc_path)
        self._init_submodules()
        return self

    def disconnect(self, save_changes: bool = False):
        """断开连接并关闭 Word。"""
        self._base.disconnect(save_changes=save_changes)
        self._nav = None
        self._text = None
        self._fmt = None
        self._find = None
        self._bm = None
        self._sel = None

    def __enter__(self) -> "WordTextOperator":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect(save_changes=(exc_type is None))
        return False

    def _init_submodules(self):
        """延迟初始化所有子模块（失败时记录警告，不整体崩溃）。"""
        import logging
        logger = logging.getLogger(__name__)
        try:
            self._nav = RangeNavigator(self._base)
        except Exception as e:
            logger.warning("RangeNavigator 初始化失败: %s", e)
        try:
            self._text = TextOperator(self._base)
        except Exception as e:
            logger.warning("TextOperator 初始化失败: %s", e)
        try:
            self._fmt = TextFormatter(self._base)
        except Exception as e:
            logger.warning("TextFormatter 初始化失败: %s", e)
        try:
            self._find = FindReplace(self._base)
        except Exception as e:
            logger.warning("FindReplace 初始化失败: %s", e)
        try:
            self._bm = BookmarkOperator(self._base)
        except Exception as e:
            logger.warning("BookmarkOperator 初始化失败: %s", e)
        try:
            self._sel = SelectionOperator(self._base)
        except Exception as e:
            logger.warning("SelectionOperator 初始化失败: %s", e)

    # ========================================================================
    # 快捷属性
    # ========================================================================

    @property
    def word_app(self):
        """Word Application 对象。"""
        return self._base.word_app

    @property
    def document(self):
        """当前文档对象。"""
        return self._base.document

    @property
    def base(self) -> WordBase:
        """基础连接器。"""
        return self._base

    @property
    def nav(self) -> RangeNavigator:
        """Range 导航模块。"""
        return self._nav

    @property
    def text(self) -> TextOperator:
        """文本操作模块。"""
        return self._text

    @property
    def fmt(self) -> TextFormatter:
        """格式操作模块。"""
        return self._fmt

    @property
    def find(self) -> FindReplace:
        """查找替换模块。"""
        return self._find

    @property
    def bm(self) -> BookmarkOperator:
        """书签操作模块。"""
        return self._bm

    @property
    def sel(self) -> SelectionOperator:
        """Selection 操作模块。"""
        return self._sel

    # ========================================================================
    # 统一 API（高频操作，直接暴露）
    # ========================================================================

    # -- 文本读取 --
    def get_text(self, start: int, end: int) -> str:
        """读取指定字符范围 [start, end) 的文本。"""
        rng = self._base.get_range(start, end)
        return rng.Text

    def get_full_text(self) -> str:
        """读取整个文档的文本。"""
        return self._base.document.Content.Text

    def get_selection_text(self) -> str:
        """读取当前选中的文本。"""
        return self._sel.selection_text

    # -- 查找 --
    def find(self, text: str, whole_word: bool = False, match_case: bool = False) -> Optional["CDispatch"]:
        """
        在文档中查找第一个匹配项。

        Returns:
            匹配 Range，未找到返回 None
        """
        rng = self._base.document.Content
        return self._find.find_next_in_range(rng, text, whole_word, match_case)

    def find_all(self, text: str) -> list:
        """查找所有匹配项，返回位置信息列表。"""
        rng = self._base.document.Content
        return self._find.find_all_positions(rng, text)

    def count_occurrences(self, text: str) -> int:
        """统计关键词在文档中的出现次数。"""
        return self._find.count_matches(self._base.document.Content, text)

    # -- 替换 --
    def replace(
        self,
        find_text: str,
        replace_text: str,
        whole_word: bool = False,
        match_case: bool = False,
        replace_all: bool = True,
    ) -> int:
        """
        替换文档中的文本。

        Returns:
            替换次数
        """
        return self._find.replace_in_document(
            find_text, replace_text, whole_word, match_case
        )

    # -- 书签 --
    def create_bookmark(self, name: str, start: int, end: int) -> bool:
        """在指定范围创建书签。"""
        rng = self._base.get_range(start, end)
        return self._bm.create(rng, name) is not None

    def go_to_bookmark(self, name: str) -> Optional["CDispatch"]:
        """跳转到指定书签。"""
        return self._bm.navigate_by_bookmark(name)

    def get_bookmarks(self) -> list:
        """获取所有书签列表。"""
        return self._bm.list_all()

    # -- 格式（便捷方法） --
    def set_bold(self, rng: "CDispatch", bold: bool = True):
        """设置 Range 内文本为加粗。"""
        self._fmt.set_bold(rng, bold)

    def set_italic(self, rng: "CDispatch", italic: bool = True):
        """设置 Range 内文本为斜体。"""
        self._fmt.set_italic(rng, italic)

    def set_underline(self, rng: "CDispatch", underline: Union[str, int] = "single"):
        """设置 Range 内文本下划线。"""
        self._fmt.set_underline(rng, underline)

    def set_font_color(self, rng: "CDispatch", color: Union[str, int]):
        """设置 Range 内文本颜色。"""
        self._fmt.set_font_color(rng, color)

    def set_font_name(self, rng: "CDispatch", name: str):
        """设置 Range 内文本字体。"""
        self._fmt.set_font_name(rng, name)

    def set_font_size(self, rng: "CDispatch", size: float):
        """设置 Range 内文本字号。"""
        self._fmt.set_font_size(rng, size)

    def set_paragraph_alignment(self, rng: "CDispatch", align: Union[str, int]):
        """设置 Range 内段落对齐。"""
        self._fmt.set_alignment(rng, align)

    def set_highlight(self, rng: "CDispatch", color: Union[str, int] = "yellow"):
        """设置 Range 内文本高亮。"""
        self._fmt.set_highlight(rng, color)

    # -- 范围工具 --
    def get_range(self, start: int, end: int) -> "CDispatch":
        """获取指定字符范围的 Range 对象。"""
        return self._base.get_range(start, end)

    def get_full_range(self) -> "CDispatch":
        """获取整个文档的 Range 对象。"""
        return self._nav.get_full_range()

    def get_selection_range(self) -> "CDispatch":
        """获取当前 Selection 对应的 Range 对象。"""
        return self._sel.selection_range

    def select(self, rng: "CDispatch"):
        """选中文档中指定的 Range。"""
        rng.Select()

    # -- 插入操作 --
    def insert_text(self, rng: "CDispatch", text: str, before: bool = True) -> "CDispatch":
        """
        在 Range 处插入文本。

        Args:
            rng: 插入位置的参考 Range
            text: 要插入的文本
            before: True=插入到 Range 之前，False=插入到 Range 之后

        Returns:
            插入后的 Range
        """
        if before:
            return self._text.insert_before(rng, text)
        else:
            return self._text.insert_after(rng, text)

    def insert_page_break(self, rng: "CDispatch"):
        """在 Range 处插入分页符。"""
        self._text.insert_page_break(rng)

    # -- 删除操作 --
    def delete_range(self, rng: "CDispatch"):
        """删除 Range 的内容。"""
        rng.Delete()

    def delete_selection(self):
        """删除当前选中的内容。"""
        self._sel.delete_selection()

    # -- 大小写转换 --
    def to_uppercase(self, rng: "CDispatch"):
        """将 Range 文本转为全大写。"""
        self._text.to_uppercase(rng)

    def to_lowercase(self, rng: "CDispatch"):
        """将 Range 文本转为全小写。"""
        self._text.to_lowercase(rng)

    def to_title_case(self, rng: "CDispatch"):
        """将 Range 文本转为标题格式。"""
        self._text.to_title_case(rng)

    # -- 统计 --
    def char_count(self, rng: "CDispatch") -> int:
        """统计字符数。"""
        return self._text.char_count(rng)

    def word_count(self, rng: "CDispatch") -> int:
        """统计单词数。"""
        return self._text.word_count(rng)

    # -- 文档操作 --
    def new_document(self):
        """新建空白文档。"""
        self._base.new_document()
        self._init_submodules()

    def save(self, path: Optional[str] = None):
        """保存文档。"""
        self._base.save_document(path)

    # ========================================================================
    # Range API（透明代理到子模块）
    # ========================================================================

    def expand_to_word(self, rng: "CDispatch") -> int:
        """将 Range 扩展到完整单词。"""
        return self._nav.expand_to_word(rng)

    def expand_to_sentence(self, rng: "CDispatch") -> int:
        """将 Range 扩展到完整句子。"""
        return self._nav.expand_to_sentence(rng)

    def expand_to_paragraph(self, rng: "CDispatch") -> int:
        """将 Range 扩展到完整段落。"""
        return self._nav.expand_to_paragraph(rng)

    def collapse(self, rng: "CDispatch", direction: str = "start"):
        """折叠 Range。"""
        self._nav.collapse(rng, direction)

    def move(self, rng: "CDispatch", unit: int = 4, count: int = 1) -> int:
        """移动 Range。"""
        return self._nav.move(rng, unit, count)

    def compare_ranges(self, rng1: "CDispatch", rng2: "CDispatch") -> int:
        """比较两个 Range 的位置关系。"""
        return self._nav.compare_location(rng1, rng2)

    # ========================================================================
    # 通配符 / 格式查找快捷方法
    # ========================================================================

    def find_wildcards(self, pattern: str, replace_text: Optional[str] = None) -> int:
        """
        使用通配符在文档中查找或替换。

        Args:
            pattern: 通配符模式
            replace_text: 可选，替换文本

        Returns:
            替换次数（提供 replace_text 时），否则 0/1 表示是否找到
        """
        rng = self._base.document.Content
        return self._find.find_wildcards_in_range(rng, pattern, replace_text)

    def find_with_format(
        self, text: str, bold: Optional[bool] = None, italic: Optional[bool] = None
    ) -> bool:
        """带格式约束的查找。"""
        rng = self._base.document.Content
        return self._find.find_with_format_in_range(rng, text, bold=bold, italic=italic)

    def replace_with_format(
        self,
        find_text: str,
        replace_text: str,
        bold: bool = False,
        italic: bool = False,
    ) -> int:
        """替换文本并应用格式。"""
        rng = self._base.document.Content
        return self._find.replace_with_format(
            rng, find_text, replace_text, bold=bold, italic=italic
        )

    def batch_replace(self, replacements: dict) -> dict:
        """
        批量替换。

        Args:
            replacements: { "原词": "新词", ... }

        Returns:
            每个词的替换次数
        """
        rng = self._base.document.Content
        return self._find.batch_replace(rng, replacements)

    # ========================================================================
    # 快捷书签操作
    # ========================================================================

    def bookmark_text(self, name: str, text: str) -> bool:
        """
        查找文本并为其添加书签。

        Args:
            name: 书签名
            text: 要查找并书签化的文本

        Returns:
            是否成功
        """
        rng = self.find(text)
        if rng:
            return self._bm.create(rng, name) is not None
        return False

    def wrap_with_bookmarks(
        self, find_text: str, open_name: str, close_name: str
    ) -> bool:
        """
        查找文本，并在其两侧创建成对的书签。

        Args:
            find_text: 要查找的文本
            open_name: 前置书签名
            close_name: 后置书签名

        Returns:
            是否成功
        """
        rng = self.find(find_text)
        if not rng:
            return False
        start = rng.Start
        end = rng.End
        rng1 = self._base.get_range(start, end)
        rng2 = self._base.get_range(start, end)
        ok1 = self._bm.create(rng1, open_name) is not None
        ok2 = self._bm.create(rng2, close_name) is not None
        return ok1 and ok2
