# -*- coding: utf-8 -*-
"""
Word COM 基础连接器

封装 pywin32 的 Word COM 对象创建与生命周期管理。
所有其他模块都依赖于此模块。
"""

from __future__ import annotations

import sys
import os
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    import win32com.client
    from win32com.client import CDispatch

# 模块级标记：CoInitialize 是否已在本进程执行过（避免重复调用报错）
_COINIT_DONE = False


class WordBase:
    """Word COM 基础连接器，负责 Word 应用程序和文档的创建与管理。"""

    _instance: Optional["WordBase"] = None

    def __init__(
        self,
        visible: bool = True,
        display_alerts: bool = False,
        screen_updating: bool = True,
    ):
        """
        初始化 Word COM 连接器。

        Args:
            visible: 是否显示 Word 窗口
            display_alerts: 是否显示警告对话框
            screen_updating: 是否刷新屏幕
        """
        self._word_app: Optional["CDispatch"] = None
        self._word_app_com: Optional["win32com.client.VARIANT"] = None
        self._document: Optional["CDispatch"] = None
        self._visible = visible
        self._display_alerts = display_alerts
        self._screen_updating = screen_updating
        self._owned = False  # 是否由本实例创建的 Word 进程

    # -------------------------------------------------------------------------
    # COM 初始化
    # -------------------------------------------------------------------------

    def connect(self, doc_path: Optional[str] = None) -> "WordBase":
        """
        连接到 Word 应用程序（复用已有进程或新建）。

        Args:
            doc_path: 要打开的文档路径，如果为 None 则不打开文档

        Returns:
            self
        """
        import win32com.client
        import pythoncom

        global _COINIT_DONE
        if not _COINIT_DONE:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
            _COINIT_DONE = True

        try:
            self._word_app = win32com.client.Dispatch("Word.Application")
        except Exception:
            try:
                self._word_app = win32com.client.Dispatch("Ketdel.Application.8")
            except Exception:
                raise RuntimeError("无法启动或连接到 Word 应用程序，请确保已安装 Microsoft Word")

        try:
            self._word_app.Visible = self._visible
            self._word_app.DisplayAlerts = 0 if self._display_alerts else 2
            self._word_app.ScreenUpdating = self._screen_updating
        except Exception:
            pass

        self._owned = True

        if doc_path:
            self.open_document(doc_path)

        return self

    def disconnect(self, save_changes: bool = False, doc_path: Optional[str] = None):
        """
        断开与 Word 的连接。

        Args:
            save_changes: 是否保存更改
            doc_path: 保存路径（如果不想使用原路径）
        """
        if self._document:
            if save_changes and self._document.Saved is False:
                self._document.Save()
            self._document.Close(SaveChanges=save_changes)
            self._document = None

        if self._word_app and self._owned:
            try:
                self._word_app.Quit()
            except Exception:
                pass
            self._word_app = None

    def __enter__(self) -> "WordBase":
        return self.connect()

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect(save_changes=(exc_type is None))
        return False

    # -------------------------------------------------------------------------
    # 应用程序对象
    # -------------------------------------------------------------------------

    @property
    def word_app(self) -> "CDispatch":
        """返回 Word Application 对象。"""
        if self._word_app is None:
            raise RuntimeError("Word 未连接，请先调用 connect()")
        return self._word_app

    @property
    def is_connected(self) -> bool:
        """检查是否已连接到 Word。"""
        return self._word_app is not None

    # -------------------------------------------------------------------------
    # 文档操作
    # -------------------------------------------------------------------------

    @property
    def document(self) -> "CDispatch":
        """返回当前活动文档对象。"""
        if self._document is None:
            self._document = self.word_app.ActiveDocument
        return self._document

    @property
    def active_document_name(self) -> str:
        """返回当前活动文档的名称。"""
        return self.document.Name

    def open_document(self, path: str, read_only: bool = False) -> "CDispatch":
        """
        打开一个 Word 文档。

        Args:
            path: 文档路径
            read_only: 是否以只读模式打开

        Returns:
            Document 对象
        """
        self._document = self.word_app.Documents.Open(path, ReadOnly=read_only)
        return self._document

    def new_document(self) -> "CDispatch":
        """创建一个新文档。"""
        self._document = self.word_app.Documents.Add()
        return self._document

    def save_document(self, path: Optional[str] = None):
        """
        保存当前文档。

        Args:
            path: 可选，指定保存路径
        """
        if path:
            self.document.SaveAs(path)
        else:
            self.document.Save()

    # -------------------------------------------------------------------------
    # Selection & Range 快捷访问
    # -------------------------------------------------------------------------

    @property
    def selection(self) -> "CDispatch":
        """返回 Selection 对象。"""
        try:
            return self.word_app.Selection
        except Exception as e:
            raise RuntimeError(
                f"无法获取 Word Selection（Word 可能已关闭或 COM 连接已失效）: {e}"
            ) from None

    @property
    def range(self) -> "CDispatch":
        """返回 Selection.CreateMethod 对象（等价于 Selection.Range）。"""
        return self.selection.Range

    def get_range(self, start: int, end: int) -> "CDispatch":
        """
        根据字符位置创建 Range 对象。

        Args:
            start: 起始字符位置
            end: 结束字符位置

        Returns:
            Range 对象
        """
        return self.document.Range(Start=start, End=end)

    def get_paragraph_range(self, para_index: int) -> "CDispatch":
        """
        获取指定段落的 Range 对象。

        Args:
            para_index: 段落索引（从 0 开始）

        Returns:
            Range 对象
        """
        return self.document.Paragraphs(para_index + 1).Range

    # -------------------------------------------------------------------------
    # 查找 / 替换 / 执行 统一入口
    # -------------------------------------------------------------------------

    def execute_find(
        self,
        text: str,
        find_text: str,
        replace_text: Optional[str] = None,
        whole_word: bool = False,
        match_case: bool = False,
        forward: bool = True,
        wrap: int = 0,
    ) -> bool:
        """
        在当前 Selection 中执行 Find（或 Find+Replace）。

        Args:
            text: 要查找的文本
            find_text: 同 text（向后兼容）
            replace_text: 替换文本（None 表示仅查找）
            whole_word: 全字匹配
            match_case: 区分大小写
            forward: 向前查找
            wrap: 查找循环方式（0=不循环，1=循环）

        Returns:
            是否找到
        """
        find = self.selection.Find
        find.ClearFormatting()
        find.Text = find_text or text
        find.WholeWord = whole_word
        find.MatchCase = match_case
        find.Forward = forward
        find.Wrap = wrap

        if replace_text is not None:
            find.Replacement.ClearFormatting()
            find.Replacement.Text = replace_text
            return find.Execute(Replace=2)  # wdReplaceAll = 2
        else:
            return find.Execute()

    # -------------------------------------------------------------------------
    # 通用工具
    # -------------------------------------------------------------------------

    @staticmethod
    def rgb_to_int(r: int, g: int, b: int) -> int:
        """将 RGB 转为 Word 颜色整数值。"""
        return (r << 16) | (g << 8) | b

    @staticmethod
    def int_to_rgb(color: int) -> tuple[int, int, int]:
        """将 Word 颜色整数值转为 RGB 元组。"""
        r = (color >> 16) & 0xFF
        g = (color >> 8) & 0xFF
        b = color & 0xFF
        return r, g, b
