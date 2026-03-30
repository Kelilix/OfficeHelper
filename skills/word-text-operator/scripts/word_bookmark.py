# -*- coding: utf-8 -*-
"""
Word 书签操作模块

提供书签的创建、读取、更新、删除能力。
支持命名书签、隐藏书签、快速书签（Quick Bookmarks）。
也包含 Range 级别的书签化操作（MakeEditable、MakeUnavailable 等）。
"""

from __future__ import annotations
from typing import TYPE_CHECKING, Optional, List, Dict

if TYPE_CHECKING:
    from win32com.client import CDispatch


class BookmarkOperator:
    """Word 书签操作封装。"""

    def __init__(self, word_base):
        self._wb = word_base

    # ========================================================================
    # 创建
    # ========================================================================

    def create(
        self,
        rng: "CDispatch",
        name: str,
        add_to_quick_access: bool = False,
    ) -> Optional["CDispatch"]:
        """
        在 Range 处创建书签。

        Args:
            rng: 要添加书签的 Range
            name: 书签名（唯一，不能包含空格和特殊字符）
            add_to_quick_access: 是否添加到快速访问工具栏

        Returns:
            书签对象，失败返回 None
        """
        name = name.strip()
        if not name:
            return None

        try:
            bm = self._wb.document.Bookmarks.Add(Name=name, Range=rng)
            return bm
        except Exception:
            return None

    def create_at_selection(self, name: str) -> Optional["CDispatch"]:
        """在当前 Selection 处创建书签。"""
        return self.create(self._wb.selection.Range, name)

    def create_quick_bookmark(self, rng: "CDispatch", label: str) -> Optional["CDispatch"]:
        """
        创建快速书签（无名称，用 Label 标识，供后续定位）。

        实际上 Word 标准书签必须有名称，此方法将 label 转为合法书签名。
        """
        safe_name = label.replace(" ", "_").replace("/", "_")
        return self.create(rng, safe_name)

    # ========================================================================
    # 读取
    # ========================================================================

    def list_all(self) -> List[Dict[str, any]]:
        """
        返回文档中所有书签的列表。

        Returns:
            [{ "name": str, "start": int, "end": int }, ...]
        """
        result = []
        for bm in self._wb.document.Bookmarks:
            result.append(
                {
                    "name": bm.Name,
                    "start": bm.Range.Start,
                    "end": bm.Range.End,
                    "text": bm.Range.Text,
                }
            )
        return result

    def get(self, name: str) -> Optional["CDispatch"]:
        """
        获取指定书签对象。

        Args:
            name: 书签名

        Returns:
            Bookmark 对象，未找到返回 None
        """
        try:
            return self._wb.document.Bookmarks(name)
        except Exception:
            return None

    def exists(self, name: str) -> bool:
        """检查书签是否存在。"""
        return self._wb.document.Bookmarks.Exists(name)

    def get_range(self, name: str) -> Optional["CDispatch"]:
        """获取书签对应的 Range。"""
        bm = self.get(name)
        return bm.Range if bm else None

    def get_text(self, name: str) -> str:
        """获取书签内的文本内容。"""
        bm = self.get(name)
        return bm.Range.Text if bm else ""

    # ========================================================================
    # 更新 / 移动
    # ========================================================================

    def update_range(self, name: str, new_start: int, new_end: int) -> bool:
        """
        重新定位书签到新的 Range。

        Args:
            name: 书签名
            new_start: 新起始位置
            new_end: 新结束位置

        Returns:
            是否成功
        """
        bm = self.get(name)
        if not bm:
            return False
        try:
            new_rng = self._wb.document.Range(Start=new_start, End=new_end)
            bm.Range = new_rng
            return True
        except Exception:
            return False

    def rename(self, old_name: str, new_name: str) -> bool:
        """
        重命名书签。

        Returns:
            是否成功
        """
        bm = self.get(old_name)
        if not bm:
            return False
        try:
            bm.Name = new_name
            return True
        except Exception:
            return False

    def select(self, name: str) -> bool:
        """
        选中文书签所在的内容（跳转到书签）。

        Returns:
            是否成功
        """
        bm = self.get(name)
        if bm:
            bm.Range.Select()
            return True
        return False

    # ========================================================================
    # 删除
    # ========================================================================

    def delete(self, name: str) -> bool:
        """
        删除指定书签（不删除书签内的文本内容）。

        Returns:
            是否成功
        """
        bm = self.get(name)
        if bm:
            try:
                bm.Delete()
                return True
            except Exception:
                return False
        return False

    def delete_all(self):
        """删除所有书签（保留内容）。"""
        while self._wb.document.Bookmarks.Count > 0:
            self._wb.document.Bookmarks(1).Delete()

    def delete_in_range(self, rng: "CDispatch"):
        """
        删除 Range 内的所有书签。
        删除后 Range 内的书签被移除，但内容不变。
        """
        bookmarks_to_delete = [
            bm.Name
            for bm in self._wb.document.Bookmarks
            if rng.InRange(bm.Range)
        ]
        for name in bookmarks_to_delete:
            self.delete(name)

    # ========================================================================
    # Range 书签化操作
    # ========================================================================

    def make_range_editable(self, rng: "CDispatch", reader: bool = False):
        """
        标记 Range 为可编辑区域（窗体保护模式下有用）。

        Args:
            reader: 是否仅允许阅读（不可编辑）
        """
        rng.MakeEditabledual = 0 if reader else 1

    def make_unavailable(self, rng: "CDispatch"):
        """标记 Range 为不可用区域（灰显）。"""
        rng.MakeUnAvailable10()

    # ========================================================================
    # 导航辅助
    # ========================================================================

    def navigate_by_bookmark(self, name: str) -> Optional["CDispatch"]:
        """
        跳转到书签位置，返回书签的 Range（也可用于 GoTo）。

        Returns:
            书签 Range，未找到返回 None
        """
        return self.get_range(name)

    def get_bookmark_info(self, name: str) -> Optional[dict]:
        """
        获取书签详细信息。

        Returns:
            {"name", "start", "end", "text", "story_type"}
        """
        bm = self.get(name)
        if not bm:
            return None
        rng = bm.Range
        return {
            "name": bm.Name,
            "start": rng.Start,
            "end": rng.End,
            "text": rng.Text,
            "story_type": rng.StoryType,
        }

    def export_bookmarks(self, filepath: str):
        """
        将所有书签导出为 JSON 文件。

        Args:
            filepath: 输出文件路径
        """
        import json

        data = self.list_all()
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def import_bookmarks(self, filepath: str) -> int:
        """
        从 JSON 文件导入书签（根据 start/end 位置重建书签）。

        Args:
            filepath: 书签数据文件路径

        Returns:
            成功导入的书签数量
        """
        import json

        with open(filepath, "r", encoding="utf-8") as f:
            data = json.load(f)

        count = 0
        for item in data:
            try:
                rng = self._wb.document.Range(
                    Start=item["start"], End=item["end"]
                )
                self.create(rng, item["name"])
                count += 1
            except Exception:
                pass

        return count
