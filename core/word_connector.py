"""
Word文档连接器模块
通过COM接口操作Word文档
"""

import os
import time
import shutil
from pathlib import Path
from typing import Optional, List, Dict, Any
from dataclasses import dataclass, field
from datetime import datetime


class WordNotFoundError(Exception):
    """Word未安装异常"""
    pass


class WordConnectionError(Exception):
    """Word连接异常"""
    pass


@dataclass
class DocumentInfo:
    """文档信息"""
    path: str
    name: str
    fullname: str
    page_count: int = 0
    word_count: int = 0
    paragraphs_count: int = 0
    active: bool = False


@dataclass
class ParagraphFormat:
    """段落格式"""
    alignment: str = "left"  # left, center, right, justify
    line_spacing: float = 1.0
    first_line_indent: float = 0.0  # 磅
    left_indent: float = 0.0
    right_indent: float = 0.0
    space_before: float = 0.0
    space_after: float = 0.0


@dataclass
class FontFormat:
    """字体格式"""
    name: str = "宋体"
    size: float = 12.0
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: str = "black"  # RGB or name


@dataclass
class PageSetup:
    """页面设置"""
    paper_size: str = "A4"  # A4, Letter, etc.
    orientation: str = "portrait"  # portrait, landscape
    top_margin: float = 2.54  # cm
    bottom_margin: float = 2.54
    left_margin: float = 3.17
    right_margin: float = 3.17
    header_distance: float = 1.5
    footer_distance: float = 1.75


class UndoManager:
    """撤销管理器"""

    def __init__(self, max_undo: int = 20):
        self._undo_stack: List[tuple] = []
        self._redo_stack: List[tuple] = []
        self._max_undo = max_undo

    def push(self, action: str, undo_func, redo_func, *args):
        """压入操作"""
        self._undo_stack.append((action, undo_func, redo_func, args))
        self._redo_stack.clear()
        if len(self._undo_stack) > self._max_undo:
            self._undo_stack.pop(0)

    def undo(self):
        """撤销"""
        if not self._undo_stack:
            return False
        action, undo_func, redo_func, args = self._undo_stack.pop()
        undo_func(*args)
        self._redo_stack.append((action, undo_func, redo_func, args))
        return True

    def redo(self):
        """重做"""
        if not self._redo_stack:
            return False
        action, undo_func, redo_func, args = self._redo_stack.pop()
        redo_func(*args)
        self._undo_stack.append((action, undo_func, redo_func, args))
        return True

    def can_undo(self) -> bool:
        return len(self._undo_stack) > 0

    def can_redo(self) -> bool:
        return len(self._redo_stack) > 0

    def clear(self):
        self._undo_stack.clear()
        self._redo_stack.clear()


class WordConnector:
    """Word文档连接器"""

    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        self._word_app = None
        self._document = None
        self._undo_manager = UndoManager()
        self._backup_dir = Path.home() / ".office_helper" / "backups"
        self._backup_dir.mkdir(parents=True, exist_ok=True)
        self._current_file = None
        self._initialized = True

    def connect(self, visible: bool = True) -> bool:
        """
        连接到Word应用

        Args:
            visible: 是否显示Word窗口

        Returns:
            bool: 连接是否成功
        """
        try:
            import win32com.client
            import pythoncom

            # 初始化COM
            pythoncom.CoInitialize()

            # 创建Word应用实例
            self._word_app = win32com.client.Dispatch("Word.Application")
            self._word_app.Visible = visible
            self._word_app.DisplayAlerts = False

            return True

        except Exception as e:
            print(f"连接Word失败: {e}")
            return False

    def is_connected(self) -> bool:
        """
        检查是否已连接（通过探测 COM 代理有效性）。
        注意：本方法只检查状态，不尝试重新连接（重新连接的逻辑在调用方）。
        """
        if self._word_app is None:
            return False
        try:
            _ = self._word_app.ActiveDocument
            return True
        except Exception:
            self._word_app = None
            return False

    def get_main_window_hwnd(self) -> Optional[int]:
        """
        获取当前 Word 进程主窗口句柄（用于嵌入宿主窗口）。
        即使窗口已被 SetParent 成子窗口，COM 仍返回同一 HWND，FindWindow 则无法找到。
        """
        if not self._word_app:
            return None
        try:
            h = int(self._word_app.Hwnd)
            return h if h else None
        except Exception:
            return None

    def get_embed_target_hwnd(self) -> Optional[int]:
        """
        用于定位/贴附 Word 窗口：优先当前活动文档窗口（Word SDI 下一文档一窗格，比 Application.Hwnd 更准）。
        """
        if not self._word_app:
            return None
        try:
            aw = self._word_app.ActiveWindow
            if aw is not None:
                for name in ("Hwnd", "HWND", "hwnd"):
                    try:
                        h = int(getattr(aw, name))
                        if h:
                            return h
                    except Exception:
                        continue
        except Exception:
            pass
        return self.get_main_window_hwnd()

    def _setup_document_window_for_embed(self) -> None:
        """
        文档打开/创建后调用：禁用「任务栏合并显示 Word 图标」，
        使 Word 不再在任务栏上以独立程序组出现。
        """
        if not self._word_app:
            return
        try:
            self._word_app.ShowWindowsInTaskbar = False
        except Exception as e:
            print(f"设置 ShowWindowsInTaskbar 失败: {e}")

    def _prepare_for_embed(self) -> Optional[int]:
        """
        文档打开后、为嵌入做的最后准备：
        隐藏 OpusApp 主窗口（仅显示文档编辑区），
        返回 ActiveWindow 的 Hwnd（文档窗格）。
        若有多窗口则全部关闭后新建。
        """
        if not self._word_app or not self._document:
            return None

        try:
            # 先尝试 NewWindow（制造一个独立窗口，OpusApp 变空）
            try:
                nw = self._word_app.ActiveWindow.NewWindow()
                self._word_app.ActiveWindow.Close()
            except Exception:
                pass

            hwnd = self.get_embed_target_hwnd()
            return hwnd

        except Exception as e:
            print(f"_prepare_for_embed 失败: {e}")
            return self.get_embed_target_hwnd()

    def _restore_main_window(self) -> None:
        """
        退出嵌入前恢复 OpusApp 主窗口可见（取消 ShowWindowsInTaskbar=False 的影响）。
        """
        if not self._word_app:
            return
        try:
            self._word_app.ShowWindowsInTaskbar = True
        except Exception as e:
            print(f"恢复 ShowWindowsInTaskbar 失败: {e}")

    def open_document(self, file_path: str, read_only: bool = False) -> DocumentInfo:
        """
        打开文档

        Args:
            file_path: 文档路径
            read_only: 是否只读打开

        Returns:
            DocumentInfo: 文档信息
        """
        if not self.is_connected():
            raise WordConnectionError("未连接到Word")

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")

        # 备份
        self._backup_current(file_path)

        # 打开文档
        try:
            self._document = self._word_app.Documents.Open(
                os.path.abspath(file_path),
                ReadOnly=read_only
            )
            self._current_file = file_path
            self._setup_document_window_for_embed()
            return self._get_document_info()

        except Exception as e:
            raise WordConnectionError(f"打开文档失败: {e}")

    def create_document(self) -> bool:
        """创建新文档"""
        if not self.is_connected():
            raise WordConnectionError("未连接到Word")

        try:
            self._document = self._word_app.Documents.Add()
            self._current_file = None
            self._setup_document_window_for_embed()
            return True

        except Exception as e:
            raise WordConnectionError(f"创建文档失败: {e}")

    def save_document(self, file_path: Optional[str] = None) -> bool:
        """
        保存文档

        Args:
            file_path: 保存路径，默认覆盖当前文档

        Returns:
            bool: 保存是否成功
        """
        if not self._document:
            return False

        try:
            if file_path:
                self._document.SaveAs(os.path.abspath(file_path))
                self._current_file = file_path
            else:
                self._document.Save()
            return True

        except Exception as e:
            print(f"保存文档失败: {e}")
            return False

    def close_document(self, save_changes: bool = True):
        """关闭文档"""
        if self._document:
            try:
                self._document.Close(SaveChanges=save_changes)
            except:
                pass
            self._document = None
            self._current_file = None

    def quit(self, save_changes: bool = False):
        """退出Word"""
        self.close_document(save_changes)
        self._restore_main_window()
        if self._word_app:
            try:
                self._word_app.Quit()
            except:
                pass
            self._word_app = None

    def _get_document_info(self) -> DocumentInfo:
        """获取文档信息"""
        if not self._document:
            return None

        try:
            return DocumentInfo(
                path=self._current_file or "",
                name=self._document.Name,
                fullname=self._document.FullName,
                page_count=self._document.ComputeStatistics(2),  # wdStatisticPages = 2
                word_count=self._document.ComputeStatistics(0),  # wdStatisticWords = 0
                paragraphs_count=self._document.ComputeStatistics(3),  # wdStatisticParagraphs = 3
                active=True
            )
        except:
            return DocumentInfo(
                path=self._current_file or "",
                name="",
                fullname="",
                active=True
            )

    def get_paragraphs(self) -> List[Dict[str, Any]]:
        """获取所有段落"""
        if not self._document:
            return []

        paragraphs = []
        for i, para in enumerate(self._document.Paragraphs):
            try:
                para_info = {
                    'index': i,
                    'text': para.Range.Text.strip() if para.Range.Text else "",
                    'alignment': self._get_alignment(para.Alignment),
                    'line_spacing': para.Format.LineSpacing,
                    'font_name': para.Range.Font.Name,
                    'font_size': para.Format.Size,
                    'bold': para.Range.Font.Bold == -1,
                    'italic': para.Range.Font.Italic == -1
                }
                paragraphs.append(para_info)
            except:
                pass

        return paragraphs

    def get_text(self) -> str:
        """获取文档全部文本"""
        if not self._document:
            return ""

        try:
            return self._document.Content.Text
        except:
            return ""

    def set_font(self, font_name: str, size: Optional[float] = None,
                 bold: Optional[bool] = None, italic: Optional[bool] = None,
                 start: int = 0, end: int = -1) -> bool:
        """
        设置字体

        Args:
            font_name: 字体名称
            size: 字号
            bold: 是否加粗
            italic: 是否斜体
            start: 起始位置
            end: 结束位置 (-1表示到末尾)

        Returns:
            bool: 设置是否成功
        """
        if not self._document:
            return False

        try:
            if end == -1:
                end = self._document.Content.End

            range_obj = self._document.Range(start, end)

            if font_name:
                range_obj.Font.Name = font_name

            if size is not None:
                range_obj.Font.Size = size

            if bold is not None:
                range_obj.Font.Bold = -1 if bold else 0

            if italic is not None:
                range_obj.Font.Italic = -1 if italic else 0

            return True

        except Exception as e:
            print(f"设置字体失败: {e}")
            return False

    def set_paragraph_alignment(self, alignment: str, paragraph_index: int = -1) -> bool:
        """
        设置段落对齐

        Args:
            alignment: 对齐方式 (left, center, right, justify)
            paragraph_index: 段落索引 (-1表示当前段落)

        Returns:
            bool: 设置是否成功
        """
        if not self._document:
            return False

        try:
            alignment_map = {
                'left': 0,      # wdAlignParagraphLeft
                'center': 1,    # wdAlignParagraphCenter
                'right': 2,     # wdAlignParagraphRight
                'justify': 3   # wdAlignParagraphJustify
            }

            if paragraph_index == -1:
                para = self._word_app.Selection.ParagraphFormat
            else:
                para = self._document.Paragraphs(paragraph_index).Format

            para.Alignment = alignment_map.get(alignment, 0)
            return True

        except Exception as e:
            print(f"设置对齐失败: {e}")
            return False

    def set_line_spacing(self, spacing: float, paragraph_index: int = -1) -> bool:
        """
        设置行距

        Args:
            spacing: 行距倍数 (1.0, 1.5, 2.0等)
            paragraph_index: 段落索引

        Returns:
            bool: 设置是否成功
        """
        if not self._document:
            return False

        try:
            if paragraph_index == -1:
                para = self._word_app.Selection.ParagraphFormat
            else:
                para = self._document.Paragraphs(paragraph_index).Format

            para.LineSpacingRule = 1  # wdLineSpaceMultiple
            para.LineSpacing = spacing * 12  # 转换为磅

            return True

        except Exception as e:
            print(f"设置行距失败: {e}")
            return False

    def set_page_setup(self, page_setup: PageSetup) -> bool:
        """
        设置页面设置

        Args:
            page_setup: 页面设置对象

        Returns:
            bool: 设置是否成功
        """
        if not self._document:
            return False

        try:
            ps = self._document.PageSetup

            # 纸张大小
            paper_sizes = {'A4': 7, 'Letter': 1}
            ps.PaperSize = paper_sizes.get(page_setup.paper_size, 7)

            # 方向
            ps.Orientation = 1 if page_setup.orientation == "portrait" else 0  # wdPortrait/wdLandscape

            # 边距 (转换为磅，1cm=28.35磅)
            cm_to_pt = 28.35
            ps.TopMargin = page_setup.top_margin * cm_to_pt
            ps.BottomMargin = page_setup.bottom_margin * cm_to_pt
            ps.LeftMargin = page_setup.left_margin * cm_to_pt
            ps.RightMargin = page_setup.right_margin * cm_to_pt
            ps.HeaderDistance = page_setup.header_distance * cm_to_pt
            ps.FooterDistance = page_setup.footer_distance * cm_to_pt

            return True

        except Exception as e:
            print(f"设置页面失败: {e}")
            return False

    def insert_text(self, text: str, at_position: int = -1) -> bool:
        """
        插入文本

        Args:
            text: 要插入的文本
            at_position: 插入位置 (-1表示当前光标处)

        Returns:
            bool: 插入是否成功
        """
        if not self._document:
            return False

        try:
            if at_position == -1:
                range_obj = self._word_app.Selection.Range
            else:
                range_obj = self._document.Range(at_position, at_position)

            range_obj.Text = text
            return True

        except Exception as e:
            print(f"插入文本失败: {e}")
            return False

    def apply_style(self, style_name: str, start: int = 0, end: int = -1) -> bool:
        """
        应用样式

        Args:
            style_name: 样式名称
            start: 起始位置
            end: 结束位置

        Returns:
            bool: 应用是否成功
        """
        if not self._document:
            return False

        try:
            if end == -1:
                end = self._document.Content.End

            range_obj = self._document.Range(start, end)
            range_obj.Style = style_name

            return True

        except Exception as e:
            print(f"应用样式失败: {e}")
            return False

    def get_styles(self) -> List[str]:
        """获取文档中的所有样式"""
        if not self._document:
            return []

        try:
            styles = []
            for style in self._document.Styles:
                try:
                    styles.append(style.NameLocal)
                except:
                    pass
            return styles
        except:
            return []

    def add_page_number(self, position: str = "bottom", format: str = "1") -> bool:
        """
        添加页码

        Args:
            position: 位置 (top/bottom)
            format: 页码格式

        Returns:
            bool: 添加是否成功
        """
        if not self._document:
            return False

        try:
            # 获取页脚或页眉
            if position == "top":
                section = self._document.Sections(1).Headers(1)
            else:
                section = self._document.Sections(1).Footers(1)

            # 添加页码
            page_num = section.PageNumbers.Add()
            page_num.NumberFormat = format

            return True

        except Exception as e:
            print(f"添加页码失败: {e}")
            return False

    def insert_table(self, rows: int, cols: int, at_position: int = -1) -> bool:
        """
        插入表格

        Args:
            rows: 行数
            cols: 列数
            at_position: 插入位置

        Returns:
            bool: 插入是否成功
        """
        if not self._document:
            return False

        try:
            if at_position == -1:
                range_obj = self._word_app.Selection.Range
            else:
                range_obj = self._document.Range(at_position, at_position)

            table = self._document.Tables.Add(range_obj, rows, cols)
            return True

        except Exception as e:
            print(f"插入表格失败: {e}")
            return False

    def undo(self) -> bool:
        """撤销操作"""
        if self._undo_manager.can_undo():
            return self._undo_manager.undo()
        return False

    def redo(self) -> bool:
        """重做操作"""
        if self._undo_manager.can_redo():
            return self._undo_manager.redo()
        return False

    def can_undo(self) -> bool:
        return self._undo_manager.can_undo()

    def can_redo(self) -> bool:
        return self._undo_manager.can_redo()

    def _backup_current(self, file_path: str):
        """备份当前文档"""
        if not os.path.exists(file_path):
            return

        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.basename(file_path)
            backup_name = f"{filename}.{timestamp}.bak"
            backup_path = self._backup_dir / backup_name
            shutil.copy2(file_path, backup_path)

            # 清理旧备份（保留最近10个）
            backups = sorted(self._backup_dir.glob(f"{filename}.*.bak"))
            for old_backup in backups[:-10]:
                old_backup.unlink()

        except Exception as e:
            print(f"备份失败: {e}")

    @staticmethod
    def _get_alignment(wd_align: int) -> str:
        """转换Word对齐方式到字符串"""
        alignments = {
            0: "left",
            1: "center",
            2: "right",
            3: "justify"
        }
        return alignments.get(wd_align, "left")

    # ====== 选区操作 ======

    def has_selection(self) -> bool:
        """检查是否有选中文本"""
        if not self._word_app:
            return False
        try:
            sel = self._word_app.Selection
            return sel.Start != sel.End
        except:
            return False

    def get_selection_text(self) -> str:
        """获取当前选中的文本"""
        if not self._word_app:
            return ""
        try:
            return self._word_app.Selection.Text.strip()
        except:
            return ""

    def get_selection_range(self) -> tuple:
        """获取当前选区的范围 (start, end)"""
        if not self._word_app:
            return (0, 0)
        try:
            sel = self._word_app.Selection
            return (sel.Start, sel.End)
        except:
            return (0, 0)

    def select_paragraph(self) -> bool:
        """选中当前段落"""
        if not self._word_app:
            return False
        try:
            self._word_app.Selection.Expand(4)  # wdParagraph = 4
            return True
        except:
            return False

    # ====== 字体操作 ======

    def set_font(self, font_name: Optional[str] = None, size: Optional[float] = None,
                  bold: Optional[bool] = None, italic: Optional[bool] = None,
                  underline: Optional[bool] = None) -> bool:
        """
        设置当前选区的字体属性

        Args:
            font_name: 字体名称
            size: 字号
            bold: 是否加粗
            italic: 是否斜体
            underline: 是否有下划线

        Returns:
            bool: 设置是否成功
        """
        if not self._word_app:
            return False

        try:
            sel = self._word_app.Selection
            font = sel.Font

            if font_name is not None:
                font.Name = font_name

            if size is not None:
                font.Size = float(size)

            if bold is not None:
                font.Bold = -1 if bold else 0

            if italic is not None:
                font.Italic = -1 if italic else 0

            if underline is not None:
                font.Underline = -1 if underline else 0  # -1=wdUnderlineSingle, 0=wdUnderlineNone

            return True

        except Exception as e:
            print(f"设置字体失败: {e}")
            return False

    def set_font_color(self, color: str) -> bool:
        """
        设置当前选区文字的颜色

        Args:
            color: 颜色值，可以是 "000000" 格式的十六进制，或颜色名称

        Returns:
            bool: 设置是否成功
        """
        if not self._word_app:
            return False

        try:
            # 颜色映射
            color_map = {
                "黑色": "000000", "黑色": "000000",
                "红色": "FF0000", "红色": "FF0000",
                "蓝色": "0000FF", "蓝色": "0000FF",
                "绿色": "00FF00", "绿色": "00FF00",
                "白色": "FFFFFF", "白色": "FFFFFF",
                "黄色": "FFFF00", "黄色": "FFFF00"
            }

            # 转换颜色名称为十六进制
            if color in color_map:
                color = color_map[color]

            # 移除 # 号
            color = color.lstrip("#")

            # 转换为 Word 颜色格式
            if len(color) == 6:
                r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
                wd_color = (b << 16) | (g << 8) | r

                self._word_app.Selection.Font.Color = wd_color
                return True

            return False

        except Exception as e:
            print(f"设置颜色失败: {e}")
            return False

    # ====== 段落操作 ======

    def set_indent(self, first_line: Optional[int] = None,
                   left_indent: Optional[int] = None,
                   right_indent: Optional[int] = None,
                   indent_type: str = "first_line") -> bool:
        """
        设置段落缩进

        Args:
            first_line: 首行缩进（磅），负值表示悬挂缩进
            left_indent: 左边缩进（磅）
            right_indent: 右边缩进（磅）
            indent_type: 缩进类型（first_line/left/right）

        Returns:
            bool: 设置是否成功
        """
        if not self._word_app:
            return False

        try:
            para_fmt = self._word_app.Selection.ParagraphFormat

            if first_line is not None and indent_type == "first_line":
                para_fmt.FirstLineIndent = float(first_line)

            if left_indent is not None and indent_type == "left":
                para_fmt.LeftIndent = float(left_indent)

            if right_indent is not None and indent_type == "right":
                para_fmt.RightIndent = float(right_indent)

            return True

        except Exception as e:
            print(f"设置缩进失败: {e}")
            return False

    def set_paragraph_spacing(self, before: Optional[float] = None,
                              after: Optional[float] = None) -> bool:
        """
        设置段落间距

        Args:
            before: 段前间距（磅）
            after: 段后间距（磅）

        Returns:
            bool: 设置是否成功
        """
        if not self._word_app:
            return False

        try:
            para_fmt = self._word_app.Selection.ParagraphFormat

            if before is not None:
                para_fmt.SpaceBefore = float(before)

            if after is not None:
                para_fmt.SpaceAfter = float(after)

            return True

        except Exception as e:
            print(f"设置段落间距失败: {e}")
            return False

    def set_line_spacing(self, spacing: float) -> bool:
        """
        设置行距

        Args:
            spacing: 行距倍数 (1.0=单倍, 1.5=1.5倍, 2.0=2倍)

        Returns:
            bool: 设置是否成功
        """
        if not self._word_app:
            return False

        try:
            para_fmt = self._word_app.Selection.ParagraphFormat
            para_fmt.LineSpacingRule = 1  # wdLineSpaceMultiple
            para_fmt.LineSpacing = spacing * 20.7  # 转换为磅（12磅字×1.725）

            return True

        except Exception as e:
            print(f"设置行距失败: {e}")
            return False

    def set_alignment(self, alignment: str) -> bool:
        """
        设置段落对齐（操作当前选中的段落）

        Args:
            alignment: 对齐方式 (left/center/right/justify)

        Returns:
            bool: 设置是否成功
        """
        if not self._word_app:
            return False

        try:
            alignment_map = {
                'left': 0,
                'center': 1,
                'right': 2,
                'justify': 3
            }

            para_fmt = self._word_app.Selection.ParagraphFormat
            para_fmt.Alignment = alignment_map.get(alignment, 0)

            return True

        except Exception as e:
            print(f"设置对齐失败: {e}")
            return False

    # ====== 页面操作 ======

    def set_page_margins(self, top: Optional[float] = None,
                        bottom: Optional[float] = None,
                        left: Optional[float] = None,
                        right: Optional[float] = None) -> bool:
        """
        设置页边距

        Args:
            top: 上边距（厘米）
            bottom: 下边距（厘米）
            left: 左边距（厘米）
            right: 右边距（厘米）

        Returns:
            bool: 设置是否成功
        """
        if not self._document:
            return False

        try:
            cm_to_pt = 28.35
            ps = self._document.PageSetup

            if top is not None:
                ps.TopMargin = top * cm_to_pt
            if bottom is not None:
                ps.BottomMargin = bottom * cm_to_pt
            if left is not None:
                ps.LeftMargin = left * cm_to_pt
            if right is not None:
                ps.RightMargin = right * cm_to_pt

            return True

        except Exception as e:
            print(f"设置页边距失败: {e}")
            return False

    def set_paper_size(self, paper_size: str) -> bool:
        """
        设置纸张大小

        Args:
            paper_size: 纸张大小 (A4/Letter/A3/A5)

        Returns:
            bool: 设置是否成功
        """
        if not self._document:
            return False

        try:
            paper_sizes = {
                'A4': 7,       # wdPaperA4
                'Letter': 1,   # wdPaperLetter
                'A3': 8,       # wdPaperA3
                'A5': 11       # wdPaperA5
            }

            self._document.PageSetup.PaperSize = paper_sizes.get(paper_size.upper(), 7)
            return True

        except Exception as e:
            print(f"设置纸张大小失败: {e}")
            return False

    def set_page_orientation(self, orientation: str) -> bool:
        """
        设置页面方向

        Args:
            orientation: 方向 (portrait/landscape)

        Returns:
            bool: 设置是否成功
        """
        if not self._document:
            return False

        try:
            # wdOrientationPortrait = 1, wdOrientationLandscape = 0
            self._document.PageSetup.Orientation = 1 if orientation == "portrait" else 0
            return True

        except Exception as e:
            print(f"设置页面方向失败: {e}")
            return False
