"""
截屏管理模块
提供屏幕截图、区域选择、图像处理等功能
"""

import base64
import io
from pathlib import Path
from typing import Optional, Tuple, List
from dataclasses import dataclass
from PIL import Image
import mss
import numpy as np


@dataclass
class ScreenRegion:
    """屏幕区域"""
    x: int
    y: int
    width: int
    height: int

    def to_dict(self):
        return {
            'x': self.x,
            'y': self.y,
            'width': self.width,
            'height': self.height
        }


@dataclass
class ScreenshotResult:
    """截屏结果"""
    success: bool
    image_data: Optional[str] = None  # base64
    width: Optional[int] = None
    height: Optional[int] = None
    region: Optional[dict] = None
    error: Optional[str] = None

    def to_dict(self):
        return {
            'success': self.success,
            'image_data': self.image_data,
            'width': self.width,
            'height': self.height,
            'region': self.region,
            'error': self.error
        }


class ScreenshotManager:
    """截屏管理器"""

    def __init__(self):
        self._screenshot_dir = Path.home() / ".office_helper" / "screenshots"
        self._screenshot_dir.mkdir(parents=True, exist_ok=True)
        self._last_screenshot: Optional[Image.Image] = None

    def capture_full_screen(self, monitor: int = 1) -> ScreenshotResult:
        """
        捕获整个屏幕

        Args:
            monitor: 显示器编号 (1-based)

        Returns:
            ScreenshotResult: 截屏结果
        """
        try:
            with mss.mss() as sct:
                # 获取指定显示器
                monitors = sct.monitors
                if monitor > len(monitors) - 1:
                    monitor = 1

                monitor_info = monitors[monitor]
                bbox = (
                    monitor_info["left"],
                    monitor_info["top"],
                    monitor_info["left"] + monitor_info["width"],
                    monitor_info["top"] + monitor_info["height"]
                )

                # 截屏
                img = sct.grab(bbox)

                # 转换为PIL Image
                self._last_screenshot = self._mss_to_pil(img)

                # 转换为base64
                image_data = self._pil_to_base64(self._last_screenshot)

                return ScreenshotResult(
                    success=True,
                    image_data=image_data,
                    width=self._last_screenshot.width,
                    height=self._last_screenshot.height,
                    region={
                        'x': bbox[0],
                        'y': bbox[1],
                        'width': bbox[2] - bbox[0],
                        'height': bbox[3] - bbox[1]
                    }
                )

        except Exception as e:
            return ScreenshotResult(
                success=False,
                error=f"截屏失败: {str(e)}"
            )

    def capture_region(self, region: ScreenRegion) -> ScreenshotResult:
        """
        捕获指定区域

        Args:
            region: 屏幕区域

        Returns:
            ScreenshotResult: 截屏结果
        """
        try:
            with mss.mss() as sct:
                bbox = (region.x, region.y,
                       region.x + region.width,
                       region.y + region.height)

                img = sct.grab(bbox)
                self._last_screenshot = self._mss_to_pil(img)
                image_data = self._pil_to_base64(self._last_screenshot)

                return ScreenshotResult(
                    success=True,
                    image_data=image_data,
                    width=self._last_screenshot.width,
                    height=self._last_screenshot.height,
                    region=region.to_dict()
                )

        except Exception as e:
            return ScreenshotResult(
                success=False,
                error=f"区域截屏失败: {str(e)}"
            )

    def capture_window(self, window_title: str) -> ScreenshotResult:
        """
        捕获指定窗口

        Args:
            window_title: 窗口标题（支持模糊匹配）

        Returns:
            ScreenshotResult: 截屏结果
        """
        try:
            import win32gui
            import win32ui
            import win32con
            import win32api

            # 查找窗口
            hwnd = win32gui.FindWindow(None, window_title)
            if not hwnd:
                # 尝试模糊匹配
                hwnd = self._find_window_by_title(window_title)

            if not hwnd:
                return ScreenshotResult(
                    success=False,
                    error=f"未找到窗口: {window_title}"
                )

            # 获取窗口位置
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top

            # 捕获窗口
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()

            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)

            result = saveDC.BitBlt(
                (0, 0), (width, height),
                mfcDC, (0, 0),
                win32con.SRCCOPY
            )

            # 转换为PIL Image
            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)
            self._last_screenshot = Image.frombuffer(
                'RGB',
                (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                bmpstr, 'raw', 'BGRX', 0, 1
            )

            # 释放资源
            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwndDC)

            # 转换为base64
            image_data = self._pil_to_base64(self._last_screenshot)

            return ScreenshotResult(
                success=True,
                image_data=image_data,
                width=width,
                height=height,
                region={'x': left, 'y': top, 'width': width, 'height': height}
            )

        except Exception as e:
            return ScreenshotResult(
                success=False,
                error=f"窗口截屏失败: {str(e)}"
            )

    def _find_window_by_title(self, title: str):
        """模糊查找窗口"""
        import win32gui

        result = []

        def enum_handler(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if title.lower() in window_title.lower():
                    result.append(hwnd)

        win32gui.EnumWindows(enum_handler, None)
        return result[0] if result else None

    def get_monitors(self) -> List[dict]:
        """
        获取所有显示器信息

        Returns:
            List[dict]: 显示器列表
        """
        try:
            with mss.mss() as sct:
                monitors = sct.monitors
                # 跳过第一个（所有显示器合并）
                return [
                    {
                        'id': i,
                        'width': m['width'],
                        'height': m['height'],
                        'x': m['left'],
                        'y': m['top']
                    }
                    for i, m in enumerate(monitors[1:], 1)
                ]
        except Exception:
            return []

    def save_screenshot(self, filepath: Optional[str] = None) -> Optional[str]:
        """
        保存截图到文件

        Args:
            filepath: 保存路径，默认保存到 screenshots 目录

        Returns:
            str: 保存的文件路径
        """
        if self._last_screenshot is None:
            return None

        if filepath is None:
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filepath = str(self._screenshot_dir / f"screenshot_{timestamp}.png")

        self._last_screenshot.save(filepath)
        return filepath

    def load_image_from_file(self, filepath: str) -> ScreenshotResult:
        """
        从文件加载图片

        Args:
            filepath: 图片文件路径

        Returns:
            ScreenshotResult: 图片结果
        """
        try:
            img = Image.open(filepath)
            self._last_screenshot = img
            image_data = self._pil_to_base64(img)

            return ScreenshotResult(
                success=True,
                image_data=image_data,
                width=img.width,
                height=img.height
            )
        except Exception as e:
            return ScreenshotResult(
                success=False,
                error=f"加载图片失败: {str(e)}"
            )

    def load_image_from_base64(self, image_data: str) -> ScreenshotResult:
        """
        从base64加载图片

        Args:
            image_data: base64编码的图片数据

        Returns:
            ScreenshotResult: 图片结果
        """
        try:
            img_bytes = base64.b64decode(image_data)
            img = Image.open(io.BytesIO(img_bytes))
            self._last_screenshot = img

            return ScreenshotResult(
                success=True,
                image_data=image_data,
                width=img.width,
                height=img.height
            )
        except Exception as e:
            return ScreenshotResult(
                success=False,
                error=f"解析图片失败: {str(e)}"
            )

    def get_last_screenshot(self) -> Optional[str]:
        """获取最后一次截图的base64数据"""
        if self._last_screenshot:
            return self._pil_to_base64(self._last_screenshot)
        return None

    @staticmethod
    def _mss_to_pil(img) -> Image.Image:
        """将mss截图转换为PIL Image"""
        return Image.frombytes("RGB", img.size, img.bgra, "raw", "BGRX")

    @staticmethod
    def _pil_to_base64(img: Image.Image) -> str:
        """将PIL Image转换为base64"""
        buffer = io.BytesIO()
        # 转换为RGB模式（如果需要）
        if img.mode != 'RGB':
            img = img.convert('RGB')
        img.save(buffer, format='PNG')
        return base64.b64encode(buffer.getvalue()).decode('utf-8')

    @staticmethod
    def resize_image(img: Image.Image, max_width: int = 1920, max_height: int = 1080) -> Image.Image:
        """调整图片大小"""
        if img.width <= max_width and img.height <= max_height:
            return img

        ratio = min(max_width / img.width, max_height / img.height)
        new_size = (int(img.width * ratio), int(img.height * ratio))
        return img.resize(new_size, Image.Resampling.LANCZOS)

    def capture_and_resize(self, max_width: int = 1920, max_height: int = 1080) -> ScreenshotResult:
        """截屏并自动调整大小"""
        result = self.capture_full_screen()
        if result.success and self._last_screenshot:
            self._last_screenshot = self.resize_image(
                self._last_screenshot, max_width, max_height
            )
            result.image_data = self._pil_to_base64(self._last_screenshot)
            result.width = self._last_screenshot.width
            result.height = self._last_screenshot.height
        return result
