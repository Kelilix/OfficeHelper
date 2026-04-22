# -*- coding: utf-8 -*-
"""
OfficeHelperBackend PyInstaller 打包配置
用法: pyinstaller OfficeHelperBackend.spec --clean
"""

import sys
import os
from pathlib import Path

block_cipher = None

# 解决 PyInstaller spec 中 __file__ 未定义的问题
try:
    _spec_file = __file__
except NameError:
    _spec_file = sys.argv[0]

PROJECT_ROOT = Path(_spec_file).parent.resolve()

# ── hidden imports ─────────────────────────────────────────────────────────
# win32com 核心（必须显式声明，PyInstaller 无法自动检测 COM 导入）
WIN32COM_IMPORTS = [
    # COM 核心
    "pythoncom",
    "pywintypes",
    # win32api 模块
    "win32api",
    "win32con",
    # win32com 包
    "win32com",
    "win32com.client",
    "win32com.server",
    "win32com.server.util",
    "win32com.server.policy",
    "win32com.util",
    # win32gui 模块
    "win32gui",
    "win32gui_struct",
    # win32 文件/进程/服务
    "win32clipboard",
    "win32console",
    "win32event",
    "win32file",
    "win32process",
    "win32service",
    "win32ts",
    "win32profile",
    "win32security",
    "win32trace",
    # comtypes（部分 win32com 高级功能依赖）
    "comtypes",
    "comtypes._otypes",
    "comtypes.client._code_cache",
    "comtypes.client._helpers",
    "comtypes.server",
]

# FastAPI / uvicorn 相关（必须显式声明）
FASTAPI_IMPORTS = [
    "uvicorn",
    "uvicorn.logging",
    "uvicorn.loops",
    "uvicorn.loops.auto",
    "uvicorn.loops.uvloop",
    "uvicorn.protocols",
    "uvicorn.protocols.http",
    "uvicorn.protocols.http.auto",
    "uvicorn.protocols.websockets",
    "uvicorn.protocols.websockets.auto",
    "uvicorn.protocols.websockets.ultra",
    "uvicorn.lifespan",
    "uvicorn.lifespan.on",
    "uvicorn.config",
    "starlette",
    "starlette.responses",
    "starlette.middleware",
    "starlette.middleware.cors",
    "starlette.middleware.gzip",
    "starlette.requests",
    "starlette.routing",
    "starlette.status",
    "fastapi",
    "pydantic",
    "pydantic.main",
    "pydantic.fields",
    "pydantic.validators",
    "pydantic.error_wrappers",
]

# 业务依赖
BIZ_IMPORTS = [
    "requests",
    "httpx",
    "openai",
    "anthropic",
    "psutil",
    "jinja2",
    "jinja2.ext",
    "markdown",
    "PIL",
    "PIL.Image",
]

ALL_HIDDEN = WIN32COM_IMPORTS + FASTAPI_IMPORTS + BIZ_IMPORTS

# ── 数据文件（skills 目录、config.json、core 目录）────────────────────────────
DATAS = [
    (str(PROJECT_ROOT / "config.json"), "."),
    (str(PROJECT_ROOT / "skills"), "skills"),
    (str(PROJECT_ROOT / "core"), "core"),
]

# ── 分析入口点 ────────────────────────────────────────────────────────────
a = Analysis(
    [str(PROJECT_ROOT / "pservice" / "main.py")],
    pathex=[str(PROJECT_ROOT)],
    binaries=[
        # 手动包含 Python 核心 DLL（PyInstaller 无法自动检测非 PATH 路径）
        (r"D:\soft\python3.14\python314.dll", "."),
        (r"D:\soft\python3.14\vcruntime140.dll", "."),
        (r"D:\soft\python3.14\vcruntime140_1.dll", "."),
        (r"D:\soft\python3.14\python3.dll", "."),
    ],
    datas=DATAS,
    hiddenimports=ALL_HIDDEN,
    hookspath=[],
    hooksconfig={},
    keys=[],
    exclude_binaries=False,
    ignore_config_dll_errors=True,  # 忽略 DLL 配置错误
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# ── 打包 Python 模块 ────────────────────────────────────────────────────────
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ── 生成 EXE ────────────────────────────────────────────────────────────────
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="OfficeHelperBackend",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,  # GUI 程序，不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# ── 收集非 Python 二进制文件（dll、pyd）──────────────────────────────────
# PyInstaller 默认把所有二进制文件放 _internal 子目录。
# 但前端 Tauri 只打包 exe 本身，不打包 _internal，导致 DLL 找不到。
# 因此在 COLLECT 之后，将 _internal 中的核心 DLL 复制到 dist 根目录。
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="OfficeHelperBackend",
)

import shutil
dist_root = Path(PROJECT_ROOT) / "dist" / "OfficeHelperBackend"
internal_dir = dist_root / "_internal"
for dll in ("python314.dll", "python3.dll", "vcruntime140.dll", "vcruntime140_1.dll"):
    src = internal_dir / dll
    dst = dist_root / dll
    if src.exists():
        shutil.copy2(src, dst)

