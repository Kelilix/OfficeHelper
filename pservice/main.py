"""
OfficeHelper Python 服务
FastAPI HTTP API，供 Word Web 加载项调用
"""

import sys
import os
import logging
import importlib
import importlib.util
from pathlib import Path

# ── 统一设置 sys.path（在所有模块 import 之前）───────────────────────────────
_project_root = Path(__file__).parent.parent
if _project_root.exists():
    sys.path.insert(0, str(_project_root))

# 提前加载 skill/scripts/ 下的模块，用 importlib.util 绕过包搜索冲突
# 加载后注册到 sys.modules，使后续代码 "from scripts.xxx import Class" 正常工作。
# 关键：必须先加载所有子模块（word_base 等），再加载主模块，
# 因为主模块用相对导入 "from .word_base import ..."，需要子模块已在 sys.modules 中。


def _preload_skill_modules():
    """
    用 importlib.util 显式加载 skills/word-*/scripts/ 下的模块文件，
    并注册到 sys.modules，完全绕开 Python 的包搜索路径冲突。

    策略：word_text_operator.py 内部有相对导入（from .word_base import ...），
    直接用 importlib.util 加载会失败。正确做法是先加载所有子模块，
    再加载主模块，确保相对导入的依赖已在 sys.modules 中。
    """

    for skill_name in ("word-text-operator", "word-paragraph-operator"):
        scripts_dir = _project_root / "skills" / skill_name / "scripts"
        if not scripts_dir.exists():
            continue

        # 清除 Python 的路径缓存（防止它记住 win32/scripts 冲突）
        if str(scripts_dir) in sys.path_importer_cache:
            del sys.path_importer_cache[str(scripts_dir)]
        if str(scripts_dir.parent) in sys.path_importer_cache:
            del sys.path_importer_cache[str(scripts_dir.parent)]

        if str(scripts_dir) not in sys.path:
            sys.path.insert(0, str(scripts_dir))

        # 尝试让 Python 正常导入 scripts 包
        # 注意：可能被 win32/scripts 抢占，所以有兜底逻辑
        if "scripts" not in sys.modules:
            importlib.import_module("scripts")

    # 兜底：确保两个主模块都已注册
    if "scripts.word_text_operator" not in sys.modules:
        _load_module_files(
            _project_root / "skills" / "word-text-operator" / "scripts",
            "scripts.word_text_operator", "word_text_operator",
            ["word_base", "word_range_navigation", "word_text_operations",
             "word_selection", "word_find_replace", "word_format", "word_bookmark"]
        )
    if "scripts.word_paragraph_operator" not in sys.modules:
        _load_module_files(
            _project_root / "skills" / "word-paragraph-operator" / "scripts",
            "scripts.word_paragraph_operator", "word_paragraph_operator",
            []
        )


def _load_module_files(scripts_dir: Path, module_name: str, filename: str,
                       submodules: list):
    """逐文件加载所有模块（解决相对导入问题）。"""
    # 1. 先加载所有子模块（无相对导入的依赖模块）
    for sub in submodules:
        sub_file = scripts_dir / f"{sub}.py"
        if not sub_file.exists():
            continue
        sub_name = f"scripts.{sub}"
        if sub_name in sys.modules:
            continue
        spec = importlib.util.spec_from_file_location(sub_name, sub_file)
        if spec is None:
            continue
        mod = importlib.util.module_from_spec(spec)
        sys.modules[sub_name] = mod
        try:
            spec.loader.exec_module(mod)
        except Exception:
            pass

    # 2. 加载主模块
    main_file = scripts_dir / f"{filename}.py"
    if not main_file.exists():
        return
    if module_name in sys.modules:
        return
    spec = importlib.util.spec_from_file_location(module_name, main_file)
    if spec is None:
        return
    mod = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = mod
    try:
        spec.loader.exec_module(mod)
    except Exception:
        pass


# 立即执行预加载
_preload_skill_modules()

# ── 原有初始化 ──────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)

from api import app
from api.routes import router as api_router
from api.service import word_service, llm_service

# 注册路由
app.include_router(api_router)

# 健康检查
@app.get("/")
def root():
    return {
        "name": "OfficeHelper Python Service",
        "version": "1.0.0",
        "word_connected": word_service.is_connected(),
    }

@app.get("/health")
def health():
    return {"status": "ok", "word": word_service.is_connected()}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "main:app",
        host="127.0.0.1",
        port=8765,
        reload=False,
        log_level="info",
    )
