"""
OfficeHelper Python Service
FastAPI HTTP API for Word Web Add-in
"""
import os
import sys
import logging
import importlib
import importlib.util
import shutil
import tempfile
import ctypes
from pathlib import Path

# ── exe 模式调试：写入日志文件 ──────────────────────────────────────────────
def _write_debug_log(msg):
    log_path = Path(tempfile.gettempdir()) / "oh_debug.log"
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(msg + "\n")

_write_debug_log("=" * 60)
_write_debug_log("OfficeHelper exe 启动")
_write_debug_log(f"  sys._MEIPASS     = {repr(getattr(sys, '_MEIPASS', 'NOT_SET'))}")
_write_debug_log(f"  sys.frozen       = {getattr(sys, 'frozen', False)}")
_write_debug_log(f"  sys.executable  = {sys.executable}")
_write_debug_log(f"  os.getcwd()      = {os.getcwd()}")
_write_debug_log(f"  __file__         = {__file__}")
_write_debug_log(f"  __file__ parent  = {str(Path(__file__).parent)}")
_write_debug_log(f"  __file__ xparent  = {str(Path(__file__).parent.parent)}")

# ── 冻结模式：PyInstaller 已在启动时正确配置 sys.path[0] = _MEIPASS ────────
# 整个 pservice 包已被编译进 exe，直接 import 即可，无需额外处理


def _get_bundle_root():
    if getattr(sys, "_MEIPASS", None):
        return Path(sys._MEIPASS)
    return Path(__file__).parent.parent


def _extract_resources():
    bundle_root = _get_bundle_root()
    tmpdir = Path(tempfile.gettempdir()) / "OfficeHelperBackend"
    tmpdir.mkdir(exist_ok=True)

    for name in ("skills", "config.json"):
        src = bundle_root / name
        dst = tmpdir / name
        if not dst.exists():
            if src.exists() and src.is_dir():
                shutil.copytree(src, dst)
            elif src.exists() and src.is_file():
                shutil.copy2(src, dst)

    return tmpdir


_project_root = Path(__file__).parent.parent
if _project_root.exists():
    sys.path.insert(0, str(_project_root))

_is_frozen = getattr(sys, "_MEIPASS", None)
_extracted_root = _extract_resources()
if _is_frozen and _extracted_root.exists() and str(_extracted_root) not in sys.path:
    sys.path.insert(0, str(_extracted_root))


def _preload_skill_modules():
    for skill_name, module_name, filename, submodules in [
        ("word-text-operator", "scripts.word_text_operator", "word_text_operator",
         ["word_base", "word_range_navigation", "word_text_operations",
          "word_selection", "word_find_replace", "word_format", "word_bookmark"]),
        ("word-paragraph-operator", "scripts.word_paragraph_operator", "word_paragraph_operator", []),
    ]:
        bundle_root = _get_bundle_root()
        scripts_dir = bundle_root / "skills" / skill_name / "scripts"
        if not scripts_dir.exists():
            scripts_dir = _extracted_root / "skills" / skill_name / "scripts"
        if not scripts_dir.exists():
            continue
        if str(scripts_dir) not in sys.path:
            sys.path.insert(0, str(scripts_dir))
        if module_name not in sys.modules:
            _load_module_files(scripts_dir, module_name, filename, submodules)


def _load_module_files(scripts_dir, module_name, filename, submodules):
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


_preload_skill_modules()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
    filename=None,  # 调试模式下写入控制台
    force=True,
)

try:
    from pservice.api import app
    from pservice.api.routes import router as api_router
    from pservice.api.service import word_service, llm_service
    _write_debug_log("  [OK] imports succeeded")
except Exception as _e:
    _write_debug_log(f"  [ERROR] imports failed: {_e}")
    import traceback
    _write_debug_log(traceback.format_exc())
    raise

app.include_router(api_router)


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


def run():
    # console=False 时 sys.stderr 为 None，而 uvicorn 内部会调用
    # sys.stderr.isatty()，必须在 import uvicorn 之前先分配控制台
    if sys.stderr is None:
        STD_OUTPUT_HANDLE = -11
        STD_ERROR_HANDLE = -12
        kernel32 = ctypes.windll.kernel32
        kernel32.AllocConsole()
        kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
        kernel32.GetStdHandle(STD_ERROR_HANDLE)
        sys.stdout = sys.stderr = open("CONOUT$", "w", encoding="utf-8", buffering=1)

    import uvicorn
    import logging
    _write_debug_log(f"  [run] sys.path[0] = {sys.path[0]}")
    _write_debug_log(f"  [run] pservice in sys.modules? {'pservice' in sys.modules}")
    _write_debug_log(f"  [run] pservice.api in sys.modules? {'pservice.api' in sys.modules}")
    _write_debug_log(f"  [run] app id = {id(app)}")
    _write_debug_log("  [run] 准备启动 uvicorn...")

    # 清除所有 uvicorn logger 的 handler，防止 PyInstaller console=False 时
    # sys.stderr 为 None 导致 StreamHandler.emit() 抛出 AttributeError。
    for _name in ("uvicorn", "uvicorn.error", "uvicorn.access", "uvicorn.asgi"):
        _logger = logging.getLogger(_name)
        _logger.handlers.clear()
        _logger.propagate = False

    # 直接传 app 实例（已在模块级 import），跳过 uvicorn 内部的 import_from_string()。
    # PyInstaller 打包后 import_from_string("pservice.main:app") 可能找不到路径。
    # lifespan="off" 避免 uvicorn/lifespan/on.py 的 BaseException handler
    # 把异常当作"lifespan 不支持"静默吞掉。
    config = uvicorn.Config(
        app,                          # 直接传实例，跳过 import_from_string
        host="0.0.0.0",
        port=8765,
        reload=False,
        lifespan="off",                # 禁用 lifespan，避开 BaseException 吞异常问题
        log_level="info",
        access_log=False,
        log_config=None,
    )
    server = uvicorn.Server(config)
    try:
        server.run()
    except KeyboardInterrupt:
        _write_debug_log("  [uvicorn] 被 Ctrl+C 终止")
    except SystemExit as _se:
        import traceback
        _write_debug_log(f"  [uvicorn] SystemExit code={_se.code}")
        _write_debug_log(traceback.format_exc())
    except Exception as _e:
        import traceback
        _write_debug_log(f"  [FATAL] uvicorn 启动失败: {_e}")
        _write_debug_log(traceback.format_exc())
        traceback.print_exc()
    else:
        _write_debug_log("  [uvicorn] server.run() 正常退出")
    import time
    time.sleep(5)


if __name__ == "__main__":
    run()
