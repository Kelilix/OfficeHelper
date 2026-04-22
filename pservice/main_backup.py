"""
OfficeHelper Python Service
FastAPI HTTP API for Word Web Add-in
"""

import sys
import os
import logging
import importlib
import importlib.util
import shutil
import tempfile
from pathlib import Path


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

_extracted_root = _extract_resources()
if _extracted_root.exists():
    sys.path.insert(0, str(_extracted_root))


def _preload_skill_modules():
    for skill_name, module_name, filename, submodules in [
        ("word-text-operator", "scripts.word_text_operator", "word_text_operator",
         ["word_base", "word_range_navigation", "word_text_operations",
          "word_selection", "word_find_replace", "word_format", "word_bookmark"]),
        ("word-paragraph-operator", "scripts.word_paragraph_operator", "word_paragraph_operator", []),
    ]:
        scripts_dir = _extracted_root / "skills" / skill_name / "scripts"
        if not scripts_dir.exists():
            scripts_dir = _get_bundle_root() / "skills" / skill_name / "scripts"
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
)

from api import app
from api.routes import router as api_router
from api.service import word_service, llm_service

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


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=8765,
        reload=False,
        log_level="info",
    )
