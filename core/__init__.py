# Core 模块 - 仅供 FastAPI 后端服务使用

from .word_connector import WordConnector
from .llm_service import LLMService
from .settings import settings
from .config import load_config, save_config

__all__ = [
    "WordConnector",
    "LLMService",
    "settings",
    "load_config",
    "save_config",
]
