"""
服务层模块
在进程启动时初始化 WordConnector 和 LLMService，
为 FastAPI 路由提供无状态的调用接口。
"""

import sys
import os

# 将 core/ 目录加入 import 路径
_core_path = os.path.join(os.path.dirname(__file__), "..", "..", "core")
sys.path.insert(0, os.path.abspath(_core_path))

from core.word_connector import WordConnector
from core.llm_service import LLMService
from core.settings import settings

# ── 全局单例 ──────────────────────────────────────────────────────

word_service = WordConnector()
llm_service = LLMService(config={
    "provider": settings.llm.provider,
    "api_key": settings.llm.api_key,
    "model": settings.llm.model,
    "base_url": settings.llm.base_url,
    "temperature": settings.llm.temperature,
    "max_tokens": settings.llm.max_tokens,
})

# 尝试自动连接 Word（可选，失败不阻止服务启动）
try:
    word_service.connect(visible=True)
except Exception as e:
    print(f"[OfficeHelper Service] Word 连接初始化失败: {e}")
