"""
服务层模块
在进程启动时初始化 WordConnector 和 LLMService，
为 FastAPI 路由提供无状态的调用接口。
配置全部来源于 config.json，无需传入参数。
"""

import sys
import os

# 将 core/ 目录加入 import 路径
_core_path = os.path.join(os.path.dirname(__file__), "..", "..", "core")
sys.path.insert(0, os.path.abspath(_core_path))

from core.word_connector import WordConnector
from core.llm_service import LLMService

# ── 全局单例 ─────────────────────────────────────────────────────────────

word_service = WordConnector()
llm_service = LLMService()   # 内部直接读取 config.json 中的 llm 配置

# 尝试自动连接 Word（可选，失败不阻止服务启动）
try:
    word_service.connect(visible=True)
except Exception as e:
    print(f"[OfficeHelper Service] Word 连接初始化失败: {e}")
