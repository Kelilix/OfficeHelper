"""
配置管理模块
提供应用配置的加载、保存和管理功能
"""

import os
import json
from pathlib import Path
from typing import Any, Optional
from dataclasses import dataclass, field, asdict
from dotenv import load_dotenv


# 加载环境变量
load_dotenv()


@dataclass
class LLMConfig:
    """大模型配置"""
    provider: str = "openai"  # openai / claude / ollama
    api_key: str = ""
    model: str = "gpt-4"
    base_url: str = "http://localhost:11434"
    temperature: float = 0.7
    max_tokens: int = 2000


@dataclass
class UIConfig:
    """UI配置"""
    theme: str = "light"  # light / dark
    window_width: int = 1200
    window_height: int = 800
    language: str = "zh-CN"


@dataclass
class WordConfig:
    """Word配置"""
    auto_open: bool = True
    backup_enabled: bool = True
    backup_dir: str = ""


@dataclass
class AppConfig:
    """应用总配置"""
    llm: LLMConfig = field(default_factory=LLMConfig)
    ui: UIConfig = field(default_factory=UIConfig)
    word: WordConfig = field(default_factory=WordConfig)
    debug: bool = False


class Settings:
    """配置管理类"""

    _instance: Optional['Settings'] = None
    _config: AppConfig = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        if self._config is None:
            self._load_config()

    def _get_config_path(self) -> Path:
        """获取配置文件路径"""
        config_dir = Path.home() / ".office_helper"
        config_dir.mkdir(exist_ok=True)
        return config_dir / "config.json"

    def _load_config(self):
        """加载配置"""
        config_path = self._get_config_path()

        # 默认配置
        self._config = AppConfig()

        # 从环境变量加载
        self._config.llm.provider = os.getenv("LLM_PROVIDER", self._config.llm.provider)
        self._config.llm.api_key = os.getenv("OPENAI_API_KEY", "")
        self._config.llm.model = os.getenv("LLM_MODEL", self._config.llm.model)
        self._config.llm.base_url = os.getenv("LLM_BASE_URL", self._config.llm.base_url)

        self._config.debug = os.getenv("DEBUG", "false").lower() == "true"

        # 如果有保存的配置，覆盖默认值
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    saved = json.load(f)
                self._apply_saved_config(saved)
            except Exception as e:
                print(f"加载配置文件失败: {e}")

    def _apply_saved_config(self, saved: dict):
        """应用保存的配置"""
        if 'llm' in saved:
            for key, value in saved['llm'].items():
                if hasattr(self._config.llm, key):
                    setattr(self._config.llm, key, value)

        if 'ui' in saved:
            for key, value in saved['ui'].items():
                if hasattr(self._config.ui, key):
                    setattr(self._config.ui, key, value)

        if 'word' in saved:
            for key, value in saved['word'].items():
                if hasattr(self._config.word, key):
                    setattr(self._config.word, key, value)

    def save(self):
        """保存配置"""
        config_path = self._get_config_path()

        config_dict = {
            'llm': asdict(self._config.llm),
            'ui': asdict(self._config.ui),
            'word': asdict(self._config.word),
        }

        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config_dict, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置文件失败: {e}")

    @property
    def llm(self) -> LLMConfig:
        return self._config.llm

    @property
    def ui(self) -> UIConfig:
        return self._config.ui

    @property
    def word(self) -> WordConfig:
        return self._config.word

    @property
    def is_debug(self) -> bool:
        return self._config.debug

    def get(self, key: str, default: Any = None) -> Any:
        """获取配置值"""
        keys = key.split('.')
        obj = self._config
        for k in keys:
            if hasattr(obj, k):
                obj = getattr(obj, k)
            else:
                return default
        return obj

    def set(self, key: str, value: Any):
        """设置配置值"""
        keys = key.split('.')
        obj = self._config
        for k in keys[:-1]:
            if hasattr(obj, k):
                obj = getattr(obj, k)
        if hasattr(obj, keys[-1]):
            setattr(obj, keys[-1], value)

    def reset(self):
        """重置为默认配置"""
        self._config = AppConfig()
        self.save()


# 全局配置实例
settings = Settings()
