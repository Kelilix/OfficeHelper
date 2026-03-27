"""
配置管理模块
直接从 config.json 读写配置，不依赖 dataclass / 环境变量。
部署时只需修改 config.json 文件即可。

支持的 provider（对应 config.json 中 llm.provider）：
  - openai   : OpenAI 官方 API
  - anthropic : Anthropic Claude
  - ollama   : 本地 Ollama 服务
  - qwen     : 通义千问 DashScope API

qwen 配置示例：
  {
    "llm": {
      "provider": "qwen",
      "api_key": "你的 DashScope API Key",
      "model": "qwen-plus",
      "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
      "temperature": 0.7,
      "max_tokens": 2000
    }
  }
"""

import json
from pathlib import Path
from typing import Any, Optional


# ── 路径常量 ────────────────────────────────────────────────────────────

_PROJECT_ROOT = Path(__file__).parent.parent.resolve()
_CONFIG_DIR = _PROJECT_ROOT
_CONFIG_FILE = _PROJECT_ROOT / "config.json"
_FALLBACK_CONFIG_FILE = Path.home() / ".office_helper" / "config.json"

# config.json 不存在时写入的默认内容
_DEFAULT_CONFIG = {
    "llm": {
        "provider": "openai",
        "api_key": "",
        "model": "gpt-4",
        "base_url": "https://api.openai.com/v1",
        "temperature": 0.7,
        "max_tokens": 2000
    },
    "ui": {
        "theme": "light",
        "window_width": 1200,
        "window_height": 800,
        "language": "zh-CN"
    },
    "word": {
        "auto_open": True,
        "backup_enabled": True,
        "backup_dir": ""
    },
    "debug": False
}


class Settings:
    """
    配置管理器，直接持有从 config.json 解析出来的原始字典。

    初始化 / load_config / save_config 均围绕 config.json 进行，
    不读写环境变量，不依赖 dataclass 包装。

    用法示例：
        settings.get("llm.provider")       # 获取值，不存在返回 None
        settings.get("llm.provider", "openai")  # 提供默认值
        settings.set("llm.model", "gpt-4o") # 设置值（仅修改内存）
        settings.save_config()              # 持久化到 config.json
    """

    _instance: Optional['Settings'] = None
    _config: dict = {}

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        if not self._config:
            self._load_config()

    # ── 文件操作 ────────────────────────────────────────────────────────

    @staticmethod
    def _ensure_dir():
        _CONFIG_DIR.mkdir(parents=True, exist_ok=True)

    def _load_config(self):
        """
        加载配置，按以下优先级查找：
          1. 项目根目录 config.json（优先）
          2. ~/.office_helper/config.json（迁移兼容）
          3. 均不存在 → 写入默认配置到项目根目录
        """
        # 项目根目录已有 → 直接加载
        if _CONFIG_FILE.exists():
            self._ensure_dir()
            try:
                with open(_CONFIG_FILE, 'r', encoding='utf-8-sig') as f:
                    self._config = json.load(f)
                return
            except (json.JSONDecodeError, IOError) as e:
                print(f"[Settings] 项目根目录 config.json 读取失败: {e}，尝试备选路径")
                # 继续走下面的备选逻辑

        # 项目根目录没有 → 尝试用户目录（迁移兼容）
        if _FALLBACK_CONFIG_FILE.exists():
            try:
                with open(_FALLBACK_CONFIG_FILE, 'r', encoding='utf-8-sig') as f:
                    self._config = json.load(f)
                print(f"[Settings] 从 {_FALLBACK_CONFIG_FILE} 加载配置，并迁移到项目根目录")
                # 立即写回项目根目录，保持路径一致性
                self._ensure_dir()
                with open(_CONFIG_FILE, 'w', encoding='utf-8') as f:
                    json.dump(self._config, f, ensure_ascii=False, indent=2)
                return
            except (json.JSONDecodeError, IOError) as e:
                print(f"[Settings] 用户目录 config.json 读取失败: {e}，使用默认配置")

        # 全部没有 → 写入默认配置到项目根目录
        self._config = _DEFAULT_CONFIG.copy()
        self._deep_copy_dicts(self._config)
        self._ensure_dir()
        try:
            with open(_CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, ensure_ascii=False, indent=2)
            print(f"[Settings] 未找到配置文件，已在 {_CONFIG_FILE} 生成默认配置")
        except IOError as e:
            print(f"[Settings] 写入默认配置失败: {e}")

    @staticmethod
    def _deep_copy_dicts(d: dict) -> dict:
        """递归深拷贝（将嵌套 dict 复制为独立对象）。"""
        result = {}
        for k, v in d.items():
            if isinstance(v, dict):
                result[k] = Settings._deep_copy_dicts(v)
            else:
                result[k] = v
        return result

    def load_config(self, config_path: Optional[str] = None) -> dict:
        """
        从指定路径重新加载配置，替换内存中的全部配置。

        Args:
            config_path: 配置文件路径，默认为项目根目录的 config.json

        Returns:
            加载后的完整配置字典
        """
        path = Path(config_path) if config_path else _CONFIG_FILE

        if not path.exists():
            print(f"[Settings] 配置文件不存在: {path}")
            return self._config

        try:
            with open(path, 'r', encoding='utf-8-sig') as f:
                self._config = json.load(f)
            print(f"[Settings] 配置已从 {path} 加载")
        except (json.JSONDecodeError, IOError) as e:
            print(f"[Settings] 配置文件读取失败: {e}")

        return self._config

    def save_config(self, config_path: Optional[str] = None) -> bool:
        """
        将当前配置保存到 config.json。

        Args:
            config_path: 配置文件路径，默认为项目根目录的 config.json

        Returns:
            bool: 保存是否成功
        """
        path = Path(config_path) if config_path else _CONFIG_FILE
        self._ensure_dir()

        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, ensure_ascii=False, indent=2)
            print(f"[Settings] 配置已保存到 {path}")
            return True
        except IOError as e:
            print(f"[Settings] 配置保存失败: {e}")
            return False

    def reset(self):
        """重置为默认配置并写入 config.json。"""
        self._config = self._deep_copy_dicts(_DEFAULT_CONFIG)
        self.save_config()

    # ── 配置读写 ───────────────────────────────────────────────────────

    def get(self, key: str, default: Any = None) -> Any:
        """
        按点分路径读取配置值。

        Args:
            key: 路径，如 "llm.provider"、"word.backup_enabled"
            default: 键不存在时返回的默认值

        Returns:
            配置值，不存在返回 default
        """
        parts = key.split('.')
        value = self._config
        for part in parts:
            if isinstance(value, dict) and part in value:
                value = value[part]
            else:
                return default
        return value

    def set(self, key: str, value: Any):
        """
        按点分路径设置配置值（仅修改内存，调用 save_config 才会写入文件）。

        Args:
            key: 路径，如 "llm.model"
            value: 要设置的值
        """
        parts = key.split('.')
        target = self._config
        for part in parts[:-1]:
            if part not in target:
                target[part] = {}
            target = target[part]
        target[parts[-1]] = value

    # ── 便捷属性 ───────────────────────────────────────────────────────

    @property
    def llm(self) -> dict:
        """返回 llm 配置子字典。"""
        return self._config.get("llm", {})

    @property
    def ui(self) -> dict:
        """返回 ui 配置子字典。"""
        return self._config.get("ui", {})

    @property
    def word(self) -> dict:
        """返回 word 配置子字典。"""
        return self._config.get("word", {})


# ── 全局单例 ─────────────────────────────────────────────────────────────

settings = Settings()
