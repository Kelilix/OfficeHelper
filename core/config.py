# 配置模块
from .settings import Settings, settings

__all__ = [
    "Settings",
    "settings",
    "load_config",
    "save_config",
]


def load_config(config_path: str | None = None) -> dict:
    """
    从 config.json 重新加载全部配置（替换内存）。

    Args:
        config_path: 可选，指定配置文件路径

    Returns:
        dict: 加载后的完整配置字典
    """
    return settings.load_config(config_path)


def save_config(config_path: str | None = None) -> bool:
    """
    将当前所有配置保存到 config.json。

    Args:
        config_path: 可选，指定配置文件路径

    Returns:
        bool: 保存是否成功
    """
    return settings.save_config(config_path)
