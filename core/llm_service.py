"""
大模型服务模块
提供多提供商的大模型调用支持
"""

import json
import base64
import io
import logging
from typing import Optional, List, Dict, Any
from dataclasses import dataclass, field
from enum import Enum
from abc import ABC, abstractmethod

from core.settings import settings

logger = logging.getLogger(__name__)


class LLMProvider(Enum):
    """大模型提供商"""
    OPENAI = "openai"
    ANTHROPIC = "anthropic"
    OLLAMA = "ollama"
    QWEN = "qwen"


@dataclass
class ChatMessage:
    """聊天消息"""
    role: str  # system/user/assistant
    content: str
    image: Optional[str] = None  # base64图片

    def to_dict(self):
        return {
            'role': self.role,
            'content': self.content,
            'image': self.image
        }


@dataclass
class AnalysisResult:
    """分析结果"""
    success: bool
    detected_issues: List[str] = field(default_factory=list)
    suggested_actions: List[Dict[str, Any]] = field(default_factory=list)
    execution_plan: Dict[str, Any] = field(default_factory=dict)
    response_text: str = ""
    error: Optional[str] = None


@dataclass
class ModelInfo:
    """模型信息"""
    id: str
    name: str
    provider: str
    max_tokens: int = 4096
    vision: bool = False


class BaseLLMProvider(ABC):
    """大模型提供商基类"""

    @abstractmethod
    def chat(self, messages: List[ChatMessage], **kwargs) -> str:
        """发送聊天请求"""
        pass

    @abstractmethod
    def analyze_image(self, image_data: str, prompt: str, **kwargs) -> str:
        """分析图片"""
        pass

    @abstractmethod
    def list_models(self) -> List[ModelInfo]:
        """列出可用模型"""
        pass


class OpenAIProvider(BaseLLMProvider):
    """OpenAI提供商"""

    def __init__(self, api_key: str, base_url: Optional[str] = None, model: str = "gpt-4"):
        self.api_key = api_key
        self.base_url = base_url or "https://api.openai.com/v1"
        self.model = model

    def chat(self, messages: List[ChatMessage], **kwargs) -> str:
        """发送聊天请求"""
        from openai import OpenAI

        client = OpenAI(api_key=self.api_key, base_url=self.base_url)

        # 转换消息格式
        chat_messages = []
        for msg in messages:
            if msg.image:
                # 视觉模型需要特殊格式
                content = [
                    {"type": "text", "text": msg.content},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/png;base64,{msg.image}"}
                    }
                ]
            else:
                content = msg.content
            chat_messages.append({"role": msg.role, "content": content})

        response = client.chat.completions.create(
            model=self.model,
            messages=chat_messages,
            temperature=kwargs.get('temperature', 0.7),
            max_tokens=kwargs.get('max_tokens', 2000),
            stream=kwargs.get('stream', False)
        )

        if kwargs.get('stream'):
            return response
        return response.choices[0].message.content

    def analyze_image(self, image_data: str, prompt: str, **kwargs) -> str:
        """分析图片"""
        # 使用GPT-4 Vision
        messages = [
            ChatMessage(role="user", content=prompt, image=image_data)
        ]
        return self.chat(messages, **kwargs)

    def list_models(self) -> List[ModelInfo]:
        """列出可用模型"""
        return [
            ModelInfo("gpt-4o", "GPT-4o", "openai", 128000, True),
            ModelInfo("gpt-4-turbo", "GPT-4 Turbo", "openai", 128000, True),
            ModelInfo("gpt-4", "GPT-4", "openai", 8192, False),
            ModelInfo("gpt-3.5-turbo", "GPT-3.5 Turbo", "openai", 16385, False),
        ]


class AnthropicProvider(BaseLLMProvider):
    """Anthropic Claude提供商"""

    def __init__(self, api_key: str, model: str = "claude-3-opus-20240229"):
        self.api_key = api_key
        self.model = model

    def chat(self, messages: List[ChatMessage], **kwargs) -> str:
        """发送聊天请求"""
        import anthropic

        client = anthropic.Anthropic(api_key=self.api_key)

        # 转换消息格式
        chat_messages = []
        for msg in messages[1:]:  # 跳过system消息
            if msg.image:
                content = [
                    {"type": "text", "text": msg.content},
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": msg.image
                        }
                    }
                ]
            else:
                content = msg.content
            chat_messages.append({"role": msg.role, "content": content})

        system = messages[0].content if messages and messages[0].role == "system" else ""

        response = client.messages.create(
            model=self.model,
            system=system,
            messages=chat_messages,
            temperature=kwargs.get('temperature', 0.7),
            max_tokens=kwargs.get('max_tokens', 2000),
            stream=kwargs.get('stream', False)
        )

        if kwargs.get('stream'):
            return response
        return response.content[0].text

    def analyze_image(self, image_data: str, prompt: str, **kwargs) -> str:
        """分析图片"""
        messages = [
            ChatMessage(role="user", content=prompt, image=image_data)
        ]
        return self.chat(messages, **kwargs)

    def list_models(self) -> List[ModelInfo]:
        """列出可用模型"""
        return [
            ModelInfo("claude-3-5-sonnet-20241022", "Claude 3.5 Sonnet", "anthropic", 200000, True),
            ModelInfo("claude-3-opus-20240229", "Claude 3 Opus", "anthropic", 200000, True),
            ModelInfo("claude-3-sonnet-20240229", "Claude 3 Sonnet", "anthropic", 200000, True),
            ModelInfo("claude-3-haiku-20240307", "Claude 3 Haiku", "anthropic", 200000, True),
        ]


class OllamaProvider(BaseLLMProvider):
    """Ollama本地模型提供商"""

    def __init__(self, base_url: str = "http://localhost:11434", model: str = "llama2"):
        self.base_url = base_url
        self.model = model

    def chat(self, messages: List[ChatMessage], **kwargs) -> str:
        """发送聊天请求"""
        import requests

        url = f"{self.base_url}/api/chat"

        # 转换消息格式
        chat_messages = []
        for msg in messages:
            msg_dict = {"role": msg.role, "content": msg.content}
            chat_messages.append(msg_dict)

        payload = {
            "model": self.model,
            "messages": chat_messages,
            "stream": kwargs.get('stream', False),
            "options": {
                "temperature": kwargs.get('temperature', 0.7),
                "num_predict": kwargs.get('max_tokens', 2000)
            }
        }

        response = requests.post(url, json=payload, stream=kwargs.get('stream', False))

        if kwargs.get('stream'):
            return response

        result = response.json()
        return result.get('message', {}).get('content', '')

    def analyze_image(self, image_data: str, prompt: str, **kwargs) -> str:
        """分析图片 - Ollama视觉模型"""
        # 检查是否支持视觉
        # 这里简化处理，实际需要检查模型是否支持视觉
        messages = [
            ChatMessage(role="user", content=f"{prompt}\n\n[图片数据: {image_data[:100]}...]", image=image_data)
        ]
        return self.chat(messages, **kwargs)

    def list_models(self) -> List[ModelInfo]:
        """列出可用模型"""
        import requests

        try:
            response = requests.get(f"{self.base_url}/api/tags")
            if response.status_code == 200:
                data = response.json()
                models = []
                for m in data.get('models', []):
                    models.append(ModelInfo(
                        id=m['name'],
                        name=m['name'],
                        provider="ollama",
                        vision=m.get('supports_vision', False)
                    ))
                return models
        except:
            pass

        return [
            ModelInfo("llama2", "Llama 2", "ollama", 4096, False),
            ModelInfo("mistral", "Mistral", "ollama", 8192, False),
        ]


def _normalize_openai_base_url(url: str) -> str:
    """OpenAI SDK 要求 base_url 以 /v1 结尾；若已含 /v1 则不再追加（避免 .../v1/v1 导致 404）。"""
    u = (url or "").rstrip("/")
    if u.endswith("/v1"):
        return u
    return u + "/v1"


class QwenProvider(BaseLLMProvider):
    """
    通义千问（DashScope）提供商。

    DashScope API 支持 OpenAI 兼容格式，只需设置 base_url 即可。
    模型名示例：qwen-plus, qwen-turbo, qwen-max, qwen-vl-plus 等。
    """

    def __init__(self, api_key: str, base_url: str, model: str = "qwen-plus"):
        self.api_key = api_key
        self.base_url = _normalize_openai_base_url(base_url)
        self.model = model

    def chat(self, messages: List[ChatMessage], **kwargs) -> str:
        """发送聊天请求"""
        from openai import OpenAI

        client = OpenAI(
            api_key=self.api_key,
            base_url=self.base_url,
            # DashScope 不支持通过 SDK 传 temperature/max_tokens，用 extra_body
            timeout=120.0,
        )

        chat_messages = []
        for msg in messages:
            if msg.image:
                content = [
                    {"type": "text", "text": msg.content},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/png;base64,{msg.image}"}
                    }
                ]
            else:
                content = msg.content
            chat_messages.append({"role": msg.role, "content": content})

        extra_body = {
            k: v
            for k, v in {
                "temperature": kwargs.get("temperature", 0.7),
                "max_tokens": kwargs.get("max_tokens", 2000),
            }.items()
            if v is not None
        }

        response = client.chat.completions.create(
            model=self.model,
            messages=chat_messages,
            stream=kwargs.get("stream", False),
            extra_body=extra_body if extra_body else None,
        )

        if kwargs.get("stream"):
            return response
        return response.choices[0].message.content

    def analyze_image(self, image_data: str, prompt: str, **kwargs) -> str:
        """分析图片（需要视觉模型如 qwen-vl-plus）"""
        messages = [
            ChatMessage(role="user", content=prompt, image=image_data)
        ]
        return self.chat(messages, **kwargs)

    def list_models(self) -> List[ModelInfo]:
        """列出常用 Qwen 模型"""
        return [
            ModelInfo("qwen-plus", "Qwen Plus", "qwen", 32768, False),
            ModelInfo("qwen-turbo", "Qwen Turbo", "qwen", 8192, False),
            ModelInfo("qwen-max", "Qwen Max", "qwen", 8192, False),
            ModelInfo("qwen-vl-plus", "Qwen VL Plus", "qwen", 8192, True),
            ModelInfo("qwen-vl-max", "Qwen VL Max", "qwen", 8192, True),
        ]


class LLMService:
    """大模型服务，配置全部来自 config.json"""

    def __init__(self):
        self._provider: Optional[BaseLLMProvider] = None
        self._current_provider_type = None
        self._chat_history: List[ChatMessage] = []

        self._system_prompt = """你是一个专业的Word文档格式调整助手。
你的任务是帮助用户分析和修复Word文档的格式问题。

你可以执行以下操作：
1. 分析文档截图，识别格式问题
2. 提供修复建议和执行步骤
3. 回答关于Word格式调整的问题

请用中文回复，并给出具体的操作建议。"""

        self._init_provider()

    # ── LLM 请求 / 响应日志 ─────────────────────────────────────────────

    def _log_request(self, messages: List[ChatMessage], tag: str):
        """
        打印发送给大模型的完整请求内容。
        """
        model = getattr(self._provider, "model", "?")
        base_url = getattr(self._provider, "base_url", "?")
        provider = self._current_provider_type or "?"
        sep = "─" * 64

        logger.info("\n%s\n[LLM][%s] Provider=%s  Model=%s  BaseURL=%s\n%s",
                    sep, tag, provider, model, base_url, sep)

        for i, msg in enumerate(messages):
            role = msg.role.upper()
            content = msg.content
            has_image = " [含图片]" if msg.image else ""
            # 始终截断避免日志过长，但保留足够信息用于分析
            if len(content) > 800:
                content = content[:800] + f"\n... [共 {len(content)} 字符，已截断]"
            logger.info("  [%d] %s:%s%s", i, role, content, has_image)

        logger.info(sep)

    def _log_response(self, response: str, tag: str):
        """
        打印大模型返回的完整响应内容。
        """
        content = response
        if len(content) > 1500:
            content = content[:1500] + f"\n... [共 {len(content)} 字符，已截断]"
        logger.info("[LLM][%s] 响应:\n%s", tag, content)

    def _llm_kwargs(self, stream: bool) -> dict:
        cfg = settings.llm
        return {
            "stream": stream,
            "temperature": cfg.get("temperature", 0.7),
            "max_tokens": cfg.get("max_tokens", 2000),
        }

    def _call_provider_chat(self, messages: List[ChatMessage], stream: bool, tag: str) -> str:
        """统一调用 Provider，并记录 INFO / ERROR 日志（异常会原样抛出）。"""
        model = getattr(self._provider, "model", "?")
        base_url = getattr(self._provider, "base_url", "?")
        provider = self._current_provider_type or "?"
        logger.info(
            "[LLM] 请求开始 | tag=%s provider=%s model=%s base_url=%s",
            tag,
            provider,
            model,
            base_url,
        )
        try:
            return self._provider.chat(messages, **self._llm_kwargs(stream))
        except Exception as e:
            err_body = ""
            try:
                if hasattr(e, "body") and e.body:
                    err_body = str(e.body)[:500]
                elif hasattr(e, "response") and getattr(e.response, "text", None):
                    err_body = (e.response.text or "")[:500]
            except Exception:
                pass
            logger.exception(
                "[LLM] 调用失败 | tag=%s provider=%s model=%s base_url=%s extra=%s err=%s",
                tag,
                provider,
                model,
                base_url,
                err_body or "(无响应体)",
                e,
            )
            raise

    def _init_provider(self):
        """从 config.json 读取 llm 配置并初始化对应 Provider"""
        cfg = settings.llm  # 直接读 config["llm"] 字典

        provider_type = cfg.get("provider", "openai").lower()
        api_key = cfg.get("api_key", "")
        model = cfg.get("model", "gpt-4")
        base_url = cfg.get("base_url", "")

        if provider_type == "openai":
            self._provider = OpenAIProvider(api_key, base_url or None, model)
        elif provider_type == "anthropic":
            self._provider = AnthropicProvider(api_key, model)
        elif provider_type == "ollama":
            self._provider = OllamaProvider(base_url or "http://localhost:11434", model)
        elif provider_type == "qwen":
            self._provider = QwenProvider(api_key, base_url or "https://dashscope.aliyuncs.com/compatible-mode/v1", model)
        else:
            self._provider = OpenAIProvider(api_key, base_url or None, model)

        self._current_provider_type = provider_type

    def set_provider(self, provider_type: str):
        """切换 Provider 类型（配置从 config.json 重新读取）"""
        cfg = settings.llm
        api_key = cfg.get("api_key", "")
        model = cfg.get("model", "gpt-4")
        base_url = cfg.get("base_url", "")

        if provider_type == "openai":
            self._provider = OpenAIProvider(api_key, base_url or None, model)
        elif provider_type == "anthropic":
            self._provider = AnthropicProvider(api_key, model)
        elif provider_type == "ollama":
            self._provider = OllamaProvider(base_url or "http://localhost:11434", model)
        elif provider_type == "qwen":
            self._provider = QwenProvider(api_key, base_url or "https://dashscope.aliyuncs.com/compatible-mode/v1", model)
        else:
            self._provider = OpenAIProvider(api_key, base_url or None, model)

        self._current_provider_type = provider_type

    def chat(self, message: str, stream: bool = False) -> str:
        """
        发送聊天消息

        Args:
            message: 用户消息
            stream: 是否流式输出

        Returns:
            str: 助手回复
        """
        self._chat_history.append(ChatMessage(role="user", content=message))
        if len(self._chat_history) == 1:
            self._chat_history.insert(0, ChatMessage(role="system", content=self._system_prompt))
        self._log_request(self._chat_history, tag="chat")
        response = self._call_provider_chat(self._chat_history, stream, tag="chat")
        if not stream:
            self._chat_history.append(ChatMessage(role="assistant", content=response))
            self._log_response(response, tag="chat")
        return response

    def chat_with_context(self, user_message: str, system_context: str, stream: bool = False) -> str:
        """
        带完整上下文的聊天（供 API 路由使用）。

        Args:
            user_message: 用户的实际消息
            system_context: 完整的系统上下文（包含选区、技能描述等）
            stream: 是否流式输出

        Returns:
            str: 助手回复
        """
        self._chat_history.clear()
        self._chat_history.append(ChatMessage(role="system", content=system_context))
        self._chat_history.append(ChatMessage(role="user", content=user_message))
        self._log_request(self._chat_history, tag="chat_with_context")
        response = self._call_provider_chat(self._chat_history, stream, tag="chat_with_context")
        if not stream:
            self._chat_history.append(ChatMessage(role="assistant", content=response))
            self._log_response(response, tag="chat_with_context")
        return response

    def analyze_image(self, image_data: str, user_request: str = "") -> AnalysisResult:
        """
        分析截图

        Args:
            image_data: base64编码的图片
            user_request: 用户请求

        Returns:
            AnalysisResult: 分析结果
        """
        # 构建分析提示
        prompt = self._build_analysis_prompt(user_request)
        debug_msgs = [ChatMessage(role="user", content=prompt, image=image_data)]
        self._log_request(debug_msgs, tag="analyze_image")

        try:
            response = self._provider.analyze_image(
                image_data, prompt, **self._llm_kwargs(stream=False)
            )
            self._log_response(response, tag="analyze_image")
            return self._parse_analysis_response(response)
        except Exception as e:
            logger.exception(
                "[LLM] analyze_image 失败 | provider=%s model=%s base_url=%s err=%s",
                self._current_provider_type,
                getattr(self._provider, "model", "?"),
                getattr(self._provider, "base_url", "?"),
                e,
            )
            return AnalysisResult(
                success=False,
                error=str(e)
            )

    def _build_analysis_prompt(self, user_request: str) -> str:
        """构建分析提示词"""
        base_prompt = """请分析这张Word文档截图，指出格式问题并给出修复建议。

请按以下JSON格式返回分析结果：
```json
{
  "detected_issues": ["问题1", "问题2"],
  "suggested_actions": [
    {
      "skill": "set_font",
      "params": {"font_name": "微软雅黑", "size": 12},
      "description": "设置正文字体"
    }
  ],
  "execution_plan": {
    "steps": ["步骤1", "步骤2"],
    "estimated_time": "5秒"
  }
}
```

"""

        if user_request:
            base_prompt += f"\n用户请求: {user_request}\n"

        return base_prompt

    def _parse_analysis_response(self, response: str) -> AnalysisResult:
        """解析分析响应"""
        result = AnalysisResult(success=True, response_text=response)

        try:
            # 尝试提取JSON
            if '{' in response:
                json_start = response.find('{')
                json_end = response.rfind('}') + 1
                json_str = response[json_start:json_end]
                data = json.loads(json_str)

                result.detected_issues = data.get('detected_issues', [])
                result.suggested_actions = data.get('suggested_actions', [])
                result.execution_plan = data.get('execution_plan', {})

        except json.JSONDecodeError:
            # 解析失败，将整个响应作为文本处理
            result.detected_issues = [response[:200]]

        return result

    def get_available_models(self) -> List[ModelInfo]:
        """获取可用模型列表"""
        if self._provider:
            return self._provider.list_models()
        return []

    def clear_history(self):
        """清空聊天历史"""
        self._chat_history = []

    def get_history(self) -> List[dict]:
        """获取聊天历史"""
        return [msg.to_dict() for msg in self._chat_history]

    @property
    def current_provider(self) -> str:
        return self._current_provider_type or 'openai'
