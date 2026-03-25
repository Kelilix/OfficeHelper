"""
大模型服务模块
提供多提供商的大模型调用支持
"""

import os
import json
import base64
import io
from typing import Optional, List, Dict, Any, Callable
from dataclasses import dataclass, field
from enum import Enum
from abc import ABC, abstractmethod


class LLMProvider(Enum):
    """大模型提供商"""
    OPENAI = "openai"
    ANTHROPIC = "anthropic"
    OLLAMA = "ollama"


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


class LLMService:
    """大模型服务"""

    def __init__(self, config: Optional[dict] = None):
        self._config = config or {}
        self._provider: Optional[BaseLLMProvider] = None
        self._current_provider_type = None
        self._chat_history: List[ChatMessage] = []

        # 默认系统提示
        self._system_prompt = """你是一个专业的Word文档格式调整助手。
你的任务是帮助用户分析和修复Word文档的格式问题。

你可以执行以下操作：
1. 分析文档截图，识别格式问题
2. 提供修复建议和执行步骤
3. 回答关于Word格式调整的问题

请用中文回复，并给出具体的操作建议。"""

        self._init_provider()

    def _init_provider(self):
        """初始化提供商"""
        provider_type = self._config.get('provider', 'openai').lower()
        api_key = self._config.get('api_key', '')
        model = self._config.get('model', 'gpt-4')
        base_url = self._config.get('base_url', '')

        if provider_type == 'openai':
            if not api_key:
                api_key = os.getenv('OPENAI_API_KEY', '')
            self._provider = OpenAIProvider(api_key, base_url or None, model)

        elif provider_type == 'anthropic':
            if not api_key:
                api_key = os.getenv('ANTHROPIC_API_KEY', '')
            self._provider = AnthropicProvider(api_key, model)

        elif provider_type == 'ollama':
            self._provider = OllamaProvider(base_url or "http://localhost:11434", model)

        self._current_provider_type = provider_type

    def set_provider(self, provider_type: str, **kwargs):
        """切换提供商"""
        self._config['provider'] = provider_type

        if provider_type == 'openai':
            self._provider = OpenAIProvider(
                kwargs.get('api_key', ''),
                kwargs.get('base_url', ''),
                kwargs.get('model', 'gpt-4')
            )
        elif provider_type == 'anthropic':
            self._provider = AnthropicProvider(
                kwargs.get('api_key', ''),
                kwargs.get('model', 'claude-3-opus-20240229')
            )
        elif provider_type == 'ollama':
            self._provider = OllamaProvider(
                kwargs.get('base_url', 'http://localhost:11434'),
                kwargs.get('model', 'llama2')
            )

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
        response = self._provider.chat(self._chat_history, stream=stream)
        if not stream:
            self._chat_history.append(ChatMessage(role="assistant", content=response))
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
        response = self._provider.chat(self._chat_history, stream=stream)
        if not stream:
            self._chat_history.append(ChatMessage(role="assistant", content=response))
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

        try:
            response = self._provider.analyze_image(image_data, prompt)
            return self._parse_analysis_response(response)
        except Exception as e:
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
