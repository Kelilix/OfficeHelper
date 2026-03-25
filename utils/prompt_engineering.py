"""
提示词工程模块
提供针对格式调整优化的提示词模板
"""

from typing import Dict, List, Any, Optional
from dataclasses import dataclass


@dataclass
class PromptTemplate:
    """提示词模板"""
    name: str
    template: str
    description: str
    variables: List[str]


class PromptEngine:
    """提示词引擎"""

    def __init__(self):
        self._templates = self._load_templates()

    def _load_templates(self) -> Dict[str, PromptTemplate]:
        """加载预设模板"""
        return {
            'analyze_screenshot': PromptTemplate(
                name="analyze_screenshot",
                template=self._ANALYZE_SCREENSHOT,
                description="分析截图中的文档格式问题",
                variables=['screenshot', 'user_request']
            ),
            'fix_format': PromptTemplate(
                name="fix_format",
                template=self._FIX_FORMAT,
                description="生成格式修复计划",
                variables=['issues', 'document_state']
            ),
            'explain_issue': PromptTemplate(
                name="explain_issue",
                template=self._EXPLAIN_ISSUE,
                description="解释格式问题",
                variables=['issue_type', 'context']
            ),
            'general_chat': PromptTemplate(
                name="general_chat",
                template=self._GENERAL_CHAT,
                description="通用对话",
                variables=['message', 'history']
            )
        }

    def get_template(self, name: str) -> Optional[PromptTemplate]:
        """获取模板"""
        return self._templates.get(name)

    def render(self, template_name: str, **kwargs) -> str:
        """
        渲染模板

        Args:
            template_name: 模板名称
            **kwargs: 模板变量

        Returns:
            str: 渲染后的提示词
        """
        template = self._templates.get(template_name)
        if not template:
            return ""

        result = template.template
        for key, value in kwargs.items():
            result = result.replace(f"{{{key}}}", str(value))

        return result

    def analyze_screenshot(self, user_request: str = "") -> str:
        """渲染截图分析提示"""
        return self.render('analyze_screenshot', user_request=user_request)

    def fix_format(self, issues: List[str], document_context: str = "") -> str:
        """渲染格式修复提示"""
        issues_text = "\n".join([f"- {issue}" for issue in issues])
        return self.render('fix_format', issues=issues_text, document_state=document_context)

    # 预设模板内容
    _ANALYZE_SCREENSHOT = """你是一个专业的Word文档格式分析助手。请分析下面的文档截图，指出格式问题。

分析要求：
1. 检查字体（名称、大小、粗细、颜色）是否一致
2. 检查段落对齐方式是否统一
3. 检查行距、段前段后间距是否合理
4. 检查页边距设置是否合适
5. 检查是否有空段落或多余空行
6. 检查标题样式是否规范

{user_request}

请以JSON格式返回分析结果：
```json
{
  "detected_issues": ["问题1描述", "问题2描述"],
  "issue_count": 3,
  "suggested_actions": [
    {
      "action": "set_font",
      "params": {"font_name": "微软雅黑", "size": 12},
      "target": "all",
      "description": "统一正文字体"
    }
  ],
  "health_score": 75
}
```
"""

    _FIX_FORMAT = """基于以下检测到的问题，请生成修复计划：

检测到的问题：
{issues}

文档状态：
{document_state}

请生成具体的修复步骤，以JSON格式返回：
```json
{
  "plan": [
    {
      "step": 1,
      "action": "set_font",
      "params": {...},
      "estimated_time": "2秒"
    }
  ],
  "total_time": "约10秒",
  "warnings": []
}
```
"""

    _EXPLAIN_ISSUE = """请解释以下格式问题：

问题类型：{issue_type}
上下文：{context}

请用通俗易懂的语言解释问题原因和影响，并提供解决方法。
"""

    _GENERAL_CHAT = """你是一个专业的Word文档格式调整助手。请回答用户的问题。

用户消息：{message}

历史对话：
{history}

请用中文回答，如果涉及到具体操作，请给出清晰的步骤说明。
"""


# 全局实例
prompt_engine = PromptEngine()
