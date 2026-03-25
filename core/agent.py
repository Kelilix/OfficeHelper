"""
Agent 模块 - AI 智能体
负责理解用户意图、选择 Skills、执行 Word 操作
"""

import re
import json
from typing import Optional, Dict, Any, List, Callable
from dataclasses import dataclass


@dataclass
class AgentResponse:
    """Agent 响应"""
    success: bool
    message: str
    executed_skills: List[Dict[str, Any]] = None
    code_snippet: str = ""
    error: Optional[str] = None


class WordAgent:
    """Word 文档 AI 代理"""

    def __init__(self, word_connector, llm_service, screenshot_manager=None):
        self.word = word_connector
        self.llm = llm_service
        self.screenshot = screenshot_manager

    def process(self, user_message: str, selection_text: str = "", selection_range: tuple = None) -> AgentResponse:
        """
        处理用户请求

        Args:
            user_message: 用户的消息
            selection_text: 选中的文本
            selection_range: 选区范围 (start, end)

        Returns:
            AgentResponse: 处理结果
        """
        try:
            # 检查选区
            if not self.word.has_selection():
                return AgentResponse(
                    success=False,
                    message="请先在 Word 中选中要修改的文字"
                )

            # 获取最新的选区文本（确保是最新的）
            current_selection = self.word.get_selection_text()

            # 分析用户意图
            intent = self._analyze_intent(user_message, current_selection)

            if not intent:
                return AgentResponse(
                    success=False,
                    message="抱歉，我不太理解你的需求。你可以告诉我：\n"
                            "• 把这段文字调小一号字体\n"
                            "• 首行缩进\n"
                            "• 加粗这段文字\n"
                            "• 居中对齐"
                )

            # 执行操作
            results = []
            for action in intent:
                result = self._execute_action(action)
                results.append(result)

            # 汇总结果
            success_count = sum(1 for r in results if r['success'])
            total_count = len(results)

            if success_count == total_count:
                message = f"✅ 已完成 {success_count} 项修改"
                if total_count > 1:
                    details = [r['desc'] for r in results if r['success']]
                    message += ":\n" + "\n".join(f"• {d}" for d in details)
                return AgentResponse(
                    success=True,
                    message=message,
                    executed_skills=results
                )
            else:
                message = f"⚠️ 完成 {success_count}/{total_count} 项修改"
                return AgentResponse(
                    success=False,
                    message=message,
                    executed_skills=results
                )

        except Exception as e:
            return AgentResponse(
                success=False,
                message=f"执行出错: {str(e)}",
                error=str(e)
            )

    def _analyze_intent(self, user_message: str, selection_text: str) -> List[Dict[str, Any]]:
        """
        分析用户意图

        Returns:
            List of actions to perform
        """
        # 构造分析提示
        prompt = f"""分析用户的格式化需求。

选中的文本：
{selection_text[:200] if selection_text else "(无)"}

用户需求：
{user_message}

请分析用户想要执行什么格式化操作，返回 JSON 数组格式：
[
  {{"action": "操作类型", "params": {{"参数": "值"}}}}
]

操作类型对应：
- "set_font": 设置字体 (font_name)
- "set_font_size": 设置字号 (size: 数字)
- "set_bold": 加粗 (bold: true/false)
- "set_italic": 斜体 (italic: true/false)
- "set_underline": 下划线 (underline: true/false)
- "set_font_color": 字体颜色 (color: 颜色名或十六进制)
- "set_alignment": 对齐方式 (alignment: left/center/right/justify)
- "set_line_spacing": 行距 (spacing: 倍数)
- "set_indent": 缩进 (first_line/left/right: 磅值)
- "set_paragraph_spacing": 段落间距 (space_before/space_after: 磅值)

如果用户没有指定具体值，使用默认值（如字号：四号=14磅，五号=10.5磅）。
只返回 JSON，不要其他内容。"""

        try:
            response = self.llm.chat(prompt)
            # 解析 JSON
            return self._parse_intent_response(response)
        except Exception as e:
            # 如果 LLM 调用失败，尝试本地解析
            return self._local_parse_intent(user_message)

    def _parse_intent_response(self, response: str) -> List[Dict[str, Any]]:
        """解析 LLM 返回的 JSON"""
        try:
            # 尝试提取 JSON
            json_match = re.search(r'\[.*\]', response, re.DOTALL)
            if json_match:
                actions = json.loads(json_match.group())
                return actions
        except json.JSONDecodeError:
            pass
        return []

    def _local_parse_intent(self, message: str) -> List[Dict[str, Any]]:
        """本地解析意图（备选方案）"""
        message = message.lower()
        actions = []

        # 加粗
        if any(kw in message for kw in ['加粗', '粗体', 'bold']) and '取消' not in message:
            actions.append({"action": "set_bold", "params": {"bold": True}})

        # 取消加粗
        if any(kw in message for kw in ['取消加粗', '取消粗体']) or \
           ('加粗' in message and '取消' in message):
            actions.append({"action": "set_bold", "params": {"bold": False}})

        # 斜体
        if '斜体' in message:
            if '取消' in message:
                actions.append({"action": "set_italic", "params": {"italic": False}})
            else:
                actions.append({"action": "set_italic", "params": {"italic": True}})

        # 下划线
        if '下划线' in message:
            if '取消' in message:
                actions.append({"action": "set_underline", "params": {"underline": False}})
            else:
                actions.append({"action": "set_underline", "params": {"underline": True}})

        # 字号
        size_map = {
            '一号': 26, '二号': 22, '小三': 15, '三号': 16, '小二': 18,
            '四号': 14, '小四': 12, '五号': 10.5, '小五': 9,
            '特大': 36, '特小': 7
        }
        for name, size in size_map.items():
            if name in message:
                actions.append({"action": "set_font_size", "params": {"size": size}})
                break
        else:
            # 尝试解析 "调小"、"调大"、"改成 XX 号" 等
            if '调小' in message or '缩小' in message:
                match = re.search(r'(\d+)号', message)
                if match:
                    size = size_map.get(f"{match.group(1)}号", int(match.group(1)))
                    actions.append({"action": "set_font_size", "params": {"size": size}})
            elif '调大' in message or '放大' in message:
                match = re.search(r'(\d+)号', message)
                if match:
                    size = size_map.get(f"{match.group(1)}号", int(match.group(1)))
                    actions.append({"action": "set_font_size", "params": {"size": size}})

        # 对齐
        if any(kw in message for kw in ['居中', '中间']):
            actions.append({"action": "set_alignment", "params": {"alignment": "center"}})
        if any(kw in message for kw in ['左对齐', '靠左']):
            actions.append({"action": "set_alignment", "params": {"alignment": "left"}})
        if any(kw in message for kw in ['右对齐', '靠右']):
            actions.append({"action": "set_alignment", "params": {"alignment": "right"}})
        if any(kw in message for kw in ['两端对齐', '对齐']):
            actions.append({"action": "set_alignment", "params": {"alignment": "justify"}})

        # 首行缩进
        if '缩进' in message:
            # 默认 2 字符 ≈ 21 磅
            indent_size = 21
            if '字符' in message:
                match = re.search(r'(\d+)字符', message)
                if match:
                    indent_size = int(match.group(1)) * 10.5
            actions.append({"action": "set_indent", "params": {"indent_type": "first_line", "first_line": indent_size}})

        # 行距
        if '1.5' in message or '一倍半' in message or '1.5倍' in message:
            actions.append({"action": "set_line_spacing", "params": {"spacing": 1.5}})
        elif '2倍' in message or '双倍' in message:
            actions.append({"action": "set_line_spacing", "params": {"spacing": 2.0}})
        elif '单倍' in message or '一倍' in message:
            actions.append({"action": "set_line_spacing", "params": {"spacing": 1.0}})

        # 字体颜色
        color_map = {
            '黑色': '000000', '红色': 'FF0000', '蓝色': '0000FF',
            '绿色': '00FF00', '白色': 'FFFFFF', '黄色': 'FFFF00'
        }
        for color_name, color_value in color_map.items():
            if color_name in message:
                actions.append({"action": "set_font_color", "params": {"color": color_value}})
                break

        return actions

    def _execute_action(self, action: Dict[str, Any]) -> Dict[str, Any]:
        """执行单个操作"""
        action_type = action.get('action', '')
        params = action.get('params', {})

        desc_map = {
            'set_font': f"设置字体为 {params.get('font_name', '')}",
            'set_font_size': f"设置字号为 {params.get('size', '')} 磅",
            'set_bold': f"设置加粗: {params.get('bold', False)}",
            'set_italic': f"设置斜体: {params.get('italic', False)}",
            'set_underline': f"设置下划线: {params.get('underline', False)}",
            'set_font_color': f"设置字体颜色: {params.get('color', '')}",
            'set_alignment': f"设置对齐方式: {params.get('alignment', '')}",
            'set_line_spacing': f"设置行距: {params.get('spacing', '')} 倍",
            'set_indent': f"设置缩进: {params.get('first_line', 0)} 磅",
            'set_paragraph_spacing': f"设置段落间距"
        }

        try:
            success = False

            if action_type == 'set_font':
                success = self.word.set_font(font_name=params.get('font_name'))

            elif action_type == 'set_font_size':
                success = self.word.set_font(size=float(params.get('size', 12)))

            elif action_type == 'set_bold':
                success = self.word.set_font(bold=params.get('bold', True))

            elif action_type == 'set_italic':
                success = self.word.set_font(italic=params.get('italic', True))

            elif action_type == 'set_underline':
                success = self.word.set_font(underline=params.get('underline', True))

            elif action_type == 'set_font_color':
                success = self.word.set_font_color(params.get('color', '000000'))

            elif action_type == 'set_alignment':
                success = self.word.set_alignment(params.get('alignment', 'left'))

            elif action_type == 'set_line_spacing':
                success = self.word.set_line_spacing(float(params.get('spacing', 1.0)))

            elif action_type == 'set_indent':
                indent_type = params.get('indent_type', 'first_line')
                if indent_type == 'first_line':
                    success = self.word.set_indent(
                        first_line=int(params.get('first_line', 21)),
                        indent_type='first_line'
                    )
                elif indent_type == 'left':
                    success = self.word.set_indent(
                        left_indent=int(params.get('left', 0)),
                        indent_type='left'
                    )
                elif indent_type == 'right':
                    success = self.word.set_indent(
                        right_indent=int(params.get('right', 0)),
                        indent_type='right'
                    )

            elif action_type == 'set_paragraph_spacing':
                before = params.get('space_before')
                after = params.get('space_after')
                success = self.word.set_paragraph_spacing(before=before, after=after)

            return {
                'action': action_type,
                'params': params,
                'desc': desc_map.get(action_type, action_type),
                'success': success
            }

        except Exception as e:
            return {
                'action': action_type,
                'params': params,
                'desc': desc_map.get(action_type, action_type),
                'success': False,
                'error': str(e)
            }

    def get_available_skills(self) -> str:
        """获取可用技能列表"""
        from skills import get_skill_descriptions
        return get_skill_descriptions()
