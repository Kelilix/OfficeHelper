"""
响应解析器模块
解析大模型返回的响应，提取可执行的指令
"""

import json
import re
from typing import Optional, Dict, List, Any
from dataclasses import dataclass


@dataclass
class ParsedAction:
    """解析后的操作"""
    action: str
    params: Dict[str, Any]
    target: str = "selection"
    description: str = ""


@dataclass
class ExecutionPlan:
    """执行计划"""
    actions: List[ParsedAction]
    estimated_time: str = "0秒"
    warnings: List[str] = None

    def __post_init__(self):
        if self.warnings is None:
            self.warnings = []


class ResponseParser:
    """响应解析器"""

    def __init__(self):
        # 支持的操作类型
        self._action_types = {
            'set_font': self._parse_set_font,
            'set_font_size': self._parse_set_font_size,
            'set_bold': self._parse_set_bold,
            'set_italic': self._parse_set_italic,
            'set_alignment': self._parse_set_alignment,
            'set_line_spacing': self._parse_set_line_spacing,
            'set_page_size': self._parse_set_page_size,
            'set_margin': self._parse_set_margin,
            'apply_style': self._parse_apply_style,
            'delete_paragraph': self._parse_delete_paragraph,
            'insert_text': self._parse_insert_text,
        }

    def parse(self, response: str) -> ExecutionPlan:
        """
        解析响应文本为执行计划

        Args:
            response: 大模型返回的响应

        Returns:
            ExecutionPlan: 执行计划
        """
        # 尝试从JSON解析
        try:
            json_data = self._extract_json(response)
            if json_data:
                return self._parse_json_plan(json_data)
        except Exception as e:
            pass

        # 回退到文本解析
        return self._parse_text_plan(response)

    def _extract_json(self, text: str) -> Optional[dict]:
        """提取JSON数据"""
        # 查找JSON块
        patterns = [
            r'```json\s*([\s\S]*?)\s*```',
            r'```\s*([\s\S]*?)\s*```',
            r'\{\s*"[^"]*":\s*[\s\S]*\}'
        ]

        for pattern in patterns:
            matches = re.findall(pattern, text)
            for match in matches:
                try:
                    return json.loads(match)
                except:
                    continue

        return None

    def _parse_json_plan(self, data: dict) -> ExecutionPlan:
        """解析JSON格式的计划"""
        actions = []

        # 尝试多种JSON格式
        action_keys = ['actions', 'plan', 'suggested_actions', 'execution_plan']

        for key in action_keys:
            if key in data:
                action_list = data[key]
                if isinstance(action_list, list):
                    for action_data in action_list:
                        parsed = self._parse_action_data(action_data)
                        if parsed:
                            actions.append(parsed)
                break

        # 如果没有找到actions，尝试从detected_issues构建
        if not actions and 'detected_issues' in data:
            for issue in data.get('detected_issues', []):
                actions.append(ParsedAction(
                    action='analyze_issue',
                    params={'issue': issue},
                    description=issue
                ))

        # 获取预计时间
        estimated_time = data.get('estimated_time', data.get('total_time', '0秒'))

        # 获取警告
        warnings = data.get('warnings', [])

        return ExecutionPlan(
            actions=actions,
            estimated_time=estimated_time,
            warnings=warnings
        )

    def _parse_action_data(self, action_data: Any) -> Optional[ParsedAction]:
        """解析单个操作数据"""
        if isinstance(action_data, str):
            return ParsedAction(action=action_data, params={})

        if isinstance(action_data, dict):
            action = action_data.get('action', action_data.get('skill', ''))
            params = action_data.get('params', {})
            target = action_data.get('target', 'selection')
            description = action_data.get('description', action_data.get('desc', ''))

            # 解析action
            parser = self._action_types.get(action)
            if parser:
                parsed_params = parser(params)
                params.update(parsed_params)

            return ParsedAction(
                action=action,
                params=params,
                target=target,
                description=description
            )

        return None

    def _parse_text_plan(self, text: str) -> ExecutionPlan:
        """解析纯文本计划"""
        actions = []
        lines = text.split('\n')

        for line in lines:
            # 匹配常见的操作描述
            line = line.strip()
            if not line or line.startswith('#'):
                continue

            # 尝试匹配操作模式
            action = self._detect_action_from_text(line)
            if action:
                actions.append(action)

        return ExecutionPlan(
            actions=actions,
            estimated_time=f"约{len(actions) * 2}秒"
        )

    def _detect_action_from_text(self, text: str) -> Optional[ParsedAction]:
        """从文本检测操作"""
        text_lower = text.lower()

        # 字体设置
        if '字体' in text or 'font' in text_lower:
            font_name = self._extract_font_name(text)
            if font_name:
                return ParsedAction(
                    action='set_font',
                    params={'font_name': font_name},
                    description=text
                )

        # 字号设置
        if '字号' in text or '大小' in text:
            size = self._extract_size(text)
            if size:
                return ParsedAction(
                    action='set_font_size',
                    params={'size': size},
                    description=text
                )

        # 行距
        if '行距' in text or 'line' in text_lower:
            spacing = self._extract_spacing(text)
            if spacing:
                return ParsedAction(
                    action='set_line_spacing',
                    params={'spacing': spacing},
                    description=text
                )

        # 对齐
        for align in ['左对齐', '右对齐', '居中', '两端对齐']:
            if align in text:
                align_map = {'左对齐': 'left', '右对齐': 'right', '居中': 'center', '两端对齐': 'justify'}
                return ParsedAction(
                    action='set_alignment',
                    params={'alignment': align_map[align]},
                    description=text
                )

        return None

    def _extract_font_name(self, text: str) -> Optional[str]:
        """提取字体名称"""
        common_fonts = ['微软雅黑', '宋体', '黑体', '楷体', '仿宋', '华文细黑',
                       'Arial', 'Times New Roman', 'Calibri']

        for font in common_fonts:
            if font in text:
                return font

        return None

    def _extract_size(self, text: str) -> Optional[float]:
        """提取字号"""
        # 匹配 "XX号" 或 "XX磅" 或 "XXpt"
        patterns = [
            r'(\d+)\s*号',
            r'(\d+)\s*磅',
            r'(\d+)\s*pt',
        ]

        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                size = float(match.group(1))
                # 号数转换为磅
                if '号' in text:
                    size = self._hao_to_pt(size)
                return size

        return None

    def _hao_to_pt(self, hao: float) -> float:
        """字号转换"""
        # 字号到磅的转换（近似值）
        hao_pt_map = {
            1: 5, 2: 5.5, 3: 7.5, 4: 9, 5: 10.5,
            6: 12, 7: 13.5, 8: 15, 9: 16.5, 10: 18,
            11: 19.5, 12: 21, 14: 24, 16: 27, 18: 30,
            21: 33, 24: 36, 30: 42, 36: 48, 42: 54, 48: 60
        }
        return hao_pt_map.get(int(hao), 12)

    def _extract_spacing(self, text: str) -> Optional[float]:
        """提取行距"""
        patterns = [
            r'(\d+(?:\.\d+)?)\s*倍',
            r'([\d.]+)\s*倍行距',
        ]

        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return float(match.group(1))

        return None

    # 参数解析方法
    def _parse_set_font(self, params: dict) -> dict:
        return {}

    def _parse_set_font_size(self, params: dict) -> dict:
        return {}

    def _parse_set_bold(self, params: dict) -> dict:
        bold = params.get('bold', True)
        return {'bold': bold}

    def _parse_set_italic(self, params: dict) -> dict:
        italic = params.get('italic', True)
        return {'italic': italic}

    def _parse_set_alignment(self, params: dict) -> dict:
        return {}

    def _parse_set_line_spacing(self, params: dict) -> dict:
        return {}

    def _parse_set_page_size(self, params: dict) -> dict:
        return {}

    def _parse_set_margin(self, params: dict) -> dict:
        return {}

    def _parse_apply_style(self, params: dict) -> dict:
        return {}

    def _parse_delete_paragraph(self, params: dict) -> dict:
        return {}

    def _parse_insert_text(self, params: dict) -> dict:
        return {}


# 全局实例
response_parser = ResponseParser()
