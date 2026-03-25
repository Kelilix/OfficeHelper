"""
格式修复器模块
执行格式修复操作
"""

from typing import Dict, Any, List, Optional, Callable
from dataclasses import dataclass, field
from enum import Enum
import json


class FixActionType(Enum):
    """修复操作类型"""
    SET_FONT = "set_font"
    SET_FONT_SIZE = "set_font_size"
    SET_BOLD = "set_bold"
    SET_ITALIC = "set_italic"
    SET_ALIGNMENT = "set_alignment"
    SET_LINE_SPACING = "set_line_spacing"
    SET_INDENT = "set_indent"
    SET_PAGE_SIZE = "set_page_size"
    SET_MARGIN = "set_margin"
    APPLY_STYLE = "apply_style"
    DELETE_PARAGRAPH = "delete_paragraph"
    INSERT_PARAGRAPH = "insert_paragraph"
    ADD_PAGE_BREAK = "add_page_break"


@dataclass
class FixAction:
    """修复操作"""
    action_type: FixActionType
    params: Dict[str, Any]
    target: str = "selection"  # selection/all/paragraph_N
    description: str = ""


@dataclass
class ExecutionPlan:
    """执行计划"""
    actions: List[FixAction] = field(default_factory=list)
    estimated_time: str = "0秒"
    warning: str = ""


@dataclass
class ExecutionResult:
    """执行结果"""
    success: bool
    message: str = ""
    executed_actions: List[str] = field(default_factory=list)
    failed_actions: List[Dict[str, Any]] = field(default_factory=list)


class FormatFixer:
    """格式修复器"""

    def __init__(self, word_connector):
        self._word = word_connector
        self._action_history: List[Dict[str, Any]] = []
        self._max_history = 50

    def execute_plan(self, plan: ExecutionPlan) -> ExecutionResult:
        """
        执行修复计划

        Args:
            plan: 执行计划

        Returns:
            ExecutionResult: 执行结果
        """
        result = ExecutionResult(success=True)
        executed = []
        failed = []

        for action in plan.actions:
            try:
                success = self._execute_action(action)
                if success:
                    executed.append(action.action_type.value)
                    self._action_history.append({
                        'action': action,
                        'executed': True
                    })
                else:
                    failed.append({
                        'action': action.action_type.value,
                        'error': '执行失败'
                    })
            except Exception as e:
                failed.append({
                    'action': action.action_type.value,
                    'error': str(e)
                })

        # 限制历史长度
        if len(self._action_history) > self._max_history:
            self._action_history = self._action_history[-self._max_history:]

        result.executed_actions = executed
        result.failed_actions = failed
        result.success = len(failed) == 0
        result.message = f"成功执行{len(executed)}个操作"
        if failed:
            result.message += f"，{len(failed)}个失败"

        return result

    def _execute_action(self, action: FixAction) -> bool:
        """执行单个操作"""
        action_map = {
            FixActionType.SET_FONT: self._set_font,
            FixActionType.SET_FONT_SIZE: self._set_font_size,
            FixActionType.SET_BOLD: self._set_bold,
            FixActionType.SET_ITALIC: self._set_italic,
            FixActionType.SET_ALIGNMENT: self._set_alignment,
            FixActionType.SET_LINE_SPACING: self._set_line_spacing,
            FixActionType.SET_PAGE_SIZE: self._set_page_size,
            FixActionType.SET_MARGIN: self._set_margin,
            FixActionType.APPLY_STYLE: self._apply_style,
        }

        handler = action_map.get(action.action_type)
        if handler:
            return handler(action.params, action.target)
        return False

    def _set_font(self, params: dict, target: str) -> bool:
        """设置字体"""
        font_name = params.get('font_name')
        if not font_name:
            return False
        return self._word.set_font(font_name=font_name)

    def _set_font_size(self, params: dict, target: str) -> bool:
        """设置字号"""
        size = params.get('size')
        if not size:
            return False
        return self._word.set_font(size=size)

    def _set_bold(self, params: dict, target: str) -> bool:
        """设置加粗"""
        bold = params.get('bold', True)
        return self._word.set_font(bold=bold)

    def _set_italic(self, params: dict, target: str) -> bool:
        """设置斜体"""
        italic = params.get('italic', True)
        return self._word.set_font(italic=italic)

    def _set_alignment(self, params: dict, target: str) -> bool:
        """设置对齐"""
        alignment = params.get('alignment', 'left')
        return self._word.set_paragraph_alignment(alignment)

    def _set_line_spacing(self, params: dict, target: str) -> bool:
        """设置行距"""
        spacing = params.get('spacing', 1.5)
        return self._word.set_line_spacing(spacing)

    def _set_page_size(self, params: dict, target: str) -> bool:
        """设置页面大小"""
        from .word_connector import PageSetup
        ps = PageSetup()
        ps.paper_size = params.get('paper_size', 'A4')
        return self._word.set_page_setup(ps)

    def _set_margin(self, params: dict, target: str) -> bool:
        """设置边距"""
        from .word_connector import PageSetup
        ps = PageSetup()
        ps.top_margin = params.get('top', ps.top_margin)
        ps.bottom_margin = params.get('bottom', ps.bottom_margin)
        ps.left_margin = params.get('left', ps.left_margin)
        ps.right_margin = params.get('right', ps.right_margin)
        return self._word.set_page_setup(ps)

    def _apply_style(self, params: dict, target: str) -> bool:
        """应用样式"""
        style_name = params.get('style_name')
        if not style_name:
            return False
        return self._word.apply_style(style_name)

    def create_plan_from_llm_response(self, llm_response: str) -> ExecutionPlan:
        """
        从LLM响应创建执行计划

        Args:
            llm_response: LLM返回的响应文本

        Returns:
            ExecutionPlan: 执行计划
        """
        plan = ExecutionPlan()

        try:
            # 尝试解析JSON
            if '{' in llm_response:
                json_str = llm_response[llm_response.find('{'):llm_response.rfind('}')+1]
                data = json.loads(json_str)

                if 'actions' in data:
                    for action_data in data['actions']:
                        action_type = FixActionType(action_data.get('action', 'set_font'))
                        params = action_data.get('params', {})
                        target = action_data.get('target', 'selection')
                        plan.actions.append(FixAction(
                            action_type=action_type,
                            params=params,
                            target=target,
                            description=action_data.get('description', '')
                        ))

                if 'estimated_time' in data:
                    plan.estimated_time = data['estimated_time']

        except Exception as e:
            # 解析失败，尝试简单解析
            plan.warning = f"解析LLM响应时出现问题: {str(e)}"

        return plan

    def preview_plan(self, plan: ExecutionPlan) -> str:
        """
        预览执行计划

        Args:
            plan: 执行计划

        Returns:
            str: 预览文本
        """
        preview_lines = ["执行计划预览:", ""]

        for i, action in enumerate(plan.actions, 1):
            desc = action.description or self._get_action_description(action)
            preview_lines.append(f"{i}. {desc}")

        preview_lines.append("")
        preview_lines.append(f"预计耗时: {plan.estimated_time}")

        if plan.warning:
            preview_lines.append(f"警告: {plan.warning}")

        return "\n".join(preview_lines)

    def _get_action_description(self, action: FixAction) -> str:
        """获取操作描述"""
        desc_map = {
            FixActionType.SET_FONT: f"设置字体为 {action.params.get('font_name', '')}",
            FixActionType.SET_FONT_SIZE: f"设置字号为 {action.params.get('size', '')}磅",
            FixActionType.SET_BOLD: f"{'设置' if action.params.get('bold') else '取消'}加粗",
            FixActionType.SET_ITALIC: f"{'设置' if action.params.get('italic') else '取消'}斜体",
            FixActionType.SET_ALIGNMENT: f"设置对齐方式为 {action.params.get('alignment', '')}",
            FixActionType.SET_LINE_SPACING: f"设置行距为 {action.params.get('spacing', '')}倍",
            FixActionType.SET_PAGE_SIZE: f"设置纸张大小为 {action.params.get('paper_size', '')}",
            FixActionType.SET_MARGIN: f"设置页边距 (上下左右)",
            FixActionType.APPLY_STYLE: f"应用样式 {action.params.get('style_name', '')}",
        }
        return desc_map.get(action.action_type, f"执行 {action.action_type.value}")

    def undo_last(self) -> bool:
        """撤销上一步操作"""
        if self._word.can_undo():
            return self._word.undo()
        return False

    def redo_last(self) -> bool:
        """重做上一步操作"""
        if self._word.can_redo():
            return self._word.redo()
        return False

    def get_history(self) -> List[dict]:
        """获取操作历史"""
        return [
            {
                'action': h['action'].action_type.value,
                'description': h['action'].description or self._get_action_description(h['action']),
                'executed': h['executed']
            }
            for h in self._action_history
        ]
