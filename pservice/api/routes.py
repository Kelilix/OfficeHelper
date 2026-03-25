"""
AI 聊天路由
"""

import re
import json
from typing import Optional, List, Dict, Any
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from .service import word_service, llm_service
from skills import get_skill_descriptions

router = APIRouter(prefix="/api", tags=["chat"])


# ── 请求/响应模型 ────────────────────────────────────────────────

class ChatRequest(BaseModel):
    message: str
    selection_text: str = ""
    document_name: str = ""


class ChatResponse(BaseModel):
    response: str
    success: bool = True
    error: Optional[str] = None


class WordStatusResponse(BaseModel):
    connected: bool
    document_name: Optional[str] = None
    has_selection: bool = False
    selection_text: str = ""


# ── 路由实现 ──────────────────────────────────────────────────────

@router.post("/chat", response_model=ChatResponse)
def chat(req: ChatRequest) -> ChatResponse:
    """
    核心接口：接收用户消息 + 选中文本，调用 LLM 分析意图并执行 Word 操作。
    """
    try:
        # 实时刷新选区（来自 Word）
        current_selection = req.selection_text
        if not current_selection:
            try:
                current_selection = word_service.get_selection_text()
            except Exception:
                current_selection = ""

        skills_desc = get_skill_descriptions()
        prompt = _build_prompt(req.message, current_selection, skills_desc)

        # 用 chat_with_context 传递完整上下文，系统消息会被正确构建
        llm_response = llm_service.chat_with_context(req.message, prompt)

        # 解析 LLM 返回的 JSON action 列表并执行
        actions = _parse_actions(llm_response)
        executed = []
        for action in actions:
            result = _execute_action(action)
            executed.append(result)

        # 生成自然语言摘要
        summary = _summarize_execution(llm_response, executed)

        return ChatResponse(response=summary, success=True)

    except Exception as e:
        return ChatResponse(response=f"处理失败：{str(e)}", success=False, error=str(e))


@router.get("/word/status", response_model=WordStatusResponse)
def word_status() -> WordStatusResponse:
    """返回 Word 连接状态"""
    try:
        connected = word_service.is_connected()
        has_sel = word_service.has_selection() if connected else False
        sel_text = word_service.get_selection_text() if has_sel else ""
        doc_name = word_service.get_document_name() if connected else None
        return WordStatusResponse(
            connected=connected,
            document_name=doc_name,
            has_selection=has_sel,
            selection_text=sel_text,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/word/connect")
def word_connect() -> dict:
    """显式连接 Word"""
    try:
        ok = word_service.connect(visible=True)
        return {"success": ok, "connected": ok}
    except Exception as e:
        return {"success": False, "error": str(e)}


@router.post("/word/disconnect")
def word_disconnect(save: bool = False) -> dict:
    """断开 Word 连接"""
    try:
        word_service.quit(save_changes=save)
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


# ── 内部函数 ──────────────────────────────────────────────────────

def _build_prompt(user_message: str, selection_text: str, skills_desc: str) -> str:
    return f"""你是一个专业的 Word 文档格式化助手。

## 当前上下文
用户选中了 Word 文档中的以下内容：
---
{selection_text if selection_text else "(无选中文本)"}
---

## 用户的需求
"{user_message}"

## 可用的格式化技能
{skills_desc}

## 你的任务
1. 理解用户想要对选中文本做什么修改
2. 根据需求，调用相应的 Word 操作（见上方技能列表）
3. 你必须返回一个 JSON 数组，表示要执行的操作列表

## 输出要求
请返回 JSON 数组，格式如下：
[
  {{
    "action": "set_font_size",
    "params": {{"size": 12}},
    "description": "设置字号为小三"
  }},
  {{
    "action": "set_bold",
    "params": {{"bold": true}},
    "description": "加粗"
  }}
]

如果用户只是聊天而不需要格式修改，也返回空数组 [] 并在回复文本中说明。
"""


def _parse_actions(llm_response: str) -> List[Dict[str, Any]]:
    """从 LLM 响应中提取 JSON action 数组"""
    try:
        match = re.search(r'\[.*\]', llm_response, re.DOTALL)
        if match:
            return json.loads(match.group())
    except (json.JSONDecodeError, re.error):
        pass
    return []


def _execute_action(action: Dict[str, Any]) -> Dict[str, Any]:
    """将单个 action 分发到 WordService"""
    action_type = action.get("action", "")
    params = action.get("params", {}) or {}
    desc = action.get("description", action_type)

    success = False
    try:
        if action_type == "set_font":
            success = word_service.set_font(font_name=params.get("font_name"))
        elif action_type == "set_font_size":
            success = word_service.set_font(size=float(params.get("size", 12)))
        elif action_type == "set_bold":
            success = word_service.set_font(bold=params.get("bold", True))
        elif action_type == "set_italic":
            success = word_service.set_font(italic=params.get("italic", True))
        elif action_type == "set_underline":
            success = word_service.set_font(underline=params.get("underline", True))
        elif action_type == "set_font_color":
            success = word_service.set_font_color(params.get("color", "000000"))
        elif action_type == "set_alignment":
            success = word_service.set_alignment(params.get("alignment", "left"))
        elif action_type == "set_line_spacing":
            success = word_service.set_line_spacing(float(params.get("spacing", 1.0)))
        elif action_type == "set_indent":
            success = word_service.set_indent(
                first_line=params.get("first_line", 21),
                left_indent=params.get("left_indent"),
                right_indent=params.get("right_indent"),
                indent_type=params.get("indent_type", "first_line"),
            )
        elif action_type == "set_paragraph_spacing":
            success = word_service.set_paragraph_spacing(
                before=params.get("space_before"),
                after=params.get("space_after"),
            )
        elif action_type == "set_page_margins":
            success = word_service.set_page_margins(
                top=params.get("top"),
                bottom=params.get("bottom"),
                left=params.get("left"),
                right=params.get("right"),
            )
        elif action_type == "set_paper_size":
            success = word_service.set_paper_size(params.get("paper_size", "A4"))
        elif action_type == "set_page_orientation":
            success = word_service.set_page_orientation(params.get("orientation", "portrait"))
    except Exception as e:
        return {"action": action_type, "description": desc, "success": False, "error": str(e)}

    return {"action": action_type, "description": desc, "success": success}


def _summarize_execution(llm_response: str, executed: List[Dict[str, Any]]) -> str:
    """从 LLM 响应中提取自然语言摘要（去掉 JSON 部分）"""
    # 去掉 JSON 数组部分，只保留前后的说明文字
    text = re.sub(r'\[.*\]', '', llm_response, flags=re.DOTALL).strip()

    if not text:
        # 没有自然语言，说明全是 JSON，生成摘要
        if executed:
            names = [e.get("description", e.get("action", "?")) for e in executed if e.get("success")]
            if names:
                return f"✅ 已执行：{', '.join(names)}"
        return "已完成处理。"

    # 取第一段作为摘要（去掉 markdown 代码块标记）
    text = re.sub(r'```json|```', '', text).strip()
    first_para = text.split('\n')[0].strip()
    return first_para if first_para else "已完成处理。"
