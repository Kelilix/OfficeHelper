"""
AI 聊天路由

内部两轮 LLM 调用（对用户单次 HTTP 请求透明）：
  Round 1: 用户需求 + skill 摘要 → LLM 选技能
  Round 2: 完整 SKILL.md → LLM 产出 action 并由本服务执行

可选请求参数 skill_name：若传入则跳过 Round 1，直接执行该技能（调试 / 高级用法）。
"""

import re
import json
import uuid
import logging
from typing import Optional, List, Dict, Any
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from .service import word_service, llm_service
from skills import get_skill_descriptions, get_skill_content, list_skill_names

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api", tags=["chat"])


# ── 请求/响应模型 ────────────────────────────────────────────────

class ChatRequest(BaseModel):
    message: str
    selection_text: str = ""
    document_name: str = ""
    session_id: str = ""          # 会话 ID，来自前端（如前端不传则后端自动生成）
    skill_name: str = ""           # 可选：指定后跳过自动选技能，直接按该技能执行


class WordStatusResponse(BaseModel):
    connected: bool
    document_name: Optional[str] = None
    has_selection: bool = False
    selection_text: str = ""


class ChatResponse(BaseModel):
    response: str
    success: bool = True
    error: Optional[str] = None
    session_id: str = ""
    turn: int = 0
    stage: str = "execute"       # "select" = 未能匹配技能；"execute" = 已走完执行流程
    skill_selected: Optional[str] = None   # 实际使用的技能目录名（成功执行时）


# ── 路由实现 ──────────────────────────────────────────────────────

@router.post("/chat", response_model=ChatResponse)
def chat(req: ChatRequest) -> ChatResponse:
    """
    单次请求内完成：自动选技能（可选跳过）→ 加载说明书 → LLM 生成 action → 执行 Word 操作。

    仅在无法匹配任何技能时返回 stage="select" 与提示文案；成功时直接返回执行结果摘要。
    """
    try:
        session_id = req.session_id.strip() or str(uuid.uuid4())
        history_before = len(llm_service.get_session_history(session_id))
        current_turn = (history_before // 2) + 1

        skill_explicit = (req.skill_name or "").strip()
        logger.info(
            "[/api/chat] 收到请求 | session_id=%s turn=%d direct_skill=%s msg_len=%d",
            session_id,
            current_turn,
            bool(skill_explicit),
            len(req.message or ""),
        )

        # 实时刷新选区
        current_selection = req.selection_text
        if not current_selection:
            try:
                current_selection = word_service.get_selection_text()
            except Exception:
                current_selection = ""

        skill_to_run = skill_explicit

        # ── Round 1：自动选技能（未显式指定 skill_name 时）────────────────────
        if not skill_to_run:
            skills_desc = get_skill_descriptions()
            prompt = _build_prompt_select(req.message, current_selection, skills_desc)
            llm_select = llm_service.chat_with_context(
                req.message,
                prompt,
                session_id=session_id,
            )
            skill_to_run = _parse_skill_selection(llm_select) or ""

            logger.info(
                "[/api/chat] 内部选技能完成 | session_id=%s skill=%s",
                session_id,
                skill_to_run or "(未识别)",
            )

            if not skill_to_run:
                return ChatResponse(
                    response="无法识别所需技能，请重述需求。",
                    success=True,
                    session_id=session_id,
                    turn=current_turn,
                    stage="select",
                    skill_selected=None,
                )

        # ── Round 2：按技能说明书执行 ────────────────────────────────────────
        skill_content = get_skill_content(skill_to_run)
        if not skill_content:
            raise ValueError(
                f"未找到技能：{skill_to_run}（请检查 skills/ 目录下是否存在对应目录）"
            )

        prompt = _build_prompt_execute(
            req.message,
            current_selection,
            skill_to_run,
            skill_content,
        )
        llm_response = llm_service.chat_with_context(
            req.message,
            prompt,
            session_id=session_id,
        )

        actions = _parse_actions(llm_response)
        executed = []
        for action in actions:
            result = _execute_action(action)
            executed.append(result)

        summary = _summarize_execution(llm_response, executed)

        final_turn = len(llm_service.get_session_history(session_id)) // 2

        logger.info(
            "[/api/chat] 执行完成 | session_id=%s skill=%s actions=%d",
            session_id,
            skill_to_run,
            len(actions),
        )
        return ChatResponse(
            response=summary,
            success=True,
            session_id=session_id,
            turn=final_turn,
            stage="execute",
            skill_selected=skill_to_run,
        )

    except Exception as e:
        logger.exception("[/api/chat] 失败: %s", e)
        return ChatResponse(
            response=f"处理失败：{str(e)}",
            success=False,
            error=str(e),
            session_id=req.session_id or "",
            stage="execute",
        )


@router.get("/chat/history")
def chat_history(session_id: str) -> dict:
    """
    获取指定 session_id 的多轮对话历史。

    返回格式：
    {
        "session_id": "...",
        "turns": [
            {"轮次": 1, "用户需求": "...", "回答": "..."},
            {"轮次": 2, "用户需求": "比之前小一些", "回答": "..."},
            ...
        ],
        "count": 4
    }
    """
    if not session_id:
        raise HTTPException(status_code=400, detail="session_id 不能为空")
    history = llm_service.get_session_history(session_id)
    return {
        "session_id": session_id,
        "turns": history,
        "count": len(history),
    }


@router.delete("/chat/clear")
def chat_clear(session_id: str) -> dict:
    """
    清空指定 session_id 的对话历史。

    用于：用户点击"新建对话"、切换文档窗口、或想重新开始对话时。
    """
    if not session_id:
        raise HTTPException(status_code=400, detail="session_id 不能为空")
    cleared = llm_service.clear_session(session_id)
    logger.info("[/api/chat/clear] session_id=%s cleared=%s", session_id, cleared)
    return {"session_id": session_id, "cleared": cleared}


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

def _build_prompt_select(user_message: str, selection_text: str, skills_desc: str) -> str:
    """
    Round 1 prompt：只让 LLM 根据 skill name + description 选择技能。
    token 消耗极低（仅传入 name+description，无完整说明书）。
    """
    return f"""你是一个专业的 Word 文档格式化助手。

## 当前上下文
用户选中了 Word 文档中的以下内容：
---
{selection_text if selection_text else "(无选中文本)"}
---

## 用户的需求
"{user_message}"

## 可用的技能列表
{skills_desc}

## 你的任务
仔细分析用户需求，从上方技能列表中选择最合适的一个（只选一个）。
如果不需要任何技能（如用户只是闲聊或问题咨询），返回 {{"skill": "", "reasoning": "不需要技能"}}。
如果判断需要某个技能，只返回以下 JSON 格式，不要包含其他文字：
```json
{{"skill": "技能目录名", "reasoning": "简短选择理由"}}
```
"""


def _build_prompt_execute(
    user_message: str,
    selection_text: str,
    skill_name: str,
    skill_content: str,
) -> str:
    """
    Round 2 prompt：传入选中技能的完整 SKILL.md 说明书，让 LLM 执行操作。
    skill_content 是被选中技能的完整 SKILL.md 内容（含所有操作细节）。
    """
    return f"""你是一个专业的 Word 文档格式化助手，正在使用技能「{skill_name}」。

## 当前上下文
用户选中了 Word 文档中的以下内容：
---
{selection_text if selection_text else "(无选中文本)"}
---

## 用户的需求
"{user_message}"

## 技能「{skill_name}」的完整说明书
{skill_content}

## 你的任务
1. 仔细阅读上方技能说明书，理解所有可用操作
2. 根据用户需求，决定需要调用哪些操作
3. 返回 JSON 数组表示要执行的操作列表

## 输出要求
只返回以下 JSON 格式，不要包含其他说明文字：
[
  {{
    "action": "操作名（必须与技能说明书中的一致）",
    "params": {{"参数名": "参数值"}},
    "description": "操作描述"
  }}
]

如果用户只是闲聊或问题咨询，不需要执行任何 Word 操作，请返回：[]
"""


def _parse_skill_selection(llm_response: str) -> Optional[str]:
    """
    从 Round 1 LLM 响应中解析出选中的 skill 目录名。
    返回 None 表示无法识别。
    """
    try:
        match = re.search(r'\{[^}]*"skill"\s*:\s*"([^"]+)"[^}]*\}', llm_response)
        if match:
            skill = match.group(1).strip()
            if skill:
                valid_names = list_skill_names()
                if skill in valid_names:
                    return skill
                # 尝试模糊匹配（忽略大小写）
                for name in valid_names:
                    if name.lower() == skill.lower():
                        return name
                    if skill.lower() in name.lower():
                        return name
        logger.warning("[_parse_skill_selection] 无法从 LLM 响应中解析 skill：%s", llm_response[:200])
    except Exception as e:
        logger.warning("[_parse_skill_selection] 解析异常：%s，原始响应：%s", e, llm_response[:200])
    return None


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
