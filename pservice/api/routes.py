# -*- coding: utf-8 -*-
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
import sys
import logging
from pathlib import Path
from typing import Optional, List, Dict, Any
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from .service import word_service, llm_service
from skills import get_skill_descriptions, get_skill_content, list_skill_names
from .action_registry import execute_action

# 将 skills/word-text-operator 加入 sys.path，使 scripts/ 下的相对导入可用
_scripts_parent = Path(__file__).parent.parent.parent / "skills" / "word-text-operator"
if _scripts_parent.exists():
    sys.path.insert(0, str(_scripts_parent))

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/api", tags=["chat"])


# ── 请求/响应模型 ────────────────────────────────────────────────────

class ChatRequest(BaseModel):
    message: str
    selection_text: str = ""
    document_name: str = ""
    session_id: str = ""
    skill_name: str = ""


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
    stage: str = "execute"
    skill_selected: Optional[str] = None
    executed: List[Dict[str, Any]] = []


# ── 路由实现 ────────────────────────────────────────────────────────

@router.post("/chat", response_model=ChatResponse)
def chat(req: ChatRequest) -> ChatResponse:
    try:
        session_id = req.session_id.strip() or str(uuid.uuid4())
        history_before = len(llm_service.get_session_history(session_id))
        current_turn = (history_before // 2) + 1

        skill_explicit = (req.skill_name or "").strip()
        logger.info(
            "[/api/chat] 收到请求 | session_id=%s turn=%d direct_skill=%s msg_len=%d",
            session_id, current_turn, bool(skill_explicit), len(req.message or ""),
        )

        # 实时刷新选区
        current_selection = req.selection_text
        if not current_selection:
            try:
                current_selection = word_service.get_selection_text()
            except Exception:
                current_selection = ""

        skill_to_run = skill_explicit

        # ── Round 1：自动选技能 ────────────────────────────────────
        if not skill_to_run:
            skills_desc = get_skill_descriptions()
            prompt = _build_prompt_select(req.message, current_selection, skills_desc)
            llm_select = llm_service.chat_with_context(
                req.message, prompt, session_id=session_id,
            )
            skill_to_run = _parse_skill_selection(llm_select) or ""

            logger.info(
                "[/api/chat] 内部选技能完成 | session_id=%s skill=%s",
                session_id, skill_to_run or "(未识别)",
            )

            if not skill_to_run:
                return ChatResponse(
                    response="无法识别所需技能，请重述需求。",
                    success=True, session_id=session_id,
                    turn=current_turn, stage="select", skill_selected=None,
                )

        # ── Round 2：按技能说明书执行 ────────────────────────────────
        skill_content = get_skill_content(skill_to_run)
        if not skill_content:
            raise ValueError(f"未找到技能：{skill_to_run}（请检查 skills/ 目录下是否存在对应目录）")

        # 在调用 LLM 之前，先尝试连接 Word 捕获当前选中内容的格式状态
        # 这样 LLM 决策时就能知道初始状态（如当前字号、行距等）
        current_format_state = ""
        op_for_state: Any = None
        try:
            from scripts.word_text_operator import WordTextOperator
            op_for_state = WordTextOperator()
            op_for_state._base.connect()
            op_for_state._init_submodules()
            if op_for_state._base._word_app and op_for_state._fmt:
                sel = op_for_state._base._word_app.Selection
                if sel:
                    fi = op_for_state._fmt.get_font_info(sel)
                    pf = op_for_state._fmt.get_paragraph_format_info(sel)
                    state_parts = []
                    if fi.get("name") or fi.get("size"):
                        size_str = f"{fi['size']}pt" if fi.get("size") else "未设置"
                        name_str = fi.get("name", "")
                        bold_str = "粗体" if fi.get("bold") else ""
                        italic_str = "斜体" if fi.get("italic") else ""
                        styles = " ".join(x for x in [bold_str, italic_str] if x)
                        state_parts.append(f"字体: {name_str} {size_str}" + (f" ({styles})" if styles else ""))
                    align_map = {0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐", 4: "分散对齐"}
                    pf_parts = []
                    if pf.get("alignment") is not None:
                        pf_parts.append(f"对齐: {align_map.get(pf['alignment'], pf['alignment'])}")
                    if pf.get("line_spacing"):
                        pf_parts.append(f"行距: {pf['line_spacing']}")
                    if pf.get("first_line_indent"):
                        pf_parts.append(f"首行缩进: {pf['first_line_indent']}")
                    if pf_parts:
                        state_parts.append("段落: " + " | ".join(pf_parts))
                    current_format_state = "\n".join(state_parts) if state_parts else "(未检测到格式)"
                # 不断开连接，复用 op_for_state 执行 actions
        except Exception as e:
            logger.debug("[/api/chat] 捕获当前格式状态失败（不影响执行）: %s", e)
            current_format_state = ""

        prompt = _build_prompt_execute(
            req.message, current_selection, skill_to_run, skill_content,
            current_format_state,
        )
        llm_response = llm_service.chat_with_context(
            req.message, prompt, session_id=session_id,
        )

        actions = _parse_actions(llm_response)

        # 复用 Round 2 前已连接的 operator，避免重复连接
        op: Any = op_for_state
        init_error: str = ""
        if op is None:
            try:
                from scripts.word_text_operator import WordTextOperator
                op = WordTextOperator()
                op._base.connect()
                op._init_submodules()
            except Exception as e:
                init_error = str(e)
                logger.warning("[/api/chat] WordTextOperator 初始化失败: %s", e)

        if op is None or not getattr(op._base, "_word_app", None) or not getattr(op, "_fmt", None):
            return ChatResponse(
                response=f"无法连接 Word：{init_error or 'Word 未启动或没有打开文档'}",
                success=False,
                error=f"Word 连接失败: {init_error or 'Word 未启动或没有打开文档'}",
                session_id=session_id,
                turn=len(llm_service.get_session_history(session_id)) // 2,
                stage="execute",
                skill_selected=skill_to_run,
            )

        executed = []
        for action in actions:
            result = execute_action(action, op)
            executed.append(result)

        # 将执行结果写入历史，使下一轮 LLM 能感知到刚才的操作参数
        llm_service.update_executed_result(session_id, executed)

        summary = _summarize_execution(llm_response, executed)
        final_turn = len(llm_service.get_session_history(session_id)) // 2

        logger.info(
            "[/api/chat] 执行完成 | session_id=%s skill=%s actions=%d",
            session_id, skill_to_run, len(actions),
        )
        return ChatResponse(
            response=summary,
            success=True,
            session_id=session_id,
            turn=final_turn,
            stage="execute",
            skill_selected=skill_to_run,
            executed=executed,
        )

    except Exception as e:
        logger.exception("[/api/chat] 失败: %s", e)
        return ChatResponse(
            response=f"处理失败：{str(e)}",
            success=False, error=str(e),
            session_id=req.session_id or "", stage="execute",
        )


@router.get("/chat/history")
def chat_history(session_id: str) -> dict:
    if not session_id:
        raise HTTPException(status_code=400, detail="session_id 不能为空")
    history = llm_service.get_session_history(session_id)
    return {"session_id": session_id, "turns": history, "count": len(history)}


@router.delete("/chat/clear")
def chat_clear(session_id: str) -> dict:
    if not session_id:
        raise HTTPException(status_code=400, detail="session_id 不能为空")
    cleared = llm_service.clear_session(session_id)
    logger.info("[/api/chat/clear] session_id=%s cleared=%s", session_id, cleared)
    return {"session_id": session_id, "cleared": cleared}


@router.get("/word/status", response_model=WordStatusResponse)
def word_status() -> WordStatusResponse:
    try:
        connected = word_service.is_connected()
        has_sel = word_service.has_selection() if connected else False
        sel_text = word_service.get_selection_text() if has_sel else ""
        doc_name = word_service.get_document_name() if connected else None
        return WordStatusResponse(
            connected=connected, document_name=doc_name,
            has_selection=has_sel, selection_text=sel_text,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/word/connect")
def word_connect() -> dict:
    try:
        ok = word_service.connect(visible=True)
        return {"success": ok, "connected": ok}
    except Exception as e:
        return {"success": False, "error": str(e)}


@router.post("/word/disconnect")
def word_disconnect(save: bool = False) -> dict:
    try:
        word_service.quit(save_changes=save)
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


# ── 内部函数 ──────────────────────────────────────────────────────

def _build_prompt_select(user_message: str, selection_text: str, skills_desc: str) -> str:
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
如果不需要任何技能，返回 {{"skill": "", "reasoning": "不需要技能"}}。
如果判断需要某个技能，只返回以下 JSON 格式，不要包含其他文字：
```json
{{"skill": "技能目录名", "reasoning": "简短选择理由"}}
```"""


def _build_prompt_execute(
    user_message: str, selection_text: str, skill_name: str, skill_content: str,
    current_format_state: str = "",
) -> str:
    format_section = ""
    if current_format_state:
        format_section = f"""
## 当前选中内容的格式状态
（以下是你决策时的参考起始状态，用户如需"调回"/"撤销"/"改回去"，请以此为依据）
---
{current_format_state}
---

"""

    return f"""你是一个专业的 Word 文档格式化助手，正在使用技能「{skill_name}」。

## 当前上下文
用户选中了 Word 文档中的以下内容：
---
{selection_text if selection_text else "(无选中文本)"}
---
{format_section}
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

如果用户只是闲聊或问题咨询，不需要执行任何 Word 操作，请返回：[]"""


def _parse_skill_selection(llm_response: str) -> Optional[str]:
    try:
        match = re.search(r'\{[^}]*"skill"\s*:\s*"([^"]+)"[^}]*\}', llm_response)
        if match:
            skill = match.group(1).strip()
            if skill:
                valid_names = list_skill_names()
                if skill in valid_names:
                    return skill
                for name in valid_names:
                    if name.lower() == skill.lower() or skill.lower() in name.lower():
                        return name
        logger.warning("[_parse_skill_selection] 无法解析 skill：%s", llm_response[:200])
    except Exception as e:
        logger.warning("[_parse_skill_selection] 解析异常：%s，原始响应：%s", e, llm_response[:200])
    return None


def _parse_actions(llm_response: str) -> List[Dict[str, Any]]:
    try:
        match = re.search(r'\[.*\]', llm_response, re.DOTALL)
        if match:
            return json.loads(match.group())
    except (json.JSONDecodeError, re.error):
        pass
    return []


def _summarize_execution(llm_response: str, executed: List[Dict[str, Any]]) -> str:
    text = re.sub(r'\[.*\]', '', llm_response, flags=re.DOTALL).strip()
    text = re.sub(r'```json|```', '', text).strip()

    result_parts = []
    for e in executed:
        if e.get("success") and "result" in e:
            r = e["result"]
            if isinstance(r, str) and len(r) < 200:
                result_parts.append(f"「{r}」")
            elif isinstance(r, int):
                result_parts.append(f"结果：{r}")
            elif isinstance(r, list) and len(r) <= 5:
                result_parts.append(f"找到 {len(r)} 项")
            elif isinstance(r, dict) and "text" in r:
                result_parts.append(f"「{r['text']}」")

    if text:
        first_para = text.split('\n')[0].strip()
        if result_parts:
            return f"{first_para}。{' '.join(result_parts)}"
        return first_para

    if executed:
        ok = [e.get("description", e.get("action", "?")) for e in executed if e.get("success")]
        failed = [e.get("action") for e in executed if not e.get("success")]
        parts = []
        if ok:
            parts.append(f"✅ 已执行：{', '.join(ok)}")
        if failed:
            parts.append(f"❌ 失败：{', '.join(failed)}")
        if result_parts:
            parts.append(" ".join(result_parts))
        return " ".join(parts) if parts else "已完成处理。"

    return "已完成处理。"
