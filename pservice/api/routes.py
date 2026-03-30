# -*- coding: utf-8 -*-
"""
AI 聊天路由

流程（每次 HTTP 请求）：
  1. 前端传 session_id → 后端按 UUID 隔离多轮对话历史
  2. 若 skill 未指定：内部 Round 1（LLM 选技能），不追加到历史
  3. 立即连接 Word，捕获当前格式状态（字体 + 段落）
  4. 单次 LLM 调用（带当前状态 + 对话历史），追加到历史
  5. 执行 actions（execute_action 捕获操作前的 before_state）
  6. 将执行结果追加到历史记录的 executed 字段
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
        # 前端生成 UUID，每次请求都带过来；未带时自动生成（兼容旧前端）
        session_id = req.session_id.strip() or str(uuid.uuid4())
        history_before = len(llm_service.get_session_history(session_id))
        current_turn = (history_before // 2) + 1

        skill_explicit = (req.skill_name or "").strip()
        logger.info(
            "[/api/chat] 收到请求 | session_id=%s turn=%d direct_skill=%s msg_len=%d",
            session_id, current_turn, bool(skill_explicit), len(req.message or ""),
        )

        # ── 实时刷新选区 ──────────────────────────────────────────────
        current_selection = req.selection_text
        if not current_selection:
            try:
                current_selection = word_service.get_selection_text()
            except Exception:
                current_selection = ""

        skill_to_run = skill_explicit

        # ── Round 1：内部选技能（skill 未指定时）─────────────────────────
        # 内部决策，不追加到历史，避免污染用户的对话历史
        if not skill_to_run:
            skills_desc = get_skill_descriptions()
            prompt = _build_prompt_select(req.message, current_selection, skills_desc)
            llm_select = llm_service.chat_with_context(
                req.message, prompt, session_id=session_id,
                add_to_history=False,
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

        skill_content = get_skill_content(skill_to_run)
        if not skill_content:
            raise ValueError(
                f"未找到技能：{skill_to_run}（请检查 skills/ 目录下是否存在对应目录）"
            )

        # ── 立即连接 Word，捕获当前格式状态 ─────────────────────────────
        # 关键：在 LLM 调用之前就捕获状态。
        # 这样无论是第 1 轮还是第 N 轮，LLM 都能看到"当前字体/段落是什么"。
        # 注意：每次请求都会重新连接，以反映用户最新的选区状态。
        op: Any = None
        init_error: str = ""
        current_font_state: str = ""
        current_para_state: str = ""
        try:
            from scripts.word_text_operator import WordTextOperator
            op = WordTextOperator()
            op._base.connect()
            op._init_submodules()
            if op._base._word_app and op._fmt:
                sel = op._base._word_app.Selection
                # 有选区时读选区格式；仅光标（折叠选区）时读插入点处格式（Word 仍返回有效 Font）
                if sel is not None:
                    try:
                        has_range = sel.Start != sel.End
                    except Exception as e:
                        logger.warning("[/api/chat] sel.Start/End 异常（可能sel非Range对象）: %s", e)
                        has_range = False
                    # 字体状态
                    try:
                        fi = op._fmt.get_font_info(sel)
                    except Exception as e:
                        logger.error("[/api/chat] get_font_info 异常: %s", e)
                        fi = {}
                    logger.info(
                        "[/api/chat] 原始 font_info: name=%r size=%r bold=%r italic=%r underline=%r",
                        fi.get("name"), fi.get("size"), fi.get("bold"),
                        fi.get("italic"), fi.get("underline"),
                    )
                    # 只要 fi 不是全空 dict，就说明 COM 通信正常，只是部分字段无值
                    fi_has_data = any(
                        fi.get(k) not in (None, "", 0) for k in ("name", "size", "bold", "italic", "underline")
                    )
                    if fi is None or not isinstance(fi, dict):
                        logger.warning("[/api/chat] font_info 返回异常: %r", fi)
                    elif fi_has_data or (fi.get("name") is not None) or (fi.get("size") is not None):
                        size_val = fi.get("size")
                        name_val = fi.get("name")
                        size_str = f"{size_val}pt" if size_val and size_val != 0 else "字号未知"
                        name_str = name_val if name_val else "字体名未知"
                        bold_str = "粗体" if fi.get("bold") in (-1, True) else ""
                        italic_str = "斜体" if fi.get("italic") in (-1, True) else ""
                        underline_val = fi.get("underline", 0)
                        underline_str = "下划线" if underline_val and underline_val != 0 else ""
                        styles = " ".join(x for x in [bold_str, italic_str, underline_str] if x)
                        current_font_state = (
                            f"字体: {name_str} {size_str}"
                            + (f" ({styles})" if styles else " (常规)")
                        )
                        if not has_range:
                            current_font_state += "（插入点/折叠选区，以上为光标处格式）"
                    else:
                        logger.warning(
                            "[/api/chat] font_info 全字段为空，将使用占位状态: %r", fi,
                        )
                    # 段落状态
                    try:
                        pf = op._fmt.get_paragraph_format_info(sel)
                    except Exception as e:
                        logger.error("[/api/chat] get_paragraph_format_info 异常: %s", e)
                        pf = {}
                    logger.info(
                        "[/api/chat] 原始 para_info: alignment=%r line_spacing=%r first_line_indent=%r",
                        pf.get("alignment"), pf.get("line_spacing"), pf.get("first_line_indent"),
                    )
                    align_map = {0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐", 4: "分散对齐"}
                    pf_parts = []
                    if pf.get("alignment") is not None:
                        pf_parts.append(f"对齐: {align_map.get(pf['alignment'], pf['alignment'])}")
                    if pf.get("line_spacing") not in (None, 0, ""):
                        pf_parts.append(f"行距: {pf['line_spacing']}")
                    if pf.get("first_line_indent") not in (None, 0, ""):
                        pf_parts.append(f"首行缩进: {pf['first_line_indent']}")
                    if pf_parts:
                        current_para_state = "段落: " + " | ".join(pf_parts)
        except Exception as e:
            init_error = str(e)
            logger.warning("[/api/chat] WordTextOperator 初始化失败: %s", e)

        if op is None or not getattr(op._base, "_word_app", None) or not getattr(op, "_fmt", None):
            return ChatResponse(
                response=f"无法连接 Word：{init_error or 'Word 未启动或没有打开文档'}",
                success=False,
                error=f"Word 连接失败: {init_error or 'Word 未启动或没有打开文档'}",
                session_id=session_id,
                turn=current_turn,
                stage="execute",
                skill_selected=skill_to_run,
            )

        # ── 诊断：Word 连接状态 ───────────────────────────────────────────
        logger.info(
            "[/api/chat] Word连接诊断 | op=%s word_app=%s fmt=%s",
            op is not None,
            getattr(op, "_base", None) is not None and getattr(op._base, "_word_app", None) is not None,
            getattr(op, "_fmt", None) is not None,
        )
        # ── 单次 LLM 调用（带当前状态）────────────────────────────────────
        # 关键：在 LLM 调用时就附上当前格式状态。
        # 同时把状态也拼到 user_message 里，这样在日志和历史中能原样看到【当前状态】。
        # state_parts 已由上方捕获逻辑填充。
        state_parts = [p for p in [current_font_state, current_para_state] if p]
        current_format_state = "\n".join(state_parts)

        logger.info(
            "[/api/chat] 状态捕获 | font='%s' para='%s'",
            current_font_state[:40], current_para_state[:40],
        )

        # ── 构造发给 LLM 的 user_message（含当前状态）──────────────────────
        # 固定出现【用户要求】【当前状态】，便于在日志/终端里搜索「当前状态」；
        # 无捕获数据时用占位，避免整段省略导致搜不到关键字。
        state_block = (
            current_format_state.strip()
            if current_format_state.strip()
            else "（未读取到有效格式：Word 未就绪、选区异常，或 COM 未返回字体名/字号等）"
        )
        user_msg_with_state = (
            f"【用户要求】\n{req.message}\n\n【当前状态】\n{state_block}"
        )

        prompt = _build_prompt_execute(
            req.message, current_selection, skill_to_run, skill_content,
            current_format_state,
        )
        llm_response = llm_service.chat_with_context(
            user_msg_with_state, prompt, session_id=session_id,
        )

        actions = _parse_actions(llm_response)
        if not actions:
            return ChatResponse(
                response="LLM 未返回有效操作。",
                success=True, session_id=session_id,
                turn=current_turn, stage="execute", skill_selected=skill_to_run,
            )

        # ── 执行 action ────────────────────────────────────────────────
        # execute_action 内部会捕获 before_state（操作执行前的格式快照），
        # append_executed_result 将其追加到历史记录中，
        # 使前端能看到"从 X 变为 Y"的对比。
        executed = []
        for action in actions:
            result = execute_action(action, op)
            executed.append(result)

        llm_service.append_executed_result(session_id, executed)

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
    state_for_prompt = (
        current_format_state.strip()
        if current_format_state.strip()
        else "（未读取到有效格式，以用户消息中【当前状态】块为准）"
    )
    format_section = f"""
## 当前选中内容的格式状态
（以下是你决策时的参考起始状态，用户如需"调回"/"撤销"/"改回去"，请以此为依据）
---
{state_for_prompt}
---

"""

    return f"""你是一个专业的 Word 文档格式化助手，正在使用技能「{skill_name}」。

## 当前上下文
用户选中了 Word 文档中的以下内容：
---
{selection_text if selection_text else "(无选中文本)"}
---
{format_section}## 用户的需求
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
