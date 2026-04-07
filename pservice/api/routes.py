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

# ── 加载 word-text-operator skill 模块 ────────────────────────────────────────
# 策略与 main.py 一致：用 importlib.util 逐文件加载，解决相对导入问题
_project_root = Path(__file__).parent.parent.parent


def _ensure_wto_module():
    """确保 WordTextOperator 已加载到 sys.modules。"""
    key = "scripts.word_text_operator"
    if key in sys.modules:
        return
    scripts_dir = _project_root / "skills" / "word-text-operator" / "scripts"
    _scripts_parent = Path(__file__).parent.parent.parent / "skills" / "word-text-operator"
    if str(_scripts_parent) not in sys.path:
        sys.path.insert(0, str(_scripts_parent))

    submodules = ["word_base", "word_range_navigation", "word_text_operations",
                  "word_selection", "word_find_replace", "word_format", "word_bookmark"]
    import importlib.util

    for sub in submodules:
        sub_file = scripts_dir / f"{sub}.py"
        if not sub_file.exists():
            continue
        sub_name = f"scripts.{sub}"
        if sub_name in sys.modules:
            continue
        spec = importlib.util.spec_from_file_location(sub_name, sub_file)
        if spec is None:
            continue
        mod = importlib.util.module_from_spec(spec)
        sys.modules[sub_name] = mod
        try:
            spec.loader.exec_module(mod)
        except Exception:
            pass

    main_file = scripts_dir / "word_text_operator.py"
    if main_file.exists():
        spec = importlib.util.spec_from_file_location(key, main_file)
        if spec is not None:
            mod = importlib.util.module_from_spec(spec)
            sys.modules[key] = mod
            try:
                spec.loader.exec_module(mod)
            except Exception:
                pass


_ensure_wto_module()
from scripts.word_text_operator import WordTextOperator


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

        logger.info("[/api/chat] >>> Step 1: 检查 Word 连接")
        # ── 每次请求都检查 Word 连接状态，未连接时尝试重新连接 ───────────────
        if not word_service.is_connected():
            logger.info("[/api/chat] Word 未连接，尝试重新连接...")
            reconnect_ok = word_service.connect(visible=True)
            if not reconnect_ok:
                logger.warning("[/api/chat] Word 重新连接失败")

        logger.info("[/api/chat] >>> Step 2: 读取选区")
        # ── 实时刷新选区 ──────────────────────────────────────────────
        current_selection = req.selection_text
        if not current_selection:
            try:
                current_selection = word_service.get_selection_text()
            except Exception:
                current_selection = ""

        logger.info("[/api/chat] >>> Step 3: 初始化 WordTextOperator")
        # ── Step 1: 初始化 WordTextOperator（复用 word_service 的 COM 连接） ──
        try:
            op = WordTextOperator()
            op._base._word_app = word_service._word_app
            # word_service._document 仅在通过 service 打开文档时存在；
            # 若 Word 由用户手动打开（无 service 介入），则为 None。
            # 两种情况都尝试以 ActiveDocument 作为兜底。
            if word_service._document is None:
                try:
                    op._base._document = word_service._word_app.ActiveDocument
                except Exception:
                    op._base._document = None
            else:
                op._base._document = word_service._document
            op._init_submodules()
        except Exception as e:
            logger.warning("[/api/chat] WordTextOperator init failed: %s", e)
            op = None

        # ── Step 2: 捕获当前格式状态（供 Plan/Execute 阶段 LLM 参考） ──────────
        current_format_state = ""
        if op and op._base._word_app is not None:
            try:
                rng = op._base.selection
                font_info = op._fmt.get_font_info(rng) if op._fmt else {}
                para_info = op._fmt.get_paragraph_format_info(rng) if op._fmt else {}
                lines = []
                if font_info:
                    lines.append("字体：")
                    for k, v in font_info.items():
                        if v and v != 0:
                            lines.append(f"  {k}: {v}")
                if para_info:
                    lines.append("段落：")
                    for k, v in para_info.items():
                        if v and v != 0:
                            lines.append(f"  {k}: {v}")
                current_format_state = "\n".join(lines)
            except Exception as e:
                logger.warning(
                    "[/api/chat] 读取格式状态失败（可能是 Word 已关闭或 COM 连接失效）: %s", e
                )

        logger.info("[/api/chat] >>> Step 4: Plan 阶段 - 调用 LLM 决定技能和拆分")
        # ── Step 3: Plan 阶段 - 决定每个 step 的技能和拆分 ────────────────────
        skills_desc = get_skill_descriptions()
        prompt_plan = _build_prompt_plan_select(
            req.message, current_selection, skills_desc,
            current_format_state=current_format_state,
        )
        logger.info("[/api/chat] >>> Step 4: 实际调用 LLM（Plan），即将阻塞等待响应...")
        llm_plan_resp = llm_service.chat_with_context(
            req.message, prompt_plan, session_id=session_id,
            add_to_history=False,
        )
        logger.info("[/api/chat] >>> Step 4: LLM 响应已返回，长度=%d", len(llm_plan_resp))
        plan_steps = _parse_plan_select(llm_plan_resp)
        if not plan_steps:
            # plan_steps 为 None 或空列表时走 fallback
            logger.warning(
                "[plan] 无法解析 LLM plan 响应，回退到单步执行 | resp_preview=%s",
                llm_plan_resp[:200],
            )
            skill_to_run = skill_explicit
            if not skill_to_run:
                skills_all = get_skill_descriptions()
                prompt_sel = _build_prompt_select(req.message, current_selection, skills_all)
                llm_sel = llm_service.chat_with_context(
                    req.message, prompt_sel, session_id=session_id,
                    add_to_history=False,
                )
                skill_to_run = _parse_skill_selection(llm_sel) or ""
            if not skill_to_run:
                return ChatResponse(
                    response="无法理解您的需求，请重述。",
                    success=True, session_id=session_id,
                    turn=current_turn, stage="plan",
                )
            plan_steps = [{"step": 1, "skill": skill_to_run, "description": "单步回退",
                          "selection_hint": "全文", "need_feedback": False}]

        # ── Step 4: 初始化 ParagraphOperator（用于段落索引操作） ────────────────
        para_op = None
        if op and op._base._word_app is not None:
            try:
                from scripts.word_paragraph_operator import ParagraphOperator
                para_op = ParagraphOperator(op._base)
            except Exception as e:
                logger.warning("[/api/chat] ParagraphOperator init failed: %s", e)

        # ── Step 5: 循环执行每个 step ────────────────────────────────────────
        executed_all = []
        prev_results = []
        prev_step_feedback = []   # 上一步的 feedback，供下一步 prompt 使用
        for step_def in plan_steps:
            step_num = step_def["step"]
            skill_name = step_def["skill"]
            step_desc = step_def["description"]
            step_selection_hint = step_def.get("selection_hint", "")

            skill_content = get_skill_content(skill_name)
            if not skill_content:
                logger.error("[plan] skill content empty: %s", skill_name)
                executed_all.append({"step": step_num, "skill": skill_name, "action": "",
                                    "success": False, "error": "skill %s not found" % skill_name})
                continue

            logger.info(
                "[plan] Step %d/%d | skill=%s | hint=%s | desc=%s",
                step_num, len(plan_steps), skill_name, step_selection_hint, step_desc,
            )

            prompt_exec = _build_prompt_execute(
                req.message, current_selection, skill_name, skill_content,
                current_format_state,
                selection_hint=step_selection_hint,
                prev_executed=prev_results,
                step_num=step_num,
                total_steps=len(plan_steps),
                step_def=step_def,
                prev_step_feedback=prev_step_feedback,
            )
            logger.info("[/api/chat] >>> Step 5: 执行循环 step=%d skill=%s，即将调用 LLM...", step_num, skill_name)
            try:
                llm_resp = llm_service.chat_with_context(
                    req.message, prompt_exec, session_id=session_id,
                    add_to_history=False,
                )
                logger.info("[/api/chat] >>> Step 5: step=%d LLM 响应已返回，长度=%d", step_num, len(llm_resp))
            except Exception as e:
                logger.error("[plan] Step %d LLM failed: %s", step_num, e)
                executed_all.append({"step": step_num, "skill": skill_name, "action": "",
                                    "success": False, "error": "LLM failed: %s" % e})
                continue

            actions = _parse_actions(llm_resp)
            if not actions:
                logger.warning("[plan] Step %d no valid actions", step_num)
                executed_all.append({"step": step_num, "skill": skill_name, "action": "",
                                    "success": False, "error": "no valid actions"})
                continue

            logger.info("[plan] Step %d actions: %s", step_num, [a.get("action") for a in actions])

            step_prev = list(prev_results)
            for action in actions:
                expanded = _substitute_rng_placeholder([action], step_prev)
                for sub_action in expanded:
                    result = execute_action(sub_action, op, para_op=para_op)
                    result["step"] = step_num
                    result["skill"] = skill_name
                    executed_all.append(result)
                    step_prev.append(result)
                    if result.get("success"):
                        logger.info(
                            "[plan]   OK step=%d action=%s",
                            step_num, sub_action.get("action"),
                        )
                    else:
                        logger.warning(
                            "[plan]   FAIL step=%d action=%s error=%s",
                            step_num,
                            sub_action.get("action"),
                            result.get("error"),
                        )

            prev_results = step_prev
            prev_step_feedback = step_prev   # 本步结果作为下一步的 feedback

        llm_service.append_executed_result(session_id, executed_all)
        summary = _summarize_execution("", executed_all)
        final_turn = len(llm_service.get_session_history(session_id)) // 2

        logger.info(
            "[/api/chat] done | session=%s steps=%d actions=%d ok=%d",
            session_id, len(plan_steps), len(executed_all),
            sum(1 for e in executed_all if e.get("success")),
        )
        return ChatResponse(
            response=summary,
            success=True,
            session_id=session_id,
            turn=final_turn,
            stage="execute",
            executed=executed_all,
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


# ── Plan 阶段 ──────────────────────────────────────────────────────────

def _build_prompt_plan_select(
    user_message: str, selection_text: str,
    all_skills_desc: str,
    current_format_state: str = "",
) -> str:
    """Plan 阶段：LLM 决定每步的技能和拆分意图（不声明选区范围）。"""
    selection_block = (
        f"用户选中了以下内容：\n---\n{selection_text}\n---\n\n"
        if selection_text
        else "（用户当前无选中文本）\n\n"
    )
    state_section = (
        f"## 当前选中内容的格式状态\n---\n{current_format_state.strip()}\n---\n\n"
        if current_format_state.strip()
        else ""
    )
    return f"""你是一个 Word 文档操作规划助手。

{selection_block}{state_section}## 用户的需求
"{user_message}"

## 可用技能列表
{all_skills_desc}

## 你的任务
分析用户需求，决定需要几步操作以及每步分别使用哪个技能。

**重要规则**：
- 每个 step 只能选一个技能
- 不同选区的操作必须拆为不同 step（如"第1段"和"选中部分"是不同选区，“第1行”和“选中”部分是不同选区）
- **在仅涉及一个选区的情况下，纯修改类需求（缩进、行距、字体等）必须控制为1步，绝不能拆成多个 step**（例如：同时改缩进和行距 → 1步；设字体和字号 → 1步；改对齐和首行缩进 → 1步）
- 先查后改的需求（find 后修改）→ 2 步（第一步 need_feedback=true）

**违反规则的典型错误示例**：
- ❌ 用户说"把选中文字的首行缩进改成2字符" → 错误拆成"获取段落"+"设置缩进"两个 step
- ❌ 用户说"调大行距" → 错误拆成"查格式"+"改行距"两个 step
- ❌ 用户说"把字体设成黑体，字号设成三号" → 错误拆成两个 step
- ✅ 正确：所有纯修改类操作在同一 step 内一次性完成

## 输出格式

```json
{{
  "plan": [
    {{"step": 1, "skill": "word-text-operator", "description": "...", "selection_hint": "第1段", "need_feedback": false}},
    {{"step": 2, "skill": "word-text-operator", "description": "...", "selection_hint": "用户当前选中的文本", "need_feedback": false}}
  ]
}}
```

**selection_hint**：用简短的自然语言描述该 step 的操作对象（如"第1段"、"全文"、"选中部分"）。
**need_feedback**：仅在"先查后改"场景设为 true，其余均为 false。

只返回 JSON。"""


def _parse_plan_select(llm_response: str):
    """解析 plan JSON，提取 skill 和 selection_hint。"""
    try:
        import re as _re
        m = _re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", llm_response, _re.IGNORECASE)
        text = m.group(1).strip() if m else llm_response.strip()
        for start in range(len(text)):
            ch = text[start]
            if ch == "{":
                obj = json.JSONDecoder().raw_decode(text, start)[0]
                plan = obj.get("plan") or obj.get("steps")
                if not isinstance(plan, list):
                    return None
                break
            if ch == "[":
                plan = json.JSONDecoder().raw_decode(text, start)[0]
                if not isinstance(plan, list):
                    return None
                break
        else:
            return None

        valid_names = list_skill_names()
        validated = []
        for step in plan:
            if not isinstance(step, dict):
                continue
            raw_skill = (step.get("skill") or "").strip()
            skill_name = None
            for name in valid_names:
                if name.lower() == raw_skill.lower() or raw_skill.lower() in name.lower():
                    skill_name = name
                    break
            if not skill_name:
                continue
            validated.append({
                "step": int(step.get("step", len(validated) + 1)),
                "skill": skill_name,
                "description": step.get("description", ""),
                "selection_hint": step.get("selection_hint", ""),
                "need_feedback": bool(step.get("need_feedback", False)),
            })
        return validated if validated else None
    except Exception as e:
        import traceback
        traceback.print_exc()
        logger.warning("[_parse_plan_select] 解析异常：%s，原始：%s", e, llm_response[:200])
        return None


def _substitute_rng_placeholder(
    actions, prev_results,
):
    """
    1. 从 prev_results 中提取选区：
       - 单个选区：get_paragraph_range → [s, e]；get_selection_info → {start, end}
       - 多个选区：find_all → [{start, end, text}, ...]
    2. 将后续 action 的 rng 占位符 "[start, end]" 替换为真实值
    3. 若 action 缺少 rng 参数但有查询结果，直接注入
    4. find_all 返回多个位置时，自动将单个 action 展开为多个（每个位置一条）
    """
    # ── 收集所有选区 ───────────────────────────────────────────────
    single_range = None   # 单个 [start, end]
    multiple_ranges = []   # find_all 返回的多位置列表

    for r in reversed(list(prev_results)):
        if not r.get("success"):
            continue
        res = r.get("result")
        action_name = r.get("action", "")
        if action_name == "find_all" and isinstance(res, list):
            # 收集全部位置（按 start,end 去重，防止上游重复 yield）
            seen_pos = set()
            for item in res:
                if isinstance(item, dict) and "start" in item and "end" in item:
                    key = (int(item["start"]), int(item["end"]))
                    if key not in seen_pos:
                        seen_pos.add(key)
                        multiple_ranges.append([key[0], key[1]])
            break
        elif isinstance(res, (list, tuple)) and len(res) == 2:
            single_range = list(res)
            break
        elif isinstance(res, dict) and "start" in res and "end" in res:
            single_range = [int(res["start"]), int(res["end"])]
            break

    # 有多个位置 → 展开所有 action，每个位置执行一次（查询类除外）
    if multiple_ranges:
        rng_literal = "[start, end]"
        _QUERY_ACTIONS = frozenset((
            "get_selection_info", "get_full_text", "get_text",
            "get_paragraph_range", "find", "find_all", "count_occurrences",
            "list_bookmarks", "char_count", "word_count",
            "sentence_count", "paragraph_count",
        ))
        expanded = []
        for a in actions:
            params = dict(a.get("params", {}))
            rng_val = params.get("rng")
            aname = a.get("action", "")
            # 查询类 action 只执行一次（取当前位置），其余 action 每个位置展开一条
            if aname in _QUERY_ACTIONS:
                logger.info("[rng_skip] 查询类 action 不展开 | action=%s", aname)
                resolved_a = {**a}
                if rng_val == rng_literal:
                    resolved_a = {**a, "params": {**params, "rng": multiple_ranges[0]}}
                    logger.info("[rng_substitute] 查询类 [start,end] → %s | action=%s", multiple_ranges[0], aname)
                expanded.append(resolved_a)
                continue
            if rng_val == rng_literal:
                logger.warning(
                    "[rng_substitute] find_all 多位置场景下不支持 [start, end] 占位符，"
                    "请省略 rng 参数 | action=%s", aname,
                )
            if rng_val is None or rng_val == rng_literal:
                for pos in multiple_ranges:
                    p2 = dict(params)
                    p2.pop("rng", None)
                    p2["rng"] = pos
                    expanded.append({**a, "params": p2})
                    logger.info(
                        "[rng_expand] find_all 展开 | action=%s → rng=%s",
                        aname, pos,
                    )
            else:
                expanded.append(a)
        return expanded

    if single_range is None:
        return actions

    rng_literal = "[start, end]"
    resolved = []
    for a in actions:
        params = dict(a.get("params", {}))
        rng_val = params.get("rng")
        action_name = a.get("action", "")
        if rng_val == rng_literal:
            params["rng"] = single_range
            logger.info("[rng_substitute] 替换占位符 → %s | action=%s", single_range, action_name)
            resolved.append({**a, "params": params})
        elif rng_val is None and any(
            action_name.startswith(p) for p in (
                "set_", "replace", "expand", "collapse",
                "delete_range", "clear_range", "insert_",
                "get_", "select_paragraph", "select_paragraph_range",
                "merge_with", "split_paragraph",
            )
        ):
            params["rng"] = single_range
            logger.info("[rng_inject] 注入 rng → %s | action=%s", single_range, action_name)
            resolved.append({**a, "params": params})
        else:
            resolved.append(a)
    return resolved


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
    selection_hint: str = "",
    prev_executed: list = None,
    step_num: int = 1,
    total_steps: int = 1,
    step_def: dict = None,
    prev_step_feedback: list = None,
) -> str:
    prev_executed = prev_executed or []
    step_def = step_def or {}
    prev_step_feedback = prev_step_feedback or []

    sf = current_format_state.strip() if current_format_state.strip() else "（未读取到有效格式，以用户消息中【当前状态】块为准）"
    format_section = (
        "\n\n## 当前选中内容的格式状态\n（以下是你决策时的参考起始状态）\n---\n"
        + sf + "\n---\n"
    )

    # ── 整理上一步 feedback 中的所有选区 ─────────────────────────────────
    # 注意：选区查询结果只属于当前 step，跨 step 必须重新查询。
    # 此处收集的 all_ranges 仅用于「发现 find_all 返回了多位置」时告诉 LLM 该情况，
    # 绝不意味着后续 step 可以复用这些值。
    all_ranges = []
    if prev_step_feedback:
        for r in prev_step_feedback:
            if not r.get("success"):
                continue
            res = r.get("result")
            if isinstance(res, dict) and "matches" in res:
                all_ranges.extend(res["matches"])
            elif r.get("action") == "find_all" and isinstance(res, list):
                seen_fb = set()
                for item in res:
                    if isinstance(item, dict) and "start" in item and "end" in item:
                        key = (int(item["start"]), int(item["end"]))
                        if key not in seen_fb:
                            seen_fb.add(key)
                            all_ranges.append(dict(item))

    feedback_section = ""
    if prev_step_feedback:
        lines_fb = ["\n## [上一步 feedback]"]
        for r in prev_step_feedback:
            a = r.get("action", "?")
            ok = "OK" if r.get("success") else "FAIL"
            res = r.get("result")
            info = a
            if res and isinstance(res, dict):
                if "start" in res and "end" in res:
                    info += " → 选区: %d~%d" % (res["start"], res["end"])
                elif "matches" in res:
                    info += " → 查找到 %d 处" % len(res["matches"])
            elif res and isinstance(res, list) and a == "find_all":
                uniq = []
                s2 = set()
                for it in res:
                    if isinstance(it, dict) and "start" in it and "end" in it:
                        k = (int(it["start"]), int(it["end"]))
                        if k not in s2:
                            s2.add(k)
                            uniq.append(it)
                info += " → 查找到 %d 处" % len(uniq)
            lines_fb.append("  [" + ok + "] " + info)
        if all_ranges:
            rng_lines = ["  选区列表（共 %d 个）：" % len(all_ranges)]
            for i, rng in enumerate(all_ranges, 1):
                text = rng.get("text", "")
                if text:
                    rng_lines.append("    [%d] start=%d, end=%d, text=\"%s\"" % (i, rng["start"], rng["end"], text))
                else:
                    rng_lines.append("    [%d] start=%d, end=%d" % (i, rng["start"], rng["end"]))
            lines_fb.extend(rng_lines)
        feedback_section = "\n".join(lines_fb) + "\n"

    hint_section = ""
    if step_def:
        if prev_step_feedback and all_ranges:
            count = len(all_ranges)
            requirement_line = (
                "**强制要求**：如果已有上一步 feedback，上方选区列表中共有 %d 个选区需要处理。\n"
                "本 step 只需返回操作类 action（如 set_bold、set_font_color、replace 等），\n"
                "**禁止**返回任何选区查询类 action（find_all、get_selection_info、get_paragraph_range 等）。\n"
                "系统会将操作类 action 按每个选区自动展开执行（共 %d 条），无需 LLM 生成循环。\n"
                "rng 参数使用上方选区列表中的具体 start/end 值，格式示例：\n"
                '  {"action": "set_bold", "params": {"rng": [%d, %d], "bold": true}, "description": "设置加粗"}\n'
            ) % (count, count, all_ranges[0]["start"], all_ranges[0]["end"])
        elif prev_step_feedback:
            requirement_line = (
                "**强制要求**：\n"
                "  - 选区查询结果（如 get_paragraph_range、get_selection_info 的返回值）**只属于当前 step 内**的后续 action，\n"
                "    **不可跨 step 使用**。其他 step 的选区结果与本 step 无关。\n"
                "  - 本 step 仍需先查询选区（除非操作「全文」），然后再执行操作：\n"
                '    · 操作「第N段」-> get_paragraph_range(index=N) （1-based）\n'
                '    · 操作「选中部分」-> get_selection_info()\n'
                '    · 操作「全文」-> 不需要 rng\n'
                "  - 后续格式类 action 的 rng 不要填写，系统会自动用本 step 查询结果填充。"
            )
        else:
            requirement_line = (
                "**强制要求**：\n"
                "  - 若当前 step 是纯查询类 action（find_all、get_xxx 等），则只需返回该查询 action，不要混入操作类 action。\n"
                "  - 若当前 step 不是纯查询类 action，则第一步必须调用选区查询 action：\n"
                '    · 操作「第N段」-> get_paragraph_range(index=N) （1-based）\n'
                '    · 操作「选中部分」-> get_selection_info()\n'
                '    · 操作「全文」-> 不需要 rng\n'
                "  - 后续格式类 action（set_font_name 等）的 rng 不要填写，系统会自动填充。"
            )
        hint_section = (
            feedback_section
            + "## [PLAN] 当前 step（第 %d/%d 步）\n"
            "本 step 定义：%s\n"
            + requirement_line
            + "\n"
        ) % (step_num, total_steps, str(step_def))

    pe = ""
    if prev_executed:
        lines_pe = ["\n## [已完成] 前面 step 执行结果（本轮以前）："]
        for r in prev_executed:
            a = r.get("action", "?")
            ok = "OK" if r.get("success") else "FAIL"
            desc = r.get("description", "")
            rng = r.get("result")
            info = a
            if desc:
                info += " - " + desc
            if rng and isinstance(rng, (list, tuple)):
                info += " | rng=" + str(list(rng))
            elif rng and isinstance(rng, dict):
                info += " | sel=" + str(rng.get("start")) + "~" + str(rng.get("end"))
            lines_pe.append("  [" + ok + "] " + info)
        pe = "\n".join(lines_pe) + "\n"
    else:
        pe = "\n## [已完成] 前面 step 执行结果（本轮以前）：(暂空)\n"

    st = selection_text if selection_text else "(无选中文本)"
    if prev_step_feedback and all_ranges:
        count = len(all_ranges)
        task_step2_line = (
            "2. 你只需要规划当前 step（第 %d/%d 步）的 action，不要规划其他 step 的内容。\n"
            "3. 如果已有上一步 feedback，上方选区列表中共有 %d 个选区需要处理。\n"
            "   必须只返回操作类 action，系统自动按每个选区展开执行（共 %d 条），禁止返回查询类 action。\n"
            "4. 所有 action 的 rng 参数使用选区列表中的具体 start/end 值，不要留空。"
        ) % (step_num, total_steps, count, count)
        rng_example = (
            '  {"action": "set_bold", "params": {"rng": [%d, %d], "bold": true}, "description": "设置加粗"},'
            "\n  ..."
        ) % (all_ranges[0]["start"], all_ranges[0]["end"])
    elif prev_step_feedback:
        task_step2_line = (
            "2. 你只需要规划当前 step（第 %d/%d 步）的 action，不要规划其他 step 的内容。\n"
            "3. 注意：选区查询结果（如 get_paragraph_range、get_selection_info）**只属于当前 step**。\n"
            "   其他 step 的选区结果与本 step 无关，本 step 必须自行查询（除非操作全文）。\n"
            "4. 后续格式类 action 的 rng 不要填写，系统会自动用本 step 的查询结果填充。"
        ) % (step_num, total_steps)
        rng_example = (
            '  {"action": "get_paragraph_range", "params": {"index": 1}, "description": "获取第1段范围"}\n'
            '  {"action": "set_font_name", "params": {"font_name": "黑体"}, "description": "设为黑体"}\n'
            '  {"action": "set_font_size", "params": {"size": 12.0}, "description": "设为小四"}'
        )
    else:
        task_step2_line = (
            "2. 你只需要规划当前 step（第 %d/%d 步）的 action，不要规划其他 step 的内容。\n"
            "3. 若当前 step 不是纯查询类 action（如 find_all、get_xxx），则第一步必须调用选区查询 action。\n"
            "4. 后续格式类 action 的 rng 不要填写，系统会自动用第一步的查询结果填充。"
        ) % (step_num, total_steps)
        # step1 纯查询 → 示例改为 find_all，不应出现 set_xxx
        rng_example = (
            '  {"action": "find_all", "params": {"text": "关键词"}, "description": "全文查找关键词"}\n'
            '  {"action": "get_selection_info", "params": {}, "description": "获取当前选区信息"}'
        )
    task_block = (
        "\n## 你的任务\n"
        "1. 仔细阅读上方技能说明书\n"
        + task_step2_line + "\n"
        "\n## 注意事项\n"
        "1. **强制要求**：\n"
    )
    if prev_step_feedback and all_ranges:
        count = len(all_ranges)
        requirement_note = (
            "  - 如果已有上一步 feedback，本 step 共有 %d 个选区需要处理。\n"
            "    本 step **只允许**返回操作类 action（如 set_bold、set_font_color、replace 等），\n"
            "    **禁止**返回任何选区查询类 action（find_all、get_selection_info、get_paragraph_range 等）。\n"
            "    系统自动将操作类 action 按每个选区展开执行，无需 LLM 生成循环。\n"
            "  - rng 参数使用上方选区列表中的具体 start/end 值，格式示例：\n"
            '    {"action": "set_bold", "params": {"rng": [%d, %d], "bold": true}, "description": "设置加粗"}\n'
            "    不允许留空或使用占位符。"
        ) % (count, all_ranges[0]["start"], all_ranges[0]["end"])
    elif prev_step_feedback:
        requirement_note = (
            "  - **重要**：选区查询结果（如 get_paragraph_range、get_selection_info 的返回值）\n"
            "    **只属于当前 step 内**的后续 action，**不可跨 step 使用**。\n"
            "    其他 step（如 step1）的选区结果与本 step 无关，本 step 必须自行查询。\n"
            "  - 若本 step 操作「全文」，则不需要 rng 参数。\n"
            "  - 若操作「第N段」或「选中部分」，则第一步必须调用对应的选区查询 action：\n"
            "    · 操作「第N段」-> get_paragraph_range(index=N) （1-based）\n"
            "    · 操作「选中部分」-> get_selection_info()\n"
            "  - 后续格式类 action 的 rng 不要填写，系统会自动用本 step 查询结果填充。"
        )
    else:
        requirement_note = (
            "  - 若当前 step 是纯查询类 action（find_all、get_xxx 等），则只需返回该查询 action，\n"
            "    不要混入操作类 action。\n"
            "  - 若当前 step 不是纯查询类 action，则第一步必须调用选区查询 action：\n"
            "    · 操作「第N段」-> get_paragraph_range(index=N) （1-based）\n"
            "    · 操作「选中部分」-> get_selection_info()\n"
            "    · 操作「全文」-> 不需要 rng\n"
            "  - 后续格式类 action（set_font_name 等）的 rng 不要填写，系统会自动填充。"
        )
    task_block += requirement_note + "\n"
    if prev_step_feedback and all_ranges:
        query_note = (
            "2. **step 类型限制**：\n"
            "  - **只允许**返回操作类 action（set_bold、set_font_color、replace 等）。\n"
            "  - **禁止**返回任何查询类 action（find_all、get_selection_info、get_paragraph_range、\n"
            "    get_full_text、find、count_occurrences 等），系统已在上一步完成查询。\n"
        )
        task_block += query_note + "\n"
    elif not prev_step_feedback:
        query_note = (
            "2. **step 类型限制**：\n"
            "  - 查询类型 step（用于获取文档信息，如 find_all、get_selection_info 等）\n"
            "    **只能规划纯查询类 action，不得混入 set/delete/insert 等操作类 action**。\n"
            "  - 操作类型 step（用于修改格式或内容）\n"
            "    **只能规划操作类 action，不得混入查询类 action**。\n"
            "\n"
            "  **注意**：若需求为「先查后改」，必须拆为两个独立 step（查 → 操作），\n"
            "切勿在单个 step 中混合查询与操作类 action。"
        )
        task_block += query_note + "\n"
    else:
        # prev_step_feedback 且无 all_ranges：跨 step 选区结果不可复用，本 step 必须自行查询
        query_note = (
            "2. **step 类型限制**：\n"
            "  - **必须**在当前 step 内添加选区查询 action（除非操作全文）。\n"
            "    选区查询结果只属于当前 step，**不可**使用其他 step 的选区结果。\n"
        )
        task_block += query_note + "\n"
    task_block += (
        "\n## 输出要求\n"
        "只返回 JSON 数组，不要其他文字：\n"
        "[\n"
        + rng_example + "\n"
        "]\n"
    )
    if not (prev_step_feedback and all_ranges):
        task_block += "WARNING: 不要在 params 里写 rng，系统会自动填充。\n"
    return (
        "你是一个专业的 Word 文档格式化助手，正在使用技能「" + skill_name + "」。\n"
        + format_section
        + (hint_section if hint_section else "")
        + pe
        + "\n## 当前上下文\n"
        "用户选中了 Word 文档中的以下内容：\n"
        "---\n" + st + "\n"
        "---\n"
        "\n## 用户的需求\n"
        "\"" + user_message + "\"\n"
        "\n## 技能「" + skill_name + "」的完整说明书\n"
        + skill_content + "\n"
        + task_block
    )


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
