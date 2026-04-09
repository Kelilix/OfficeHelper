# -*- coding: utf-8 -*-
"""
Action 注册表

定义所有 word-text-operator skill 支持的 action 及其参数映射。
以后新增 action，只需在这里添加条目，无需改 routes.py。

映射结构：
    action_name -> {
        "target": ("op", "fmt")          # op.fmt.method()
                   ("op", None)           # op.method()
        "method": "method_name",          # 被调用的方法名
        "rng": True/False,               # 是否需要解析 rng 参数
        "params": {...},                 # action param -> method param 映射
        "result": RESULT_xxx,            # 如何从返回值构建 result
        "returns": "result"/"none",       # 是否将返回值写入 result
    }

对于有复杂逻辑的 action（如 find/goto_bookmark），使用 "handler" 字段指向处理函数。
"""

import inspect
from typing import Any, Callable, Dict, List, Optional, Tuple, TYPE_CHECKING

if TYPE_CHECKING:
    from word_text_operator import WordTextOperator


# ── 结果提取模式 ────────────────────────────────────────────────────────────────

RESULT_NONE = "none"       # 无返回值
RESULT_SINGLE = "single"   # 直接透传（替换次数、书签列表等）
RESULT_BOOL = "bool"       # 转为 True/False
RESULT_INFO = "info"       # 返回结构化 info dict

# ── 执行前状态捕获类型 ─────────────────────────────────────────────────────────
# 每个 action 操作前，需要捕获的初始状态类型。
# "font"      → rng.Font（字体：字号、字名、粗/斜/下划线、颜色、高亮）
# "paragraph" → rng.ParagraphFormat（段落：对齐、行距、缩进、段前/段后间距）
# "border"    → rng.Borders / rng.Shading（边框与底纹）
# "content"   → rng.Text（文本内容，用于插入/删除/替换操作）
# None        → 无需捕获（如查找、导航、书签、只读等操作）

STATE_FONT = "font"
STATE_PARA = "paragraph"
STATE_BORDER = "border"
STATE_CONTENT = "content"


# ── 辅助：解析 rng 参数 ────────────────────────────────────────────────────────

def _resolve_rng(op: "WordTextOperator", rng_param) -> Any:
    """
    将 rng 参数转为 Word Range COM 对象。
    三种格式：
      - "full_document"            → 整篇文档
      - "[start, end]"            → 指定字符范围（优先于 Selection）
      - "" / None / 缺失参数       → 当前 Selection（用户选中的区域）
      - 其他                        → 当前 Selection
    """
    # Word 未连接（connect 失败或从未调用）→ 尝试重新探测
    if op is None or op._base._word_app is None:
        raise RuntimeError("Word 未连接（请确保已打开 Word 文档）")

    # 若 COM 代理已失效，is_connected 会清空 _word_app，重新抛合适的错误
    try:
        _ = op._base._word_app.ActiveDocument
    except Exception:
        op._base._word_app = None
        raise RuntimeError("Word COM 连接已失效（Word 可能已关闭），请重新连接")

    # 显式要求全文档
    if rng_param == "full_document":
        doc = op._base._document
        if doc is None:
            try:
                doc = op._base._word_app.ActiveDocument
            except Exception:
                pass
        if doc is not None:
            return doc.Content
        raise RuntimeError("没有活动 Word 文档，无法操作全文档")

    # 坐标格式 [start, end]（list 或 string 均可，优先级高于 Selection）
    if isinstance(rng_param, (list, tuple)) and len(rng_param) == 2:
        return op.get_range(int(rng_param[0]), int(rng_param[1]))
    if isinstance(rng_param, str) and rng_param.startswith("["):
        import ast
        try:
            coords = ast.literal_eval(rng_param)
            if isinstance(coords, (list, tuple)) and len(coords) == 2:
                return op.get_range(int(coords[0]), int(coords[1]))
        except Exception:
            pass

    # 无显式 rng 参数时：优先使用当前 Selection
    try:
        sel = op._base._word_app.Selection
        if sel is not None and sel.Start != sel.End:
            return sel
        # 折叠选区（光标）也返回（用于插入点操作）
        return sel
    except AttributeError:
        # pywin32 在 Word 无活动文档或 COM 状态异常时 Selection 报 AttributeError
        pass
    except Exception as e:
        # 其他 COM 错误（如 stale proxy）也走 fallback
        pass

    # 最终兜底：尝试用 ActiveDocument.Content
    try:
        doc = op._base._word_app.ActiveDocument
    except Exception:
        pass
    else:
        if doc is not None:
            return doc.Content
    raise RuntimeError("无法获取 Selection（请在 Word 中选中文本）")


# ── 状态捕获 ───────────────────────────────────────────────────────────────────

def _human_readable_state(state_type: str, raw: dict) -> str:
    """
    将原始格式状态 dict 转换为人类可读的自然语言描述。
    只展示与本次操作相关的关键属性，避免信息过载。
    """
    if state_type == STATE_FONT:
        parts = []
        size = raw.get("size")
        name = raw.get("name")
        if size:
            parts.append(f"字号: {size}pt")
        if name:
            parts.append(f"字体: {name}")
        # 样式组合
        styles = []
        if raw.get("bold"):
            styles.append("粗体")
        if raw.get("italic"):
            styles.append("斜体")
        underline_val = raw.get("underline", 0)
        if underline_val and underline_val != 0:
            styles.append("下划线")
        if styles:
            parts.append(f"样式: {', '.join(styles)}")
        else:
            parts.append("样式: 常规")
        color = raw.get("color")
        if color and color != 0xFFFFFFFF and color != 0:
            parts.append(f"颜色: #{color:06X}")
        highlight = raw.get("highlight")
        if highlight and highlight != -1:
            parts.append(f"高亮: {highlight}")
        return " | ".join(parts) if parts else "字体格式: 默认"

    elif state_type == STATE_PARA:
        parts = []
        align_map = {0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐", 4: "分散对齐"}
        align = raw.get("alignment")
        if align is not None:
            parts.append(f"对齐: {align_map.get(align, str(align))}")
        ls = raw.get("line_spacing")
        if ls:
            parts.append(f"行距: {ls}")
        fi = raw.get("first_line_indent")
        if fi:
            parts.append(f"首行缩进: {fi}")
        li = raw.get("left_indent")
        if li:
            parts.append(f"左缩进: {li}")
        ri = raw.get("right_indent")
        if ri:
            parts.append(f"右缩进: {ri}")
        sb = raw.get("space_before")
        if sb:
            parts.append(f"段前间距: {sb}")
        sa = raw.get("space_after")
        if sa:
            parts.append(f"段后间距: {sa}")
        return " | ".join(parts) if parts else "段落格式: 默认"

    elif state_type == STATE_BORDER:
        return f"边框/底纹: {raw}"

    elif state_type == STATE_CONTENT:
        text = raw.get("text", "")
        return f"内容: 「{text[:50]}{'...' if len(text) > 50 else ''}」"

    return str(raw)


def _capture_state(
    state_type: str, rng: Any, op: "WordTextOperator",
) -> Optional[dict]:
    """在操作执行前捕获初始状态，返回人类可读字符串。"""
    if not state_type:
        return None

    try:
        if state_type == STATE_FONT:
            return op._fmt.get_font_info(rng)
        elif state_type == STATE_PARA:
            return op._fmt.get_paragraph_format_info(rng)
        elif state_type == STATE_BORDER:
            return {}
        elif state_type == STATE_CONTENT:
            return {"text": rng.Text}
    except Exception:
        pass
    return None


# ── 自定义处理器 ────────────────────────────────────────────────────────────────


def _capture_para_state(
    para_op, idx: int,
) -> tuple[Optional[dict], str]:
    """
    捕获段落格式初始状态（供段落写操作 handler 调用）。
    idx 为 1-based，返回 (raw_dict, human_readable_str)。
    """
    try:
        para = para_op.get(idx)
        raw = para_op.get_format_info(para)
        text = _human_readable_state(STATE_PARA, raw) if raw else ""
        return raw, text
    except Exception:
        return None, ""


def _h_find(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """find：返回 {start, end, text} 或 None"""
    try:
        rng = op.find(
            params.get("text", ""),
            whole_word=params.get("whole_word", False),
            match_case=params.get("match_case", False),
        )
        return {
            "success": rng is not None,
            "result": {"start": rng.Start, "end": rng.End, "text": rng.Text} if rng else None
        }
    except Exception as e:
        return {"success": False, "error": f"find 失败：{e}"}


def _h_count_occurrences(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    count = op.count_occurrences(params.get("text", ""))
    return {"success": True, "result": count}


def _h_find_wildcards(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    pattern = params.get("pattern", "")
    replace_text = params.get("replace_text")
    if replace_text:
        n = op.find_wildcards(pattern, replace_text)
    else:
        n = 1 if op.find_wildcards(pattern) else 0
    return {"success": True, "result": n}


def _h_goto_bookmark(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    rng = op.go_to_bookmark(params.get("name", ""))
    if rng:
        op.select(rng)
    return {"success": rng is not None}


def _h_goto_page(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    page = int(params.get("page", 1))
    rng = op.nav.go_to_page(page)
    op.select(rng)
    return {"success": True}


def _h_goto_line(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    line = int(params.get("line", 1))
    rng = op.nav.go_to_line(line)
    op.select(rng)
    return {"success": True}


def _h_create_bookmark(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    ok = op.create_bookmark(params.get("name", ""),
                             int(params.get("start", 0)),
                             int(params.get("end", 0)))
    return {"success": ok}


def _h_bookmark_text(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    ok = op.bookmark_text(params.get("name", ""), params.get("text", ""))
    return {"success": ok}


def _h_wrap_with_bookmarks(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    ok = op.wrap_with_bookmarks(
        params.get("text", ""),
        params.get("open_name", ""),
        params.get("close_name", ""),
    )
    return {"success": ok}


def _h_insert_text(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    rng = _resolve_rng(op, params.get("rng", ""))
    before_state_raw = _capture_state(STATE_CONTENT, rng, op)
    before_state_text = _human_readable_state(STATE_CONTENT, before_state_raw) if before_state_raw else ""
    op.insert_text(rng, params.get("text", ""), before=params.get("before", True))
    result = {"success": True}
    if before_state_text:
        result["before_state"] = before_state_text
    if before_state_raw:
        result["before_state_raw"] = before_state_raw
    return result


def _h_insert_page_break(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    rng = _resolve_rng(op, params.get("rng", ""))
    op.insert_page_break(rng)
    return {"success": True}


def _h_insert_file(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    rng = _resolve_rng(op, params.get("rng", ""))
    op.text.insert_file(rng, params.get("file_path", ""))
    return {"success": True}


def _h_insert_symbol(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    rng = _resolve_rng(op, params.get("rng", ""))
    op.text.insert_symbol(
        rng,
        int(params.get("character_code", 9744)),
        params.get("font_name"),
        unicode=params.get("unicode", False),
    )
    return {"success": True}


def _h_insert_paragraph(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    rng = _resolve_rng(op, params.get("rng", ""))
    op.text.insert_paragraph(rng)
    return {"success": True}


def _h_delete_range(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    rng = _resolve_rng(op, params.get("rng", ""))
    before_state_raw = _capture_state(STATE_CONTENT, rng, op)
    before_state_text = _human_readable_state(STATE_CONTENT, before_state_raw) if before_state_raw else ""
    op.delete_range(rng)
    result = {"success": True}
    if before_state_text:
        result["before_state"] = before_state_text
    if before_state_raw:
        result["before_state_raw"] = before_state_raw
    return result


def _h_clear_range(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    rng = _resolve_rng(op, params.get("rng", ""))
    before_state_raw = _capture_state(STATE_CONTENT, rng, op)
    before_state_text = _human_readable_state(STATE_CONTENT, before_state_raw) if before_state_raw else ""
    op.text.clear(rng)
    result = {"success": True}
    if before_state_text:
        result["before_state"] = before_state_text
    if before_state_raw:
        result["before_state_raw"] = before_state_raw
    return result


def _h_select(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    rng = _resolve_rng(op, params.get("rng", ""))
    op.select(rng)
    return {"success": True}


def _h_char_count(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    rng = _resolve_rng(op, params.get("rng", ""))
    n = op.char_count(rng)
    return {"success": True, "result": n}


def _h_word_count(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    rng = _resolve_rng(op, params.get("rng", ""))
    n = op.word_count(rng)
    return {"success": True, "result": n}


def _h_sentence_count(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    rng = _resolve_rng(op, params.get("rng", ""))
    n = op.text.sentence_count(rng)
    return {"success": True, "result": n}


def _h_paragraph_count(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    rng = _resolve_rng(op, params.get("rng", ""))
    n = op.text.paragraph_count(rng)
    return {"success": True, "result": n}


def _apply_first_line_indent_chars(rng: Any, n_chars: float, op: "WordTextOperator") -> None:
    """首行缩进 n 个字符：优先 Word 的 CharacterUnitFirstLineIndent，否则按字号换算为磅。"""
    pf = rng.ParagraphFormat
    try:
        pf.CharacterUnitFirstLineIndent = int(n_chars)
        return
    except Exception:
        pass
    try:
        op._fmt.set_first_line_indent(rng, characters=float(n_chars))
        return
    except Exception:
        pass
    fs = rng.Font.Size
    pts = float(n_chars) * (float(fs) if fs and fs > 0 else 12.0)
    pf.FirstLineIndent = pts


def _h_apply_format_to_range(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """
    批量设置 [start, end] 内段落格式（索引 1-based 闭区间，与技能一致）。
    识别 params 中非空字段：alignment, spacing+rule, first_line_indent_characters / _cm,
    indent_left_characters/cm, indent_right_characters/cm, space_before, space_after（磅）。
    """
    if op._base._word_app is None:
        raise RuntimeError("Word 未连接（请确保已打开 Word 文档）")
    doc = op._base.document
    if doc is None:
        raise RuntimeError("没有活动 Word 文档")

    n = int(doc.Paragraphs.Count)
    if n < 1:
        return {"success": False, "error": "文档无段落"}

    start = int(params.get("start", 1))
    end = int(params.get("end", n))
    # 与单段 API 一致：负数索引表示从文末倒数（-1 = 最后一段）
    if start < 0:
        start = n + start + 1
    if end < 0:
        end = n + end + 1
    start = max(1, min(start, n))
    end = max(1, min(end, n))
    if start > end:
        start, end = end, start

    keys = (
        "alignment",
        "spacing",
        "rule",
        "first_line_indent_characters",
        "first_line_indent_cm",
        "indent_left_characters",
        "indent_left_cm",
        "indent_right_characters",
        "indent_right_cm",
        "space_before",
        "space_after",
    )
    if not any(params.get(k) is not None for k in keys if k not in ("rule",)):
        if params.get("spacing") is None and params.get("rule") is not None:
            pass
        elif not any(params.get(k) is not None for k in keys):
            return {"success": False, "error": "未提供任何格式参数（如 first_line_indent_characters）"}

    for i in range(start, end + 1):
        rng = doc.Paragraphs(i).Range

        al = params.get("alignment")
        if al is not None:
            op.set_paragraph_alignment(rng, al)

        sp = params.get("spacing")
        if sp is not None:
            rule = params.get("rule", "single")
            op._fmt.set_line_spacing(rng, float(sp), rule)

        fic = params.get("first_line_indent_characters")
        if fic is not None:
            _apply_first_line_indent_chars(rng, float(fic), op)

        ficm = params.get("first_line_indent_cm")
        if ficm is not None:
            op._fmt.set_first_line_indent(rng, cm=float(ficm))

        ilc = params.get("indent_left_characters")
        ilm = params.get("indent_left_cm")
        if ilc is not None or ilm is not None:
            op._fmt.set_indent_left(rng, characters=float(ilc) if ilc is not None else None,
                                    cm=float(ilm) if ilm is not None else None)

        irc = params.get("indent_right_characters")
        irm = params.get("indent_right_cm")
        if irc is not None or irm is not None:
            op._fmt.set_indent_right(rng, characters=float(irc) if irc is not None else None,
                                     cm=float(irm) if irm is not None else None)

        sb = params.get("space_before")
        if sb is not None:
            op._fmt.set_space_before_para(rng, float(sb))

        sa = params.get("space_after")
        if sa is not None:
            op._fmt.set_space_after_para(rng, float(sa))

    text_state, raw = "", None
    if para_op is not None:
        try:
            raw, text_state = _capture_para_state(para_op, start)
        except Exception:
            pass
    result = {"success": True, "result": end - start + 1}
    if text_state: result["before_state"] = text_state
    if raw: result["before_state_raw"] = raw
    return result


def _h_get_full_text(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    """返回全文范围信息（start/end），供 rng 占位符自动填充。"""
    doc = op._base._document
    if doc is None:
        try:
            doc = op._base._word_app.ActiveDocument
        except Exception:
            pass
    if doc is None:
        return {"success": False, "error": "没有活动 Word 文档"}
    content = doc.Content
    return {
        "success": True,
        "result": {
            "text": content.Text,
            "start": content.Start,
            "end": content.End,
        },
    }


def _h_get_text(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    rng = _resolve_rng(op, params.get("rng", ""))
    return {"success": True, "result": rng.Text}


def _h_get_selection_text(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    return {"success": True, "result": op.get_selection_text()}


def _h_get_paragraph_text(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    """
    与 word-paragraph-operator 技能一致：index 从 1 开始；-1 表示最后一段。
    底层 op.text.get_paragraph_text 为 0 起始（第 1 段传 0）。
    """
    if op._base._word_app is None:
        raise RuntimeError("Word 未连接（请确保已打开 Word 文档）")
    idx_raw = params.get("index")
    if idx_raw is None:
        return {"success": False, "error": "缺少参数 index"}
    idx = int(idx_raw)
    doc = op._base.document
    if doc is None:
        raise RuntimeError("没有活动 Word 文档")

    if idx == -1:
        n = int(doc.Paragraphs.Count)
        if n < 1:
            return {"success": False, "error": "文档无段落"}
        para_idx0 = n - 1
    elif idx >= 1:
        para_idx0 = idx - 1
        if idx > int(doc.Paragraphs.Count):
            return {"success": False, "error": f"段落索引超出范围（共 {doc.Paragraphs.Count} 段）"}
    else:
        return {"success": False, "error": f"无效的段落索引: {idx}（须为 ≥1 或 -1）"}

    text = op.text.get_paragraph_text(para_idx0)
    return {"success": True, "result": text}


def _h_find_all(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    import logging as _fa_logger
    _fa_logger.getLogger(__name__).info("[_h_find_all] 开始执行 | params=%s", params)
    positions = op.find_all(params.get("text", ""))
    _fa_logger.getLogger(__name__).info("[_h_find_all] 完成 | positions=%s", positions)
    return {"success": True, "result": positions}


def _h_list_bookmarks(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    return {"success": True, "result": op.get_bookmarks()}


def _h_export_bookmarks(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    op.bm.export_bookmarks(params.get("path", "bookmarks.json"))
    return {"success": True}


def _h_import_bookmarks(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    n = op.bm.import_bookmarks(params.get("path", "bookmarks.json"))
    return {"success": True, "result": n}


def _h_get_selection_info(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    try:
        return {"success": True, "result": op.sel.get_selection_info()}
    except (AttributeError, RuntimeError, Exception) as e:
        # pywin32 在 Word 某些状态（如无活动文档）或 COM 代理失效时
        # 会报 AttributeError / RuntimeError / COMError
        return {"success": False, "error": f"无法获取 Selection：{e}"}


def _h_replace(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    n = op.replace(
        params.get("find_text", ""),
        params.get("replace_text", ""),
        whole_word=params.get("whole_word", False),
        match_case=params.get("match_case", False),
        replace_all=params.get("replace_all", True),
    )
    return {"success": True, "result": n}


def _h_replace_with_format(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    n = op.replace_with_format(
        params.get("find_text", ""),
        params.get("replace_text", ""),
        bold=params.get("bold", False),
        italic=params.get("italic", False),
    )
    return {"success": True, "result": n}


def _h_batch_replace(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    return {"success": True, "result": op.batch_replace(params.get("replacements", {}))}


def _h_find_with_format(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    found = op.find_with_format(
        params.get("text", ""),
        bold=params.get("bold"),
        italic=params.get("italic"),
    )
    return {"success": found}


def _h_new_document(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    op.new_document()
    return {"success": True}


def _h_save(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    op.save(params.get("path"))
    return {"success": True}


# ── word-paragraph-operator 段落操作 handlers ─────────────────────────────────

def _h_get_paragraph_count(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """段落总数（整篇文档）。"""
    if para_op is None:
        return {"action": "get_paragraph_count", "description": "获取段落总数", "success": False, "error": "ParagraphOperator 未初始化"}
    n = para_op.count()
    return {"success": True, "result": n}


def _h_get_paragraph_by_index(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """按索引获取段落（1-based，-1=最后一段）。"""
    if para_op is None:
        return {"action": "get_paragraph_by_index", "description": "按索引获取段落", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        para = para_op.get(idx)
        return {"success": True, "result": para_op.get_text(para)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_paragraph_text(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """读取指定段落的纯文本（index 从 1 开始）。"""
    if para_op is None:
        return {"action": "get_paragraph_text", "description": "读取段落文本", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        para = para_op.get(idx)
        return {"success": True, "result": para_op.get_text(para)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_paragraph_range(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """
    获取指定段落（index）或段落范围（start~end）的字符位置 [start, end]。

    用法：
    - index: 1-based 段落索引，返回该段字符位置
    - start+end: 返回范围内所有段落的字符位置列表
    返回均为字符 Range（可直接用于 rng 参数）。
    """
    if para_op is None:
        return {"action": "get_paragraph_range", "description": "获取段落范围", "success": False, "error": "ParagraphOperator 未初始化"}
    try:
        # 单段落：index 参数（1-based）
        if "index" in params:
            idx = int(params["index"])
            paras = [para_op.get(idx)]
        else:
            start = int(params.get("start", 1))
            end = int(params.get("end", para_op.count()))
            paras = para_op.range(start, end)

        # 转换为字符位置 [start, end]
        results = []
        for p in paras:
            r = p.Range
            results.append([r.Start, r.End])

        if len(results) == 1:
            return {"success": True, "result": results[0]}
        return {"success": True, "result": results}
    except Exception as e:
        return {"action": "get_paragraph_range", "description": "获取段落范围", "success": False, "error": str(e)}


def _h_get_document_structure(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """文档结构摘要（每段 index/text/style/level/is_heading/is_empty/is_list）。"""
    if para_op is None:
        return {"action": "get_document_structure", "description": "获取文档结构", "success": False, "error": "ParagraphOperator 未初始化"}
    return {"success": True, "result": para_op.get_document_structure()}


def _h_get_outline_summary(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """大纲摘要（用于生成目录）。"""
    if para_op is None:
        return {"action": "get_outline_summary", "description": "获取大纲摘要", "success": False, "error": "ParagraphOperator 未初始化"}
    return {"success": True, "result": para_op.get_outline_summary()}


def _h_find_empty_paragraphs(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """查找所有空段落。"""
    if para_op is None:
        return {"action": "find_empty_paragraphs", "description": "查找空段落", "success": False, "error": "ParagraphOperator 未初始化"}
    paras = para_op.find_empty_paragraphs()
    return {"success": True, "result": [para_op.get_index(p) for p in paras]}


def _h_find_heading_paragraphs(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """查找所有标题段落。"""
    if para_op is None:
        return {"action": "find_heading_paragraphs", "description": "查找标题段落", "success": False, "error": "ParagraphOperator 未初始化"}
    paras = para_op.find_headings()
    return {"success": True, "result": [para_op.get_index(p) for p in paras]}


def _h_find_paragraphs_by_level(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """按大纲级别查找标题段落。"""
    if para_op is None:
        return {"action": "find_paragraphs_by_level", "description": "按级别查找标题", "success": False, "error": "ParagraphOperator 未初始化"}
    level = int(params.get("level", 1))
    paras = para_op.find_headings_by_level(level)
    return {"success": True, "result": [para_op.get_index(p) for p in paras]}


def _h_find_paragraphs_by_text(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """按文本内容查找段落。"""
    if para_op is None:
        return {"action": "find_paragraphs_by_text", "description": "按文本查找段落", "success": False, "error": "ParagraphOperator 未初始化"}
    text = params.get("text", "")
    whole_word = params.get("whole_word", False)
    paras = para_op.find_by_text(text, whole_word=whole_word)
    return {"success": True, "result": [para_op.get_index(p) for p in paras]}


def _h_get_paragraph_format_info(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """读取段落完整格式（对齐/行距/缩进/间距/样式）。"""
    if para_op is None:
        return {"action": "get_paragraph_format_info", "description": "读取段落格式", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        para = para_op.get(idx)
        return {"success": True, "result": para_op.get_format_info(para)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_paragraph_style(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """获取段落样式名称。"""
    if para_op is None:
        return {"action": "get_paragraph_style", "description": "获取样式名称", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        para = para_op.get(idx)
        return {"success": True, "result": para_op.get_style_name(para)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_outline_level(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """获取段落大纲级别。"""
    if para_op is None:
        return {"action": "get_outline_level", "description": "获取大纲级别", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        para = para_op.get(idx)
        return {"success": True, "result": para_op.get_outline_level(para)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_is_paragraph_list_item(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """判断段落是否属于编号列表。"""
    if para_op is None:
        return {"action": "is_paragraph_list_item", "description": "是否编号列表项", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        para = para_op.get(idx)
        return {"success": True, "result": para_op.is_list_item(para)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_is_paragraph_in_table(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """判断段落是否在表格内。"""
    if para_op is None:
        return {"action": "is_paragraph_in_table", "description": "是否在表格中", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        para = para_op.get(idx)
        return {"success": True, "result": para_op.is_in_table(para)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_paragraph_alignment(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """设置段落对齐方式。"""
    if para_op is None:
        return {"action": "set_paragraph_alignment", "description": "设置段落对齐", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    align = params.get("alignment", "left")
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.set_alignment(para, align)
        result = {"success": True}
        if text:
            result["before_state"] = text
        if raw:
            result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_paragraph_line_spacing(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """设置段落行间距。"""
    if para_op is None:
        return {"action": "set_paragraph_line_spacing", "description": "设置行距", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    spacing = params.get("spacing")
    rule = params.get("rule", "single")
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.set_line_spacing(para, value=spacing, rule=rule)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_paragraph_space_before(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """设置段前间距。"""
    if para_op is None:
        return {"action": "set_paragraph_space_before", "description": "设置段前间距", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    points = float(params.get("points", 0))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.set_space_before(para, points)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_paragraph_space_after(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """设置段后间距。"""
    if para_op is None:
        return {"action": "set_paragraph_space_after", "description": "设置段后间距", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    points = float(params.get("points", 0))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.set_space_after(para, points)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_paragraph_indent_left(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """左缩进。"""
    if para_op is None:
        return {"action": "set_paragraph_indent_left", "description": "设置左缩进", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        cm = params.get("cm")
        chars = params.get("characters")
        para_op.set_indent_left(para, characters=float(chars) if chars else None,
                                cm=float(cm) if cm else None)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_paragraph_indent_right(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """右缩进。"""
    if para_op is None:
        return {"action": "set_paragraph_indent_right", "description": "设置右缩进", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        cm = params.get("cm")
        chars = params.get("characters")
        para_op.set_indent_right(para, characters=float(chars) if chars else None,
                                  cm=float(cm) if cm else None)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_paragraph_first_line_indent(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """首行缩进（传负值为悬挂缩进）。"""
    if para_op is None:
        return {"action": "set_paragraph_first_line_indent", "description": "设置首行缩进", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        cm = params.get("cm")
        chars = params.get("characters")
        para_op.set_first_line_indent(para, characters=float(chars) if chars else None,
                                      cm=float(cm) if cm else None)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_paragraph_hanging_indent(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """悬挂缩进。"""
    if para_op is None:
        return {"action": "set_paragraph_hanging_indent", "description": "设置悬挂缩进", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    characters = float(params.get("characters", 0))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.set_hanging_indent(para, characters)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_paragraph_outline_level(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """设置大纲级别。"""
    if para_op is None:
        return {"action": "set_paragraph_outline_level", "description": "设置大纲级别", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    level = int(params.get("level", 0))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.set_outline_level(para, level)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_paragraph_keep_together(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """段内不分页。"""
    if para_op is None:
        return {"action": "set_paragraph_keep_together", "description": "段内不分页", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    on = bool(params.get("on", True))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.set_keep_together(para, on)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_paragraph_keep_with_next(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """与下段同页。"""
    if para_op is None:
        return {"action": "set_paragraph_keep_with_next", "description": "与下段同页", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    on = bool(params.get("on", True))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.set_keep_with_next(para, on)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_paragraph_style(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """应用样式到段落。"""
    if para_op is None:
        return {"action": "set_paragraph_style", "description": "应用样式", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    style_name = params.get("style_name", "Normal")
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.set_style(para, style_name)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_reset_paragraph_format(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """重置段落格式。"""
    if para_op is None:
        return {"action": "reset_paragraph_format", "description": "重置段落格式", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.reset_format(para)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_paragraph_border(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """给段落添加边框。"""
    if para_op is None:
        return {"action": "set_paragraph_border", "description": "添加段落边框", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.set_border(
            para,
            side=params.get("side", "bottom"),
            line_style=int(params.get("line_style", 1)),
            line_width=int(params.get("line_width", 4)),
            color=params.get("color", 0x000000),
            space=float(params.get("space", 6.0)),
        )
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_clear_paragraph_border(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """清除段落边框。"""
    if para_op is None:
        return {"action": "clear_paragraph_border", "description": "清除段落边框", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.clear_border(para)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_paragraph_shading(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """设置段落底纹。"""
    if para_op is None:
        return {"action": "set_paragraph_shading", "description": "设置段落底纹", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.set_shading(para, fill_color=params.get("fill_color", 0xFFFF00))
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_clear_paragraph_shading(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """清除段落底纹。"""
    if para_op is None:
        return {"action": "clear_paragraph_shading", "description": "清除段落底纹", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.clear_shading(para)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_list_paragraphs(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """返回所有编号/项目符号列表段落。"""
    if para_op is None:
        return {"action": "get_list_paragraphs", "description": "获取列表段落", "success": False, "error": "ParagraphOperator 未初始化"}
    paras = para_op.find_list_paragraphs()
    return {"success": True, "result": [para_op.get_index(p) for p in paras]}


def _h_get_list_level(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """获取段落在列表中的级别（1-based）。"""
    if para_op is None:
        return {"action": "get_list_level", "description": "获取列表级别", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        para = para_op.get(idx)
        return {"success": True, "result": para_op.get_list_level(para)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_list_level(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """设置段落在列表中的级别。"""
    if para_op is None:
        return {"action": "set_list_level", "description": "设置列表级别", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    level = int(params.get("level", 1))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.set_list_level(para, level)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_apply_bullet_list(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """将段落转换为项目符号列表。"""
    if para_op is None:
        return {"action": "apply_bullet_list", "description": "转为项目符号", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.apply_bullet(para)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_apply_numbered_list(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """将段落转换为编号列表。"""
    if para_op is None:
        return {"action": "apply_numbered_list", "description": "转为编号列表", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    number_format = params.get("number_format", "decimal")
    start_at = int(params.get("start_at", 1))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.apply_numbering(para, number_format=number_format, start_at=start_at)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_remove_list_format(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """移除段落的列表格式。"""
    if para_op is None:
        return {"action": "remove_list_format", "description": "移除列表格式", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        raw, text = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.remove_list_format(para)
        result = {"success": True}
        if text: result["before_state"] = text
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_paragraph_text(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """替换段落内容（保留段落格式）。"""
    if para_op is None:
        return {"action": "set_paragraph_text", "description": "替换段落文本", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    text = params.get("text", "")
    try:
        raw, text_state = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.set_text(para, text)
        result = {"success": True}
        if text_state: result["before_state"] = text_state
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_insert_text_before_paragraph(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """在段落开头插入文本。"""
    if para_op is None:
        return {"action": "insert_text_before_paragraph", "description": "段落前插入文本", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    text = params.get("text", "")
    try:
        raw, text_state = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.insert_text_before(para, text)
        result = {"success": True}
        if text_state: result["before_state"] = text_state
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_insert_text_after_paragraph(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """在段落末尾插入文本。"""
    if para_op is None:
        return {"action": "insert_text_after_paragraph", "description": "段落后插入文本", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    text = params.get("text", "")
    try:
        raw, text_state = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.insert_text_after(para, text)
        result = {"success": True}
        if text_state: result["before_state"] = text_state
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_delete_paragraph(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """删除整个段落（慎用：会合并相邻段落）。"""
    if para_op is None:
        return {"action": "delete_paragraph", "description": "删除段落", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        raw, text_state = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.delete_paragraph(para)
        result = {"success": True}
        if text_state: result["before_state"] = text_state
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_clear_paragraph(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """清空段落内容（保留段落标记）。"""
    if para_op is None:
        return {"action": "clear_paragraph", "description": "清空段落", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        raw, text_state = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        para_op.clear_paragraph(para)
        result = {"success": True}
        if text_state: result["before_state"] = text_state
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_add_paragraph_after(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """在指定段落之后插入新段落。"""
    if para_op is None:
        return {"action": "add_paragraph_after", "description": "段落后插入新段落", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    text = params.get("text", "")
    try:
        raw, text_state = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        new_para = para_op.add_paragraph_after(para)
        if text:
            para_op.set_text(new_para, text)
        result = {"success": True}
        if text_state: result["before_state"] = text_state
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_add_paragraph_before(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """在指定段落之前插入新段落。"""
    if para_op is None:
        return {"action": "add_paragraph_before", "description": "段落前插入新段落", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    text = params.get("text", "")
    try:
        raw, text_state = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        new_para = para_op.add_paragraph_before(para)
        if text:
            para_op.set_text(new_para, text)
        result = {"success": True}
        if text_state: result["before_state"] = text_state
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_merge_with_next(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """将当前段落与下一段合并。"""
    if para_op is None:
        return {"action": "merge_with_next", "description": "与下一段合并", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        raw, text_state = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        ok = para_op.merge_with_next(para)
        result = {"success": ok, "result": ok}
        if text_state: result["before_state"] = text_state
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_merge_with_previous(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """将当前段落与上一段合并。"""
    if para_op is None:
        return {"action": "merge_with_previous", "description": "与上一段合并", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        raw, text_state = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        ok = para_op.merge_with_previous(para)
        result = {"success": ok, "result": ok}
        if text_state: result["before_state"] = text_state
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_split_paragraph_by_separator(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """按分隔符将段落拆分为多个。"""
    if para_op is None:
        return {"action": "split_paragraph_by_separator", "description": "拆分段落", "success": False, "error": "ParagraphOperator 未初始化"}
    idx = int(params.get("index", 1))
    separator = params.get("separator", "\t")
    try:
        raw, text_state = _capture_para_state(para_op, idx)
        para = para_op.get(idx)
        results = para_op.split_paragraph(para, separator=separator)
        result = {"success": True, "result": [para_op.get_index(p) for p in results]}
        if text_state: result["before_state"] = text_state
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_reverse_paragraph_order(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """反转指定范围内段落顺序。"""
    if para_op is None:
        return {"action": "reverse_paragraph_order", "description": "反转段落顺序", "success": False, "error": "ParagraphOperator 未初始化"}
    start = int(params.get("start", 1))
    end = int(params.get("end", para_op.count()))
    try:
        raw, text_state = _capture_para_state(para_op, start)
        paras = para_op.reverse_order(start=start, end=end)
        result = {"success": True, "result": [para_op.get_index(p) for p in paras]}
        if text_state: result["before_state"] = text_state
        if raw: result["before_state_raw"] = raw
        return result
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_delete_empty_paragraphs(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """删除所有空段落。"""
    if para_op is None:
        return {"action": "delete_empty_paragraphs", "description": "删除所有空段落", "success": False, "error": "ParagraphOperator 未初始化"}
    paras = para_op.find_empty_paragraphs()
    count = 0
    for para in paras:
        para_op.delete_paragraph(para)
        count += 1
    result = {"success": True, "result": count}
    return result


def _h_get_paragraph_at_selection(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """获取当前 Selection 所在段落。"""
    if para_op is None:
        return {"action": "get_paragraph_at_selection", "description": "获取光标所在段落", "success": False, "error": "ParagraphOperator 未初始化"}
    para = para_op.get_paragraph_at_selection()
    return {"success": True, "result": para_op.get_index(para)}


def _h_select_paragraph(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """选中指定段落；未传 index 时选中当前 Selection 所在段落（与 move 导航配合）。"""
    if para_op is None:
        return {"action": "select_paragraph", "description": "选中段落", "success": False, "error": "ParagraphOperator 未初始化"}
    try:
        if params.get("index") is not None:
            idx = int(params["index"])
            para = para_op.get(idx)
        else:
            para = para_op.get_paragraph_at_selection()
            idx = para_op.get_index(para)
        raw, text_state = _capture_para_state(para_op, idx)
        para_op.select_paragraph(para)
        result = {"success": True}
        if text_state: result["before_state"] = text_state
        if raw: result["before_state_raw"] = raw
        return result
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_select_paragraph_range(action: Dict, params: Dict, op: "WordTextOperator", para_op=None) -> Dict[str, Any]:
    """选中指定范围内的所有段落。"""
    if para_op is None:
        return {"action": "select_paragraph_range", "description": "选中段落范围", "success": False, "error": "ParagraphOperator 未初始化"}
    start = int(params.get("start", 1))
    end = int(params.get("end", para_op.count()))
    try:
        para_op.select_range_of_paragraphs(start, end)
        return {"success": True}
    except IndexError as e:
        return {"success": False, "error": str(e)}


# ── word-page-operator 页面操作 handlers ──────────────────────────────────────

def _h_get_section_count(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """节总数。"""
    if page_op is None:
        return {"action": "get_section_count", "description": "获取节总数", "success": False, "error": "PageOperator 未初始化"}
    return {"success": True, "result": page_op.count()}


def _h_get_section_by_index(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """按索引获取节。"""
    if page_op is None:
        return {"action": "get_section_by_index", "description": "按索引获取节", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        sec = page_op.get(idx)
        return {"success": True, "result": page_op.get_index(sec)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_current_section_index(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """获取光标所在节的索引。"""
    if page_op is None:
        return {"action": "get_current_section_index", "description": "获取光标所在节", "success": False, "error": "PageOperator 未初始化"}
    return {"success": True, "result": page_op.section.get_current_section_index()}


def _h_get_page_setup_info(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """读取完整页面设置信息。"""
    if page_op is None:
        return {"action": "get_page_setup_info", "description": "读取页面设置", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        return {"success": True, "result": page_op.get_page_setup_info(idx)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_page_margins(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """读取页边距。"""
    if page_op is None:
        return {"action": "get_page_margins", "description": "读取页边距", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        return {"success": True, "result": page_op.get_page_margins(idx)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_paper_size(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """读取纸张大小。"""
    if page_op is None:
        return {"action": "get_paper_size", "description": "读取纸张大小", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        return {"success": True, "result": page_op.get_paper_size(idx)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_orientation(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """读取纸张方向。"""
    if page_op is None:
        return {"action": "get_orientation", "description": "读取纸张方向", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        return {"success": True, "result": page_op.get_orientation(idx)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_column_count(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """读取分栏数。"""
    if page_op is None:
        return {"action": "get_column_count", "description": "读取分栏数", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        return {"success": True, "result": page_op.get_column_count(idx)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_column_info(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """读取分栏详情。"""
    if page_op is None:
        return {"action": "get_column_info", "description": "读取分栏详情", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        return {"success": True, "result": page_op.get_column_info(idx)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_section_start_type(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """读取节起始类型。"""
    if page_op is None:
        return {"action": "get_section_start_type", "description": "读取节起始类型", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        return {"success": True, "result": page_op.section.get_section_start_type(idx)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_page_margins(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """设置页边距。"""
    if page_op is None:
        return {"action": "set_page_margins", "description": "设置页边距", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.set_page_margins(
            idx,
            top=float(params["top"]) if params.get("top") is not None else None,
            bottom=float(params["bottom"]) if params.get("bottom") is not None else None,
            left=float(params["left"]) if params.get("left") is not None else None,
            right=float(params["right"]) if params.get("right") is not None else None,
        )
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_set_page_margins_by_inch(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """用英寸设置页边距。"""
    if page_op is None:
        return {"action": "set_page_margins_by_inch", "description": "英寸设置页边距", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.set_page_margins_by_inch(
            idx,
            top=float(params["top"]) if params.get("top") is not None else None,
            bottom=float(params["bottom"]) if params.get("bottom") is not None else None,
            left=float(params["left"]) if params.get("left") is not None else None,
            right=float(params["right"]) if params.get("right") is not None else None,
        )
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_set_page_margins_preset(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """使用预设方案设置页边距。"""
    if page_op is None:
        return {"action": "set_page_margins_preset", "description": "预设页边距", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    preset = params.get("preset", "normal")
    try:
        page_op.set_page_margins_preset(idx, preset)
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_set_paper_size(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """设置纸张大小。"""
    if page_op is None:
        return {"action": "set_paper_size", "description": "设置纸张大小", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.set_paper_size(
            idx,
            width=float(params["width"]),
            height=float(params["height"]),
        )
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_set_paper_size_preset(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """使用预设纸张类型。"""
    if page_op is None:
        return {"action": "set_paper_size_preset", "description": "预设纸张", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    preset = params.get("preset", "A4")
    try:
        page_op.set_paper_size_preset(idx, preset)
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_set_orientation(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """设置纸张方向。"""
    if page_op is None:
        return {"action": "set_orientation", "description": "设置纸张方向", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    orient = params.get("orientation", "portrait")
    try:
        page_op.set_orientation(idx, orient)
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_set_columns(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """设置分栏数。"""
    if page_op is None:
        return {"action": "set_columns", "description": "设置分栏数", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    count = int(params.get("count", 1))
    equal = bool(params.get("equal_width", True))
    try:
        page_op.set_columns(idx, count, equal_width=equal)
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_set_columns_with_gutter(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """设置分栏数并指定栏间距。"""
    if page_op is None:
        return {"action": "set_columns_with_gutter", "description": "设置栏间距", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    count = int(params.get("count", 1))
    spacing = float(params.get("spacing", 0.75))
    equal = bool(params.get("equal_width", True))
    try:
        page_op.set_columns_with_gutter(idx, count, spacing, equal_width=equal)
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_set_columns_equal_width(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """设为等宽栏。"""
    if page_op is None:
        return {"action": "set_columns_equal_width", "description": "等宽栏", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.set_columns_equal_width(idx)
        return {"success": True}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_column_width(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """设置指定栏的宽度。"""
    if page_op is None:
        return {"action": "set_column_width", "description": "设置栏宽", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    column = int(params.get("column", 1))
    width = float(params.get("width", 8.0))
    try:
        page_op.page.set_column_width(idx, column, width)
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_apply_two_column_layout(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """应用两栏布局。"""
    if page_op is None:
        return {"action": "apply_two_column_layout", "description": "两栏布局", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    with_line = bool(params.get("with_line", True))
    try:
        page_op.apply_two_column_layout(idx, with_line)
        return {"success": True}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_delete_section_break(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """删除指定节前面的分节符，或删除所有分节符（index=0）。"""
    if page_op is None:
        return {"action": "delete_section_break", "description": "删除分节符", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 0))
    try:
        page_op.delete_section_break(idx)
        if idx == 0:
            return {"success": True, "result": "已删除所有分节符，文档合并为单节"}
        return {"success": True, "result": f"已删除第 {idx} 节前的分节符"}
    except Exception as e:
        return {"action": "delete_section_break", "description": "删除分节符", "success": False, "error": str(e)}


def _h_insert_section_break(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """插入分节符。"""
    if page_op is None:
        return {"action": "insert_section_break", "description": "插入分节符", "success": False, "error": "PageOperator 未初始化"}
    rng_param = params.get("rng")
    if rng_param:
        rng = _resolve_rng(op, rng_param)
    else:
        rng = op.selection
    break_type = params.get("type", "new_page")
    try:
        page_op.insert_section_break(rng, break_type)
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


def _h_set_section_start_type(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """设置节起始类型。"""
    if page_op is None:
        return {"action": "set_section_start_type", "description": "设置节起始类型", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    type_str = params.get("type", "new_page")
    try:
        page_op.set_section_start_type(idx, type_str)
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_set_section_start_new_page(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """节从新页开始。"""
    if page_op is None:
        return {"action": "set_section_start_new_page", "description": "节从新页开始", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.set_section_start_new_page(idx)
        return {"success": True}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_section_start_continuous(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """节设为连续。"""
    if page_op is None:
        return {"action": "set_section_start_continuous", "description": "连续节", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.set_section_start_continuous(idx)
        return {"success": True}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_first_page_different(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """设置首页不同。"""
    if page_op is None:
        return {"action": "set_first_page_different", "description": "首页不同", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    on = bool(params.get("on", True))
    try:
        page_op.set_first_page_different(idx, on)
        return {"success": True}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_odd_and_even_pages(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """设置奇偶页不同。"""
    if page_op is None:
        return {"action": "set_odd_and_even_pages", "description": "奇偶页不同", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    on = bool(params.get("on", True))
    try:
        page_op.set_odd_and_even_pages(idx, on)
        return {"success": True}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_insert_line_break(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """插入换行符。"""
    rng = _resolve_rng(op, params.get("rng"))
    rng.InsertBreak(Type=6)  # wdLineBreak = 6
    return {"success": True}


def _h_insert_column_break(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """插入栏分隔符。"""
    rng = _resolve_rng(op, params.get("rng"))
    rng.InsertBreak(Type=8)  # wdColumnBreak = 8
    return {"success": True}


def _h_remove_page_break(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """移除分页符。"""
    rng_param = params.get("rng")
    rng = _resolve_rng(op, rng_param) if rng_param else op.selection
    rng.Collapse(Direction=0)
    rng.MoveRight(Unit=1, Count=1)
    return {"success": True}


def _h_get_page_count(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """获取总页数。"""
    if page_op is None:
        return {"action": "get_page_count", "description": "获取总页数", "success": False, "error": "PageOperator 未初始化"}
    return {"success": True, "result": page_op.get_page_count()}


def _h_get_page_of_range(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """获取 Range 所在页码。"""
    if page_op is None:
        return {"action": "get_page_of_range", "description": "获取所在页码", "success": False, "error": "PageOperator 未初始化"}
    rng = _resolve_rng(op, params.get("rng"))
    return {"success": True, "result": page_op.get_page_of_range(rng)}


def _h_set_header(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """设置页眉。"""
    if page_op is None:
        return {"action": "set_header", "description": "设置页眉", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.set_header(
            idx,
            position=params.get("position", "primary"),
            text=params.get("text", ""),
            alignment=params.get("alignment", "left"),
        )
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_get_header(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """读取页眉。"""
    if page_op is None:
        return {"action": "get_header", "description": "读取页眉", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        return {"success": True, "result": page_op.get_header(idx, params.get("position", "primary"))}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_clear_header(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """清除页眉。"""
    if page_op is None:
        return {"action": "clear_header", "description": "清除页眉", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.clear_header(idx, params.get("position", "primary"))
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_set_footer(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """设置页脚。"""
    if page_op is None:
        return {"action": "set_footer", "description": "设置页脚", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.set_footer(
            idx,
            position=params.get("position", "primary"),
            text=params.get("text", ""),
            alignment=params.get("alignment", "left"),
        )
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_get_footer(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """读取页脚。"""
    if page_op is None:
        return {"action": "get_footer", "description": "读取页脚", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        return {"success": True, "result": page_op.get_footer(idx, params.get("position", "primary"))}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_clear_footer(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """清除页脚。"""
    if page_op is None:
        return {"action": "clear_footer", "description": "清除页脚", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.clear_footer(idx, params.get("position", "primary"))
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_insert_page_number_in_header(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """在页眉中插入页码。"""
    if page_op is None:
        return {"action": "insert_page_number_in_header", "description": "页眉插入页码", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.insert_page_number_in_header(
            idx,
            position=params.get("position", "primary"),
            alignment=params.get("alignment", "right"),
        )
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_insert_page_number_in_footer(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """在页脚中插入页码。"""
    if page_op is None:
        return {"action": "insert_page_number_in_footer", "description": "页脚插入页码", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.insert_page_number_in_footer(
            idx,
            position=params.get("position", "primary"),
            alignment=params.get("alignment", "center"),
        )
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_set_vertical_alignment(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """设置页面垂直对齐。"""
    if page_op is None:
        return {"action": "set_vertical_alignment", "description": "设置垂直对齐", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    align = params.get("align", "top")
    try:
        page_op.set_vertical_alignment(idx, align)
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_set_page_border(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """设置页面边框。"""
    if page_op is None:
        return {"action": "set_page_border", "description": "设置页面边框", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.set_page_border(
            idx,
            side=params.get("side", "all"),
            line_style=int(params.get("line_style", 1)),
            line_width=int(params.get("line_width", 6)),
            color=params.get("color", 0x000000),
        )
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_clear_page_border(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """清除页面边框。"""
    if page_op is None:
        return {"action": "clear_page_border", "description": "清除页面边框", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.clear_page_border(idx)
        return {"success": True}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_set_page_shading(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """设置整页背景色。"""
    if page_op is None:
        return {"action": "set_page_shading", "description": "设置页面背景", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.set_page_shading(idx, fill_color=params.get("fill_color", 0xCCE8FF))
        return {"success": True}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_clear_page_shading(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """清除页面背景色。"""
    if page_op is None:
        return {"action": "clear_page_shading", "description": "清除页面背景", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.clear_page_shading(idx)
        return {"success": True}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_apply_page_setup_to_all(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """将页面设置应用到所有节。"""
    if page_op is None:
        return {"action": "apply_page_setup_to_all", "description": "应用到全文", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        n = page_op.apply_page_setup_to_all(idx)
        return {"success": True, "result": n}
    except (IndexError, ValueError) as e:
        return {"success": False, "error": str(e)}


def _h_copy_page_setup(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """复制页面设置。"""
    if page_op is None:
        return {"action": "copy_page_setup", "description": "复制页面设置", "success": False, "error": "PageOperator 未初始化"}
    from_idx = int(params.get("from_index", 1))
    to_idx = int(params.get("to_index", 2))
    try:
        ok = page_op.copy_page_setup(from_idx, to_idx)
        return {"success": ok}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_reset_page_setup(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """重置页面设置。"""
    if page_op is None:
        return {"action": "reset_page_setup", "description": "重置页面设置", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        page_op.reset_page_setup(idx)
        return {"success": True}
    except IndexError as e:
        return {"success": False, "error": str(e)}


# ── Action 注册表 ─────────────────────────────────────────────────────────────
# target: ("op_attr", None) 表示 op.attr.method()
#         ("op_attr", "sub_attr") 表示 op.attr.sub_attr.method()
# rng: True → 该 action 接受 rng 参数，由注册表自动解析后传入

ACTION_REGISTRY: Dict[str, Dict[str, Any]] = {

    # ── 文本读取 ──────────────────────────────────────────────
    "get_full_text":       {"handler": _h_get_full_text},
    "get_text":            {"handler": _h_get_text},
    "get_selection_text":   {"handler": _h_get_selection_text},
    "get_paragraph_text":  {"handler": _h_get_paragraph_text},
    "get_paragraph_count": {"handler": _h_get_paragraph_count},
    "apply_format_to_range": {"handler": _h_apply_format_to_range},

    # ── 查找 ──────────────────────────────────────────────────
    "find":                {"handler": _h_find},
    "find_all":            {"handler": _h_find_all},
    "count_occurrences":   {"handler": _h_count_occurrences},
    "find_wildcards":      {"handler": _h_find_wildcards},
    "find_with_format":    {"handler": _h_find_with_format},

    # ── 替换 ─────────────────────────────────────────────────
    "replace":             {"handler": _h_replace,                                   "rng": False, "capture": STATE_CONTENT},
    "replace_with_format": {"handler": _h_replace_with_format,                      "rng": False, "capture": STATE_CONTENT},
    "batch_replace":      {"handler": _h_batch_replace,                              "rng": False, "capture": STATE_CONTENT},

    # ── 字体格式（op 直接调用） ────────────────────────────────
    # void COM 调用，无返回值；若用 RESULT_BOOL，bool(None) 恒为 False 会误报失败
    "set_bold":               {"target": None,               "method": "set_bold",             "rng": True,  "params": {"bold": "bold"},              "result": RESULT_NONE,  "capture": STATE_FONT},
    "set_italic":             {"target": None,               "method": "set_italic",           "rng": True,  "params": {"italic": "italic"},            "result": RESULT_NONE,  "capture": STATE_FONT},
    "set_underline":          {"target": None,               "method": "set_underline",       "rng": True,  "params": {"underline": "underline"},    "result": RESULT_NONE,  "capture": STATE_FONT},
    "set_font_color":         {"target": None,              "method": "set_font_color",      "rng": True,  "params": {"color": "color"},             "result": RESULT_NONE,  "capture": STATE_FONT},
    "set_font_name":          {"target": None,              "method": "set_font_name",        "rng": True,  "params": {"font_name": "name"},         "result": RESULT_NONE,  "capture": STATE_FONT},
    "set_font_size":          {"target": None,              "method": "set_font_size",        "rng": True,  "params": {"size": "size"},               "result": RESULT_NONE,  "capture": STATE_FONT},
    "set_highlight":          {"target": None,              "method": "set_highlight",        "rng": True,  "params": {"highlight": "highlight"},     "result": RESULT_NONE,  "capture": STATE_FONT},
    # 注意：set_paragraph_alignment 内部调用的是 rng.Bold/rng.Italic 等字体属性，
    # 不是 ParagraphFormat，故 capture 应为 STATE_FONT（不能让 LLM 误以为改的是段落格式）
    "set_paragraph_alignment":{"target": None,              "method": "set_paragraph_alignment","rng": True,"params": {"alignment": "align"},        "result": RESULT_NONE,  "capture": STATE_FONT},

    # ── 段落格式（fmt 调用） ────────────────────────────────
    "set_line_spacing":        {"target": ("fmt", None),     "method": "set_line_spacing",       "rng": True,  "params": {"spacing": "value", "rule": "rule"},                "result": RESULT_NONE,  "capture": STATE_PARA},
    "set_indent_left":         {"target": ("fmt", None),     "method": "set_indent_left",        "rng": True,  "params": {"characters": "characters", "cm": "cm"},         "result": RESULT_NONE,  "capture": STATE_PARA},
    "set_indent_right":        {"target": ("fmt", None),     "method": "set_indent_right",       "rng": True,  "params": {"characters": "characters", "cm": "cm"},         "result": RESULT_NONE,  "capture": STATE_PARA},
    "set_first_line_indent":   {"target": ("fmt", None),     "method": "set_first_line_indent",  "rng": True,  "params": {"characters": "characters", "cm": "cm"},         "result": RESULT_NONE,  "capture": STATE_PARA},
    "set_space_before":        {"target": ("fmt", None),     "method": "set_space_before_para",  "rng": True,  "params": {"points": "points"},                                          "result": RESULT_NONE,  "capture": STATE_PARA},
    "set_space_after":         {"target": ("fmt", None),     "method": "set_space_after_para",   "rng": True,  "params": {"points": "points"},                                          "result": RESULT_NONE,  "capture": STATE_PARA},
    "set_outline_level":       {"target": ("fmt", None),     "method": "set_outline_level",      "rng": True,  "params": {"level": "level"},                                            "result": RESULT_NONE,  "capture": STATE_PARA},
    "set_keep_together":       {"target": ("fmt", None),     "method": "set_keep_together",      "rng": True,  "params": {"on": "on"},                                               "result": RESULT_NONE,  "capture": STATE_PARA},
    "set_keep_with_next":      {"target": ("fmt", None),     "method": "set_keep_with_next",     "rng": True,  "params": {"on": "on"},                                               "result": RESULT_NONE,  "capture": STATE_PARA},

    # ── 边框与底纹（fmt 调用） ──────────────────────────────
    "set_border":             {"target": ("fmt", None),      "method": "set_border",            "rng": True,  "params": {"side": "side", "line_style": "line_style", "line_width": "line_width", "color": "color"}, "result": RESULT_NONE,  "capture": STATE_BORDER},
    "clear_border":           {"target": ("fmt", None),      "method": "clear_border",          "rng": True,  "params": {}},
    "set_shading":            {"target": ("fmt", None),      "method": "set_shading",            "rng": True,  "params": {"fill_color": "fill_color", "texture": "texture"},        "result": RESULT_NONE,  "capture": STATE_BORDER},
    "clear_shading":          {"target": ("fmt", None),      "method": "clear_shading",          "rng": True,  "params": {}},

    # ── Range 导航（op 直接调用） ──────────────────────────
    "expand_to_word":         {"target": None,              "method": "expand_to_word",        "rng": True,  "params": {}},
    "expand_to_sentence":     {"target": None,              "method": "expand_to_sentence",    "rng": True,  "params": {}},
    "expand_to_paragraph":    {"target": None,              "method": "expand_to_paragraph",   "rng": True,  "params": {}},
    "collapse":               {"target": None,              "method": "collapse",              "rng": True,  "params": {"direction": "direction"}},
    "move":                   {"target": None,              "method": "move",                  "rng": True,  "params": {"unit": "unit", "count": "count"}},
    "goto_bookmark":        {"handler": _h_goto_bookmark},
    "goto_page":            {"handler": _h_goto_page},
    "goto_line":           {"handler": _h_goto_line},

    # ── 书签（op / op.bm 调用） ───────────────────────────────
    "create_bookmark":     {"handler": _h_create_bookmark},
    "delete_bookmark":        {"target": ("bm", None),      "method": "delete",             "rng": False, "params": {"name": "name"},                                              "result": RESULT_BOOL},
    "delete_all_bookmarks":   {"target": ("bm", None),      "method": "delete_all",         "rng": False, "params": {}},
    "rename_bookmark":        {"target": ("bm", None),      "method": "rename",            "rng": False, "params": {"old_name": "old_name", "new_name": "new_name"},           "result": RESULT_BOOL},
    "list_bookmarks":      {"handler": _h_list_bookmarks},
    "export_bookmarks":    {"handler": _h_export_bookmarks},
    "import_bookmarks":    {"handler": _h_import_bookmarks},
    "bookmark_text":       {"handler": _h_bookmark_text},
    "wrap_with_bookmarks": {"handler": _h_wrap_with_bookmarks},

    # ── 文本操作 ──────────────────────────────────────────────
    "insert_text":         {"handler": _h_insert_text},
    "insert_page_break":   {"handler": _h_insert_page_break},
    "insert_file":         {"handler": _h_insert_file},
    "insert_symbol":       {"handler": _h_insert_symbol},
    "insert_paragraph":    {"handler": _h_insert_paragraph},
    "delete_range":        {"handler": _h_delete_range,                                  "rng": True,  "capture": STATE_CONTENT},
    "delete_selection":       {"target": None,              "method": "delete_selection",  "rng": False, "params": {}},
    "clear_range":            {"handler": _h_clear_range,                                  "rng": True,  "capture": STATE_CONTENT},
    "to_uppercase":          {"target": None,              "method": "to_uppercase",     "rng": True,  "params": {},                    "capture": STATE_CONTENT},
    "to_lowercase":          {"target": None,              "method": "to_lowercase",     "rng": True,  "params": {},                    "capture": STATE_CONTENT},
    "to_title_case":         {"target": None,              "method": "to_title_case",     "rng": True,  "params": {},                    "capture": STATE_CONTENT},
    "select":              {"handler": _h_select},

    # ── 统计 ──────────────────────────────────────────────────
    "char_count":          {"handler": _h_char_count},
    "word_count":          {"handler": _h_word_count},
    "sentence_count":      {"handler": _h_sentence_count},
    "paragraph_count":     {"handler": _h_paragraph_count},

    # ── 文档操作 ──────────────────────────────────────────────
    "new_document":        {"handler": _h_new_document},
    "save":               {"handler": _h_save},

    # ── Selection 专属操作（sel 调用） ─────────────────────
    "move_left":             {"target": ("sel", None),     "method": "move_left",            "rng": False, "params": {"unit": "unit", "count": "count", "extend": "extend"}},
    "move_right":            {"target": ("sel", None),     "method": "move_right",           "rng": False, "params": {"unit": "unit", "count": "count", "extend": "extend"}},
    "move_up":               {"target": ("sel", None),     "method": "move_up",              "rng": False, "params": {"unit": "unit", "count": "count", "extend": "extend"}},
    "move_down":             {"target": ("sel", None),     "method": "move_down",            "rng": False, "params": {"unit": "unit", "count": "count", "extend": "extend"}},
    "move_to_line_start":    {"target": ("sel", None),     "method": "move_to_line_start",   "rng": False, "params": {}},
    "move_to_line_end":      {"target": ("sel", None),     "method": "move_to_line_end",     "rng": False, "params": {}},
    "move_to_document_start":{"target": ("sel", None),     "method": "move_to_document_start","rng": False, "params": {}},
    "move_to_document_end":  {"target": ("sel", None),     "method": "move_to_document_end", "rng": False, "params": {}},
    "select_word":           {"target": ("sel", None),     "method": "select_word",          "rng": False, "params": {}},
    "select_line":           {"target": ("sel", None),     "method": "select_line",          "rng": False, "params": {}},
    "select_paragraph":      {"target": ("sel", None),     "method": "select_paragraph",     "rng": False, "params": {}},
    "select_all":            {"target": ("sel", None),     "method": "select_all",           "rng": False, "params": {}},
    "type_text":             {"target": ("sel", None),     "method": "type_text",            "rng": False, "params": {"text": "text"}},
    "clear_formatting":      {"target": ("sel", None),     "method": "clear_formatting",    "rng": False, "params": {}},
    "get_selection_info":    {"handler": _h_get_selection_info},
    "find_and_select":       {"target": ("sel", None),     "method": "find_and_select",      "rng": False, "params": {"text": "text", "whole_word": "whole_word", "match_case": "match_case"}, "result": RESULT_BOOL},
    "replace_in_selection":  {"target": ("sel", None),     "method": "replace_selection",    "rng": False, "params": {"find_text": "find_text", "replace_text": "replace_text", "replace_all": "replace_all"}, "result": RESULT_SINGLE},

    # ── word-paragraph-operator 段落操作 ──────────────────────────────
    # 段落基础访问
    "get_paragraph_count":       {"handler": _h_get_paragraph_count},
    "get_paragraph_by_index":    {"handler": _h_get_paragraph_by_index},
    # 注意：get_paragraph_text 在 word-text-operator 中已有（rng 参数），此处使用 para_op 版本（index 参数）
    "get_paragraph_range":       {"handler": _h_get_paragraph_range},
    # 段落识别与结构分析
    "get_document_structure":    {"handler": _h_get_document_structure},
    "get_outline_summary":       {"handler": _h_get_outline_summary},
    "find_empty_paragraphs":     {"handler": _h_find_empty_paragraphs},
    "find_heading_paragraphs":  {"handler": _h_find_heading_paragraphs},
    "find_paragraphs_by_level":  {"handler": _h_find_paragraphs_by_level},
    "find_paragraphs_by_text":   {"handler": _h_find_paragraphs_by_text},
    # 段落属性读取
    "get_paragraph_format_info": {"handler": _h_get_paragraph_format_info},
    "get_paragraph_style":       {"handler": _h_get_paragraph_style},
    "get_outline_level":         {"handler": _h_get_outline_level},
    "is_paragraph_list_item":    {"handler": _h_is_paragraph_list_item},
    "is_paragraph_in_table":     {"handler": _h_is_paragraph_in_table},
    # 段落属性写入（单个）
    "set_paragraph_alignment":   {"handler": _h_set_paragraph_alignment},
    "set_paragraph_line_spacing": {"handler": _h_set_paragraph_line_spacing},
    "set_paragraph_space_before": {"handler": _h_set_paragraph_space_before},
    "set_paragraph_space_after": {"handler": _h_set_paragraph_space_after},
    "set_paragraph_indent_left": {"handler": _h_set_paragraph_indent_left},
    "set_paragraph_indent_right": {"handler": _h_set_paragraph_indent_right},
    "set_paragraph_first_line_indent": {"handler": _h_set_paragraph_first_line_indent},
    "set_paragraph_hanging_indent": {"handler": _h_set_paragraph_hanging_indent},
    "set_paragraph_outline_level": {"handler": _h_set_paragraph_outline_level},
    "set_paragraph_keep_together": {"handler": _h_set_paragraph_keep_together},
    "set_paragraph_keep_with_next": {"handler": _h_set_paragraph_keep_with_next},
    "set_paragraph_style":       {"handler": _h_set_paragraph_style},
    "reset_paragraph_format":    {"handler": _h_reset_paragraph_format},
    # 段落边框与底纹
    "set_paragraph_border":      {"handler": _h_set_paragraph_border},
    "clear_paragraph_border":    {"handler": _h_clear_paragraph_border},
    "set_paragraph_shading":     {"handler": _h_set_paragraph_shading},
    "clear_paragraph_shading":   {"handler": _h_clear_paragraph_shading},
    # 编号列表操作
    "get_list_paragraphs":      {"handler": _h_get_list_paragraphs},
    "get_list_level":            {"handler": _h_get_list_level},
    "set_list_level":            {"handler": _h_set_list_level},
    "apply_bullet_list":         {"handler": _h_apply_bullet_list},
    "apply_numbered_list":       {"handler": _h_apply_numbered_list},
    "remove_list_format":        {"handler": _h_remove_list_format},
    # 段落内容操作（CRUD）
    "set_paragraph_text":        {"handler": _h_set_paragraph_text},
    "insert_text_before_paragraph": {"handler": _h_insert_text_before_paragraph},
    "insert_text_after_paragraph": {"handler": _h_insert_text_after_paragraph},
    "delete_paragraph":          {"handler": _h_delete_paragraph},
    "clear_paragraph":           {"handler": _h_clear_paragraph},
    "add_paragraph_after":       {"handler": _h_add_paragraph_after},
    "add_paragraph_before":      {"handler": _h_add_paragraph_before},
    "merge_with_next":            {"handler": _h_merge_with_next},
    "merge_with_previous":        {"handler": _h_merge_with_previous},
    "split_paragraph_by_separator": {"handler": _h_split_paragraph_by_separator},
    # 批量操作
    "reverse_paragraph_order":   {"handler": _h_reverse_paragraph_order},
    "delete_empty_paragraphs":   {"handler": _h_delete_empty_paragraphs},
    # Selection / Range 互操作
    "get_paragraph_at_selection": {"handler": _h_get_paragraph_at_selection},
    "select_paragraph":           {"handler": _h_select_paragraph},
    "select_paragraph_range":    {"handler": _h_select_paragraph_range},

    # ── word-page-operator 页面操作 ──────────────────────────────────────────
    # 节基础访问
    "get_section_count":             {"handler": _h_get_section_count},
    "get_section_by_index":          {"handler": _h_get_section_by_index},
    "get_current_section_index":     {"handler": _h_get_current_section_index},
    # 页面设置读取
    "get_page_setup_info":           {"handler": _h_get_page_setup_info},
    "get_page_margins":              {"handler": _h_get_page_margins},
    "get_paper_size":               {"handler": _h_get_paper_size},
    "get_orientation":              {"handler": _h_get_orientation},
    "get_column_count":             {"handler": _h_get_column_count},
    "get_column_info":              {"handler": _h_get_column_info},
    "get_section_start_type":       {"handler": _h_get_section_start_type},
    # 页边距设置
    "set_page_margins":             {"handler": _h_set_page_margins},
    "set_page_margins_by_inch":     {"handler": _h_set_page_margins_by_inch},
    "set_page_margins_preset":      {"handler": _h_set_page_margins_preset},
    # 纸张设置
    "set_paper_size":               {"handler": _h_set_paper_size},
    "set_paper_size_preset":       {"handler": _h_set_paper_size_preset},
    "set_orientation":              {"handler": _h_set_orientation},
    # 分栏操作
    "set_columns":                  {"handler": _h_set_columns},
    "set_columns_with_gutter":      {"handler": _h_set_columns_with_gutter},
    "set_columns_equal_width":      {"handler": _h_set_columns_equal_width},
    "set_column_width":             {"handler": _h_set_column_width},
    "apply_two_column_layout":      {"handler": _h_apply_two_column_layout},
    # 分节符
    "delete_section_break":          {"handler": _h_delete_section_break},
    "insert_section_break":         {"handler": _h_insert_section_break},
    "set_section_start_type":       {"handler": _h_set_section_start_type},
    "set_section_start_new_page":   {"handler": _h_set_section_start_new_page},
    "set_section_start_continuous":  {"handler": _h_set_section_start_continuous},
    "set_first_page_different":     {"handler": _h_set_first_page_different},
    "set_odd_and_even_pages":       {"handler": _h_set_odd_and_even_pages},
    # 分页控制
    "insert_line_break":           {"handler": _h_insert_line_break},
    "insert_column_break":          {"handler": _h_insert_column_break},
    "remove_page_break":           {"handler": _h_remove_page_break},
    "get_page_count":               {"handler": _h_get_page_count},
    "get_page_of_range":            {"handler": _h_get_page_of_range},
    # 页眉页脚
    "set_header":                   {"handler": _h_set_header},
    "get_header":                   {"handler": _h_get_header},
    "clear_header":                 {"handler": _h_clear_header},
    "set_footer":                   {"handler": _h_set_footer},
    "get_footer":                   {"handler": _h_get_footer},
    "clear_footer":                 {"handler": _h_clear_footer},
    "insert_page_number_in_header": {"handler": _h_insert_page_number_in_header},
    "insert_page_number_in_footer": {"handler": _h_insert_page_number_in_footer},
    # 页面级格式
    "set_vertical_alignment":       {"handler": _h_set_vertical_alignment},
    "set_page_border":              {"handler": _h_set_page_border},
    "clear_page_border":            {"handler": _h_clear_page_border},
    "set_page_shading":             {"handler": _h_set_page_shading},
    "clear_page_shading":           {"handler": _h_clear_page_shading},
    # 文档级操作
    "apply_page_setup_to_all":      {"handler": _h_apply_page_setup_to_all},
    "copy_page_setup":              {"handler": _h_copy_page_setup},
    "reset_page_setup":             {"handler": _h_reset_page_setup},
}


# ── word-page-operator 页面操作 handlers ──────────────────────────────────────

def _h_get_section_count(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """节总数。"""
    if page_op is None:
        return {"action": "get_section_count", "description": "获取节总数", "success": False, "error": "PageOperator 未初始化"}
    return {"success": True, "result": page_op.count()}


def _h_get_section_by_index(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """按索引获取节。"""
    if page_op is None:
        return {"action": "get_section_by_index", "description": "按索引获取节", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        sec = page_op.get(idx)
        return {"success": True, "result": page_op.get_index(sec)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_current_section_index(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """获取光标所在节的索引。"""
    if page_op is None:
        return {"action": "get_current_section_index", "description": "获取光标所在节", "success": False, "error": "PageOperator 未初始化"}
    return {"success": True, "result": page_op.section.get_current_section_index()}


def _h_get_page_setup_info(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """读取完整页面设置信息。"""
    if page_op is None:
        return {"action": "get_page_setup_info", "description": "读取页面设置", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        return {"success": True, "result": page_op.get_page_setup_info(idx)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_page_margins(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """读取页边距。"""
    if page_op is None:
        return {"action": "get_page_margins", "description": "读取页边距", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        return {"success": True, "result": page_op.get_page_margins(idx)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_paper_size(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """读取纸张大小。"""
    if page_op is None:
        return {"action": "get_paper_size", "description": "读取纸张大小", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        return {"success": True, "result": page_op.get_paper_size(idx)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_orientation(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """读取纸张方向。"""
    if page_op is None:
        return {"action": "get_orientation", "description": "读取纸张方向", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        return {"success": True, "result": page_op.get_orientation(idx)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


def _h_get_column_count(action: Dict, params: Dict, op: "WordTextOperator", page_op=None) -> Dict[str, Any]:
    """读取分栏数。"""
    if page_op is None:
        return {"action": "get_column_count", "description": "读取分栏数", "success": False, "error": "PageOperator 未初始化"}
    idx = int(params.get("index", 1))
    try:
        return {"success": True, "result": page_op.get_column_count(idx)}
    except IndexError as e:
        return {"success": False, "error": str(e)}


# ── 分发器 ─────────────────────────────────────────────────────────────────────

def execute_action(action: Dict[str, Any], op: Optional["WordTextOperator"] = None,
                  para_op=None, page_op=None) -> Dict[str, Any]:
    """
    根据 action["action"] 从注册表查到映射，构造参数，调用对应方法。

    新增 action 只需在 ACTION_REGISTRY 中添加一条，
    或定义一个新的 handler 函数并注册，无需修改本文件以外的代码。
    """
    import logging as _ar_logger
    _ar_logger.getLogger(__name__).info(
        "[execute_action] 进入 | action=%s params=%s",
        action.get("action", ""),
        action.get("params", {}),
    )
    action_type = action.get("action", "")
    params = action.get("params", {}) or {}
    desc = action.get("description", action_type)

    spec = ACTION_REGISTRY.get(action_type)

    # ── 有自定义 handler ────────────────────────────────────────
    if spec and "handler" in spec:
        try:
            state_type = spec.get("capture")
            before_state_raw = None
            before_state_text = ""
            # 优先用 handler 自己返回的 before_state（handler 知道如何正确捕获自身状态）。
            # 兜底：若 handler 未自行捕获，由 execute_action 代为捕获。
            if state_type:
                rng = _resolve_rng(op, params.get("rng", ""))
                before_state_raw = _capture_state(state_type, rng, op)
                before_state_text = _human_readable_state(state_type, before_state_raw) if before_state_raw else ""

            sig = inspect.signature(spec["handler"])
            handler_kwargs = {}
            if "para_op" in sig.parameters:
                handler_kwargs["para_op"] = para_op
            if "page_op" in sig.parameters:
                handler_kwargs["page_op"] = page_op
            result = spec["handler"](action, params, op, **handler_kwargs)
            result.setdefault("success", True)
            result.setdefault("action", action_type)
            result.setdefault("description", desc)
            # handler 若已自行捕获了 before_state，保留之；否则用上方兜底捕获的值
            if not result.get("before_state") and before_state_text:
                result["before_state"] = before_state_text
            if not result.get("before_state_raw") and before_state_raw:
                result["before_state_raw"] = before_state_raw
            import logging as _ar_logger
            _ar_logger.getLogger(__name__).info(
                "[execute_action] handler 完成返回 | action=%s success=%s result=%s",
                action_type, result.get("success"), str(result.get("result", ""))[:100],
            )
            return result
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {"action": action_type, "description": desc, "success": False, "error": str(e)}

    # ── 有映射 spec ────────────────────────────────────────────
    if spec is None:
        return {"action": action_type, "description": desc, "success": False,
                "error": f"未知的 action: {action_type}"}

    try:
        target_attr, sub_attr = spec.get("target") or (None, None)  # e.g. ("fmt", None) or None
        method_name = spec["method"]
        needs_rng = spec.get("rng", False)
        param_map = spec.get("params", {})

        # 解析 target 对象：None → 直接用 op；("fmt", None) → op.fmt
        if target_attr is None:
            target = op
        else:
            target = getattr(op, target_attr)
            if sub_attr:
                target = getattr(target, sub_attr)

        # ── 执行前捕获初始状态 ──────────────────────────────
        # _resolve_rng 内部已处理无 rng 参数时的兜底（优先当前 Selection），
        # 故 state capture 不依赖 needs_rng——即使 rng=False，只要注册了 capture 就应捕获。
        state_type = spec.get("capture")
        before_state_raw = None
        before_state_text = ""
        rng = None
        if state_type or needs_rng:
            rng = _resolve_rng(op, params.get("rng", ""))
        if state_type:
            before_state_raw = _capture_state(state_type, rng, op)
            before_state_text = _human_readable_state(state_type, before_state_raw) if before_state_raw else ""

        # 构造调用参数
        kwargs = {}
        for action_param, method_param in param_map.items():
            raw_val = params.get(action_param)
            if raw_val is not None:
                # 自动类型转换
                if method_param in ("size", "spacing", "points", "level"):
                    raw_val = float(raw_val)
                elif method_param in ("unit", "count", "line_style", "line_width"):
                    raw_val = int(raw_val)
                elif method_param == "on":
                    raw_val = bool(raw_val)
                kwargs[method_param] = raw_val

        # 注入 rng
        if needs_rng and rng is not None:
            kwargs["rng"] = rng

        # 调用
        ret = getattr(target, method_name)(**kwargs)

        # 构建返回值
        result_mode = spec.get("result", RESULT_NONE)
        if result_mode == RESULT_NONE:
            ret_dict = {"action": action_type, "description": desc, "success": True}
        elif result_mode == RESULT_SINGLE:
            ret_dict = {"action": action_type, "description": desc, "success": True, "result": ret}
        elif result_mode == RESULT_BOOL:
            ret_dict = {"action": action_type, "description": desc, "success": bool(ret)}
        else:
            ret_dict = {"action": action_type, "description": desc, "success": True, "result": ret}

        # 附加初始状态（供历史记录使用）
        if before_state_text:
            ret_dict["before_state"] = before_state_text
        if before_state_raw:
            ret_dict["before_state_raw"] = before_state_raw

        return ret_dict

    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"action": action_type, "description": desc, "success": False, "error": str(e)}