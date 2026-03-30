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
      - "" / None / 缺失参数       → 当前 Selection（默认，优先使用！）
      - "[start, end]"             → 指定字符范围
      - 其他                        → 当前 Selection
    """
    # Word 未连接（connect 失败或从未调用）
    if op._base._word_app is None:
        raise RuntimeError("Word 未连接（请确保已打开 Word 文档）")

    # 优先使用当前 Selection（用户鼠标选中的区域）
    if rng_param != "full_document":
        try:
            sel = op._base._word_app.Selection
            if sel is not None and sel.Start != sel.End:
                return sel
        except Exception:
            pass

    # 显式要求全文档时才用 Content
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

    # 坐标格式 [start, end]
    if isinstance(rng_param, str) and rng_param.startswith("["):
        import ast
        try:
            coords = ast.literal_eval(rng_param)
            if isinstance(coords, (list, tuple)) and len(coords) == 2:
                return op.get_range(int(coords[0]), int(coords[1]))
        except Exception:
            pass

    # 无参数或非标准参数：尝试用全文档代替（用户未选中文本时）
    doc = op._base._document
    if doc is None:
        try:
            doc = op._base._word_app.ActiveDocument
        except Exception:
            pass
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

def _h_find(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    """find：返回 {start, end, text} 或 None"""
    rng = op.find(
        params.get("text", ""),
        whole_word=params.get("whole_word", False),
        match_case=params.get("match_case", False),
    )
    result = {
        "success": rng is not None,
        "result": {"start": rng.Start, "end": rng.End, "text": rng.Text} if rng else None
    }
    return result


def _h_count_occurrences(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    count = op.count_occurrences(params.get("text", ""))
    return {"success": True, "result": count}


def _h_find_wildcards(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    pattern = params.get("pattern", "")
    replace_text = params.get("replace_text")
    if replace_text:
        n = op.find_wildcards(pattern, replace_text)
    else:
        n = 1 if op.find_wildcards(pattern) else 0
    return {"success": True, "result": n}


def _h_goto_bookmark(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    rng = op.go_to_bookmark(params.get("name", ""))
    if rng:
        op.select(rng)
    return {"success": rng is not None}


def _h_goto_page(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
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


def _h_get_full_text(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    return {"success": True, "result": op.get_full_text()}


def _h_get_text(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    rng = _resolve_rng(op, params.get("rng", ""))
    return {"success": True, "result": rng.Text}


def _h_get_selection_text(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    return {"success": True, "result": op.get_selection_text()}


def _h_find_all(action: Dict, params: Dict, op: "WordTextOperator") -> Dict[str, Any]:
    positions = op.find_all(params.get("text", ""))
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
    return {"success": True, "result": op.sel.get_selection_info()}


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


# ── Action 注册表 ─────────────────────────────────────────────────────────────
# target: ("op_attr", None) 表示 op.attr.method()
#         ("op_attr", "sub_attr") 表示 op.attr.sub_attr.method()
# rng: True → 该 action 接受 rng 参数，由注册表自动解析后传入

ACTION_REGISTRY: Dict[str, Dict[str, Any]] = {

    # ── 文本读取 ──────────────────────────────────────────────
    "get_full_text":       {"handler": _h_get_full_text},
    "get_text":            {"handler": _h_get_text},
    "get_selection_text":   {"handler": _h_get_selection_text},

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
}


# ── 分发器 ─────────────────────────────────────────────────────────────────────

def execute_action(action: Dict[str, Any], op: Optional["WordTextOperator"] = None) -> Dict[str, Any]:
    """
    根据 action["action"] 从注册表查到映射，构造参数，调用对应方法。

    新增 action 只需在 ACTION_REGISTRY 中添加一条，
    或定义一个新的 handler 函数并注册，无需修改本文件以外的代码。
    """
    action_type = action.get("action", "")
    params = action.get("params", {}) or {}
    desc = action.get("description", action_type)

    spec = ACTION_REGISTRY.get(action_type)

    # ── 有自定义 handler ────────────────────────────────────────
    if spec and "handler" in spec:
        try:
            # handler 也支持 capture
            state_type = spec.get("capture")
            rng = None
            before_state_raw = None
            before_state_text = ""
            if state_type and spec.get("rng"):
                rng = _resolve_rng(op, params.get("rng", ""))
                before_state_raw = _capture_state(state_type, rng, op)
                before_state_text = _human_readable_state(state_type, before_state_raw) if before_state_raw else ""

            result = spec["handler"](action, params, op)
            result.setdefault("success", True)
            if before_state_text:
                result["before_state"] = before_state_text
            if before_state_raw:
                result["before_state_raw"] = before_state_raw
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
        state_type = spec.get("capture")
        rng = None
        before_state_raw = None
        before_state_text = ""
        if state_type and needs_rng:
            rng = _resolve_rng(op, params.get("rng", ""))
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

        # 注入 rng（已在上方解析过）
        if needs_rng:
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