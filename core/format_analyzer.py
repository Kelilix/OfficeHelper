"""
格式分析器模块
分析Word文档格式问题
"""

from typing import List, Dict, Any, Optional
from dataclasses import dataclass, field
from enum import Enum


class FormatIssueType(Enum):
    """格式问题类型"""
    FONT_NAME = "font_name"          # 字体不一致
    FONT_SIZE = "font_size"         # 字号不统一
    PARAGRAPH_ALIGNMENT = "alignment" # 对齐问题
    LINE_SPACING = "line_spacing"   # 行距问题
    INDENT = "indent"               # 缩进问题
    MARGIN = "margin"               # 边距问题
    PAGE_SIZE = "page_size"         # 页面尺寸问题
    ORIENTATION = "orientation"     # 方向问题
    STYLE = "style"                 # 样式问题
    SPELLING = "spelling"           # 拼写错误
    EMPTY_PARAGRAPH = "empty_para"  # 空段落
    PAGE_BREAK = "page_break"       # 分页问题


@dataclass
class FormatIssue:
    """格式问题"""
    issue_type: FormatIssueType
    message: str
    severity: str = "warning"  # error/warning/info
    location: str = ""  # 位置描述
    suggested_fix: str = ""
    related_paragraph: Optional[int] = None


@dataclass
class FormatAnalysisResult:
    """格式分析结果"""
    document_path: str
    total_issues: int = 0
    errors: int = 0
    warnings: int = 0
    infos: int = 0
    issues: List[FormatIssue] = field(default_factory=list)
    summary: Dict[str, Any] = field(default_factory=dict)


class FormatAnalyzer:
    """格式分析器"""

    def __init__(self):
        self._common_fonts = {
            '中文': ['微软雅黑', '宋体', '黑体', '楷体', '仿宋', '华文细黑', '苹方', '思源黑体'],
            '英文': ['Arial', 'Times New Roman', 'Calibri', 'Helvetica', 'Georgia', 'Verdana']
        }

    def analyze_document(self, document_info: dict, paragraphs: List[dict]) -> FormatAnalysisResult:
        """
        分析文档格式

        Args:
            document_info: 文档信息
            paragraphs: 段落列表

        Returns:
            FormatAnalysisResult: 分析结果
        """
        result = FormatAnalysisResult(
            document_path=document_info.get('path', '')
        )

        # 分析各项格式
        self._analyze_font_consistency(paragraphs, result)
        self._analyze_paragraph_alignment(paragraphs, result)
        self._analyze_line_spacing(paragraphs, result)
        self._analyze_empty_paragraphs(paragraphs, result)
        self._analyze_heading_styles(paragraphs, result)

        # 统计
        result.total_issues = len(result.issues)
        for issue in result.issues:
            if issue.severity == "error":
                result.errors += 1
            elif issue.severity == "warning":
                result.warnings += 1
            else:
                result.infos += 1

        # 生成摘要
        result.summary = self._generate_summary(result)

        return result

    def _analyze_font_consistency(self, paragraphs: List[dict], result: FormatAnalysisResult):
        """分析字体一致性"""
        fonts = {}
        for i, para in enumerate(paragraphs):
            if not para.get('text'):
                continue
            font_name = para.get('font_name', 'Unknown')
            fonts[font_name] = fonts.get(font_name, 0) + 1

        if len(fonts) > 1:
            # 字体不统一
            sorted_fonts = sorted(fonts.items(), key=lambda x: x[1], reverse=True)
            main_font = sorted_fonts[0][0]
            other_fonts = [f for f, c in sorted_fonts[1:]]

            result.issues.append(FormatIssue(
                issue_type=FormatIssueType.FONT_NAME,
                severity="warning",
                message=f"文档使用了多种字体: {', '.join(fonts.keys())}",
                location="全文",
                suggested_fix=f"建议统一使用{main_font}字体",
                related_paragraph=None
            ))

    def _analyze_paragraph_alignment(self, paragraphs: List[dict], result: FormatAnalysisResult):
        """分析段落对齐"""
        alignments = {}
        for i, para in enumerate(paragraphs):
            if not para.get('text'):
                continue
            align = para.get('alignment', 'left')
            alignments[align] = alignments.get(align, 0) + 1

        # 检测混合对齐
        mixed_aligns = [a for a, c in alignments.items() if c < len(paragraphs) * 0.1]
        if len(mixed_aligns) > 1:
            result.issues.append(FormatIssue(
                issue_type=FormatIssueType.PARAGRAPH_ALIGNMENT,
                severity="info",
                message=f"段落对齐方式不统一",
                location="全文",
                suggested_fix="建议统一对齐方式",
                related_paragraph=None
            ))

    def _analyze_line_spacing(self, paragraphs: List[dict], result: FormatAnalysisResult):
        """分析行距"""
        spacings = {}
        for i, para in enumerate(paragraphs):
            if not para.get('text'):
                continue
            spacing = round(para.get('line_spacing', 12) / 12, 1)
            spacings[spacing] = spacings.get(spacing, 0) + 1

        if len(spacings) > 1:
            main_spacing = max(spacings.items(), key=lambda x: x[1])[0]
            result.issues.append(FormatIssue(
                issue_type=FormatIssueType.LINE_SPACING,
                severity="warning",
                message=f"行距不统一，主要行距为{main_spacing}倍",
                location="全文",
                suggested_fix=f"建议统一设置行距为{main_spacing}倍",
                related_paragraph=None
            ))

    def _analyze_empty_paragraphs(self, paragraphs: List[dict], result: FormatAnalysisResult):
        """分析空段落"""
        empty_count = sum(1 for p in paragraphs if not p.get('text', '').strip())
        if empty_count > 5:
            result.issues.append(FormatIssue(
                issue_type=FormatIssueType.EMPTY_PARAGRAPH,
                severity="info",
                message=f"发现{empty_count}个空段落",
                location="全文",
                suggested_fix="建议删除空段落以优化文档结构",
                related_paragraph=None
            ))

    def _analyze_heading_styles(self, paragraphs: List[dict], result: FormatAnalysisResult):
        """分析标题样式"""
        # 检测可能的标题（字体较大或加粗）
        potential_headings = []
        for i, para in enumerate(paragraphs):
            text = para.get('text', '')
            if not text or len(text) > 100:
                continue
            font_size = para.get('font_size', 12)
            is_bold = para.get('bold', False)

            # 大字体或加粗可能是标题
            if font_size >= 16 or is_bold:
                potential_headings.append({
                    'index': i,
                    'text': text[:50],
                    'size': font_size,
                    'bold': is_bold
                })

        if potential_headings:
            result.issues.append(FormatIssue(
                issue_type=FormatIssueType.STYLE,
                severity="info",
                message=f"发现{len(potential_headings)}个可能的标题",
                location="全文",
                suggested_fix="建议应用标准标题样式以便于目录生成",
                related_paragraph=None
            ))

    def _generate_summary(self, result: FormatAnalysisResult) -> dict:
        """生成摘要"""
        return {
            'total_issues': result.total_issues,
            'errors': result.errors,
            'warnings': result.warnings,
            'infos': result.infos,
            'issue_types': list(set(i.issue_type.value for i in result.issues)),
            'health_score': max(0, 100 - result.errors * 10 - result.warnings * 5)
        }

    def get_fix_suggestions(self, issue: FormatIssue) -> List[dict]:
        """获取修复建议"""
        suggestions_map = {
            FormatIssueType.FONT_NAME: [
                {'action': 'set_font', 'params': {'font_name': '微软雅黑'}, 'description': '统一使用微软雅黑'}
            ],
            FormatIssueType.FONT_SIZE: [
                {'action': 'set_font_size', 'params': {'size': 12}, 'description': '设置正文字体为12磅'}
            ],
            FormatIssueType.LINE_SPACING: [
                {'action': 'set_line_spacing', 'params': {'spacing': 1.5}, 'description': '设置1.5倍行距'}
            ],
            FormatIssueType.PARAGRAPH_ALIGNMENT: [
                {'action': 'set_alignment', 'params': {'alignment': 'left'}, 'description': '设置左对齐'}
            ]
        }
        return suggestions_map.get(issue.issue_type, [])
