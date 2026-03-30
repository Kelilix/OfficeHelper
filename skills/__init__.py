# Skills 模块
# 标准格式 Skill 集合，供 AI Agent 调用 Word COM 操作

import os
from pathlib import Path
from typing import List, Dict, Optional
import re


class Skill:
    """单个 Skill 定义"""

    def __init__(self, skill_dir: Path):
        self.name = skill_dir.name
        self.path = skill_dir
        self.skill_md = skill_dir / "SKILL.md"
        self.description = ""     # frontmatter 描述
        self.actions = []         # 正文中的操作列表，如 ["set_font", "set_font_size", ...]
        self._load_metadata()

    def _load_metadata(self):
        """从 SKILL.md 中提取 frontmatter 描述 + 正文操作列表"""
        if not self.skill_md.exists():
            return

        content = self.skill_md.read_text(encoding="utf-8")

        # 提取 YAML frontmatter 的 description
        match = re.match(r'^---\n(.*?)\n---', content, re.DOTALL)
        if match:
            frontmatter = match.group(1)
            for line in frontmatter.split('\n'):
                if line.startswith('description:'):
                    self.description = line.replace('description:', '').strip()
                    break

        # 提取正文中的操作列表（### N. 操作名 (action_name)）
        action_matches = re.findall(r'### \d+\.\s+.+\s+\(([a-zA-Z_][a-zA-Z0-9_]*)\)', content)
        self.actions = action_matches

    def get_content(self) -> str:
        """获取 Skill 的完整内容"""
        if self.skill_md.exists():
            return self.skill_md.read_text(encoding="utf-8")
        return ""


class SkillLoader:
    """Skill 加载器"""

    def __init__(self, skills_dir: Optional[Path] = None):
        if skills_dir is None:
            # 默认使用项目下的 skills 目录
            self.skills_dir = Path(__file__).parent
        else:
            self.skills_dir = Path(skills_dir)

        self._skills: List[Skill] = []
        self._load_skills()

    def _load_skills(self):
        """加载所有 Skill"""
        self._skills = []
        if not self.skills_dir.exists():
            return

        for item in self.skills_dir.iterdir():
            if item.is_dir() and (item / "SKILL.md").exists():
                self._skills.append(Skill(item))

    def list_skills(self) -> List[Dict[str, str]]:
        """列出所有 Skill 的名称、描述和操作列表"""
        return [
            {"name": s.name, "description": s.description, "actions": s.actions}
            for s in self._skills
        ]

    def get_skill(self, name: str) -> Optional[Skill]:
        """根据名称获取 Skill"""
        for s in self._skills:
            if s.name == name:
                return s
        return None

    def get_all_descriptions(self) -> str:
        """
        生成所有 Skill 的描述（name + frontmatter 描述 + 操作列表）。

        用于第一轮：让 LLM 感知有哪些技能可用，以便选择最合适的一个。
        """
        lines = ["## 可用的 Word 操作技能\n"]
        lines.append("请根据用户需求，从以下技能中选择最合适的（只选一个）。返回格式：\n"
                     '```json\n{"skill": "技能目录名", "reasoning": "简短选择理由"}\n```\n\n')
        for s in self._skills:
            lines.append(f"- **{s.name}**：{s.description}")
            if s.actions:
                lines.append(f"  支持操作：{', '.join(s.actions)}")
            lines.append("")
        return "\n".join(lines)

    def get_all_full(self) -> str:
        """
        生成所有 Skill 的完整内容（用于第二轮 Skill 被选中后）。

        注意：实际不会一次性传全部；本方法保留用于按 name 获取单个。
        """
        parts = []
        for s in self._skills:
            content = s.get_content()
            if content:
                parts.append(f"\n### {s.name}\n{content}\n")
        return "\n".join(parts)


def get_skill_loader() -> SkillLoader:
    """获取 Skill 加载器（每次返回新实例，确保实时反映 skills 目录变化）。"""
    return SkillLoader()


def list_available_skills() -> List[Dict[str, str]]:
    """列出所有可用技能"""
    return get_skill_loader().list_skills()


def get_skill_descriptions() -> str:
    """获取所有技能的简短描述（name + description），用于第一轮 LLM 选技能。"""
    return get_skill_loader().get_all_descriptions()


def get_skill_content(name: str) -> str:
    """根据 skill 目录名获取其完整 SKILL.md 内容，用于第二轮执行。"""
    skill = get_skill_loader().get_skill(name)
    if skill:
        return skill.get_content()
    return ""


def list_skill_names() -> List[str]:
    """返回所有 skill 的目录名列表。"""
    return [s.name for s in get_skill_loader()._skills]
