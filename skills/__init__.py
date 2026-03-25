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

        # 解析 frontmatter
        self.description = ""
        self._load_metadata()

    def _load_metadata(self):
        """从 SKILL.md 中提取 frontmatter 的 description"""
        if not self.skill_md.exists():
            return

        content = self.skill_md.read_text(encoding="utf-8")

        # 提取 YAML frontmatter
        match = re.match(r'^---\n(.*?)\n---', content, re.DOTALL)
        if match:
            frontmatter = match.group(1)
            for line in frontmatter.split('\n'):
                if line.startswith('description:'):
                    self.description = line.replace('description:', '').strip()
                    break

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
        """列出所有 Skill 的名称和描述（用于 AI 决策）"""
        return [
            {"name": s.name, "description": s.description}
            for s in self._skills
        ]

    def get_skill(self, name: str) -> Optional[Skill]:
        """根据名称获取 Skill"""
        for s in self._skills:
            if s.name == name:
                return s
        return None

    def get_all_descriptions(self) -> str:
        """生成所有 Skill 的描述文本（供 AI 阅读）"""
        lines = ["## 可用的 Word 操作技能\n"]
        lines.append("AI Agent 应根据用户需求，从以下技能中选择合适的操作。\n")

        for skill in self._skills:
            content = skill.get_content()
            if content:
                lines.append(f"\n### {skill.name}\n")
                lines.append(content)

        return "\n".join(lines)


# 全局 Skill 加载器
_loader: Optional[SkillLoader] = None


def get_skill_loader() -> SkillLoader:
    """获取全局 Skill 加载器"""
    global _loader
    if _loader is None:
        _loader = SkillLoader()
    return _loader


def list_available_skills() -> List[Dict[str, str]]:
    """列出所有可用技能"""
    return get_skill_loader().list_skills()


def get_skill_descriptions() -> str:
    """获取所有技能的详细描述"""
    return get_skill_loader().get_all_descriptions()
