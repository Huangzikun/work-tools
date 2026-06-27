"""教案生成核心模块（从 teaching_plan_v2/ 抽取并参数化）。"""

from .models import LessonPlan
from .parser import DEFAULT_SYSTEM_PROMPT, SyllabusParser
from .docx_builder import DocxBuilder

__all__ = ["LessonPlan", "SyllabusParser", "DocxBuilder", "DEFAULT_SYSTEM_PROMPT"]
