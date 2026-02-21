"""
教案生成工具V2

使用AI从教学大纲自动生成教案
"""

__version__ = "2.0.0"

from .models import LessonPlan
from .config import AppConfig
from .generator import LessonPlanGenerator

__all__ = [
    "LessonPlan",
    "AppConfig",
    "LessonPlanGenerator",
]
