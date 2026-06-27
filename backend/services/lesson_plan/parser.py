"""教学大纲解析器（参数化版，复制自 teaching_plan_v2/parser.py 并改造）。

改造点：
- 移除 ai_client 依赖，直接用 common.llm_client.LLMClient
- 增加 progress_callback(done, total, label) 钩子，方便服务层上报进度
- system_prompt 由外部传入（高级用户可改）
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Callable, List, Optional

from docx import Document

from .models import LessonPlan

logger = logging.getLogger(__name__)


DEFAULT_SYSTEM_PROMPT = (
    Path(__file__).parent / "default_system_prompt.txt"
).read_text(encoding="utf-8").strip()


LESSON_FIELDS = """
{
    "课次": "由调用方指定，本批次必须按顺序填入期望值，不可为 0 或空",
    "授课内容": "课程章节标题",
    "授课学时": "应为一个整数，如3或5",
    "知识目标": "学生应掌握的知识点",
    "能力目标": "学生应具备的能力",
    "情感与价值观目标": "学生应培养的情感和价值观",
    "重点": "本课的重点内容",
    "难点": "本课的难点内容",
    "教学工具或资源": "使用的教学工具或资源",
    "课堂导入": "课堂导入内容和方式",
    "课堂导入时间分配": "如10分钟",
    "课堂内容": "必填！要求300-500字，按时间顺序详细描述教学步骤，每步包含时间分配、教学内容、学生活动；说明教师活动（讲解内容、案例、引导方式）和学生活动（讨论、任务、实践、展示）；整堂课选用1-2种教学方法（讲授法、案例分析法、小组讨论法、情景模拟法、任务驱动法、项目教学法、翻转课堂等）；明确融入思政元素（如工匠精神、家国情怀、法治意识、职业道德等）；课堂内容分节要清晰，每个主要教学环节（如导入、新授、练习、总结）之间使用换行分隔，保持排版美观，便于阅读",
    "课堂内容时间分配": "如180分钟",
    "教学方法与设计": "教学方法和设计思路",
    "课堂小结": "课堂小结内容",
    "课堂小结时间分配": "如10分钟",
    "课后作业": "课后作业内容",
    "课后作业时间分配": "如60分钟",
    "教学反思": "教学反思（可选，可为空字符串）"
}
"""


ProgressCallback = Callable[[int, int, str], None]


class SyllabusParser:
    def __init__(self, llm_client, system_prompt: Optional[str] = None):
        self.llm = llm_client
        self.system_prompt = system_prompt or DEFAULT_SYSTEM_PROMPT

    def parse(
        self,
        syllabus_path: str,
        total_lessons: int,
        batch_size: int = 2,
        progress_callback: Optional[ProgressCallback] = None,
    ) -> List[LessonPlan]:
        syllabus_content = self._read_syllabus(syllabus_path)
        logger.info("syllabus length=%d chars", len(syllabus_content))

        all_lessons: List[LessonPlan] = []
        if progress_callback:
            progress_callback(0, total_lessons, "开始解析大纲")

        for start in range(1, total_lessons + 1, batch_size):
            end = min(start + batch_size - 1, total_lessons)
            expected_numbers = list(range(start, end + 1))
            label = f"正在生成第 {start}-{end} 个教案"
            logger.info(label)
            if progress_callback:
                progress_callback(start - 1, total_lessons, label)

            user_prompt = self._build_user_prompt(
                syllabus_content, start, end, total_lessons
            )

            try:
                response_text = self.llm.generate_with_retry(
                    system_prompt=self.system_prompt,
                    user_prompt=user_prompt,
                    json_output=True,
                )
                lessons_data = self._parse_json_response(response_text)
                lessons = [LessonPlan.from_dict(d) for d in lessons_data]

                # 强制按期望课次赋值：LLM 输出的"课次"字段不可信（可能为 0/null/跳号），
                # 由调用方按批次顺序强制覆盖为 start, start+1, ..., end
                for i, lesson in enumerate(lessons):
                    if i < len(expected_numbers):
                        lesson.课次 = expected_numbers[i]

                # LLM 输出数量不足时补空占位（避免表格缺失）
                while len(lessons) < len(expected_numbers):
                    idx = len(lessons)
                    placeholder = LessonPlan(
                        课次=expected_numbers[idx],
                        授课内容=f"第 {expected_numbers[idx]} 次课（AI 未生成，待补充）",
                    )
                    lessons.append(placeholder)
                    logger.warning(
                        "LLM 返回教案数量不足，已补占位：期望第 %d 课",
                        expected_numbers[idx],
                    )

                # LLM 输出数量超出时截断
                if len(lessons) > len(expected_numbers):
                    logger.warning(
                        "LLM 返回教案数量超出，已截断：实际 %d 期望 %d",
                        len(lessons),
                        len(expected_numbers),
                    )
                    lessons = lessons[: len(expected_numbers)]

                all_lessons.extend(lessons)
                logger.info("generated %d lessons (expected %d-%d)",
                            len(lessons), start, end)

                if progress_callback:
                    progress_callback(end, total_lessons, f"已完成 {end}/{total_lessons}")
            except Exception:
                logger.exception("生成第 %d-%d 个教案失败", start, end)
                raise

        return all_lessons

    def _read_syllabus(self, file_path: str) -> str:
        doc = Document(file_path)
        parts = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    parts.append(cell.text)
        return "\n".join(parts)

    def _build_user_prompt(
        self,
        syllabus_content: str,
        start_lesson: int,
        end_lesson: int,
        total_lessons: int,
    ) -> str:
        expected = "、".join(str(i) for i in range(start_lesson, end_lesson + 1))
        return (
            f"用户要求生成 {total_lessons} 个教案，对应教学大纲中不同课次。\n"
            f"当前批次必须输出第 {start_lesson}~{end_lesson} 个教案，共 {end_lesson - start_lesson + 1} 个对象。\n"
            f"每个对象的「课次」字段必须严格按顺序填入：{expected}\n\n"
            f"字段列表如下：\n{LESSON_FIELDS}\n\n"
            f"教学大纲内容如下：\n{syllabus_content}\n\n"
            f"约束：\n"
            f"- 必须输出且仅输出 {end_lesson - start_lesson + 1} 个教案对象，不能多也不能少。\n"
            f"- 「课次」字段必须是上述期望值之一，按顺序对应，不可为 0 或空。\n"
            f"- 内容必须紧扣教学大纲，按章节顺序合理分配到每个课次，不可编造。\n"
            f"- 如果大纲未提及，则对象内对应项设置为空字符串。\n"
            f"- 如果一项中包含列表，每一条目后用换行分隔。\n"
        )

    def _parse_json_response(self, response_text: str) -> List[dict]:
        try:
            data = json.loads(response_text)
            if isinstance(data, dict):
                if "plans" in data:
                    return data["plans"]
                if "data" in data:
                    return data["data"]
                return [data]
            if isinstance(data, list):
                return data
            raise ValueError(f"未知 JSON 格式: {type(data)}")
        except json.JSONDecodeError as e:
            logger.error("JSON 解析失败: %s", e)
            logger.error("响应前 500 字符: %s", response_text[:500])
            raise ValueError(f"AI 返回不是有效 JSON: {e}") from e
