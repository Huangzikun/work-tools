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


DEFAULT_USER_PROMPT = (
    Path(__file__).parent / "default_user_prompt.txt"
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
    def __init__(
        self,
        llm_client,
        system_prompt: Optional[str] = None,
        user_prompt: Optional[str] = None,
    ):
        self.llm = llm_client
        self.system_prompt = system_prompt or DEFAULT_SYSTEM_PROMPT
        # user_prompt 是用户的「补充要求」：非空时作为额外区块追加到系统自动拼接的
        # user prompt 末尾（课次范围/字段/大纲/数量约束仍由系统保证）。
        self.user_prompt = (user_prompt or "").strip() or DEFAULT_USER_PROMPT
        # 时间预算默认值（parse() 会按 total_hours/total_lessons 重新计算）。
        # 默认每教案 5 课时 × 40 分钟 = 200 分钟，兼容旧逻辑。
        self.total_lessons = 0
        self.total_hours: Optional[int] = None
        self.per_lesson_hours = 5
        self.per_lesson_minutes = 200

    def parse(
        self,
        syllabus_path: str,
        total_lessons: int,
        batch_size: int = 2,
        progress_callback: Optional[ProgressCallback] = None,
        total_hours: Optional[int] = None,
    ) -> List[LessonPlan]:
        syllabus_content = self._read_syllabus(syllabus_path)
        logger.info("syllabus length=%d chars", len(syllabus_content))

        # 按课时数动态计算每教案可用时间：
        # 每教案课时 = round(总课时 / 教案数)，每课时 40 分钟 → 每教案分钟数。
        # 不整除时四舍五入到整数课时（如 50/17≈2.94 → 3 课时）。
        self.total_lessons = total_lessons
        self.total_hours = total_hours
        if total_hours and total_hours > 0 and total_lessons > 0:
            self.per_lesson_hours = int(total_hours / total_lessons + 0.5)
        else:
            self.per_lesson_hours = 5
        self.per_lesson_minutes = self.per_lesson_hours * 40
        logger.info(
            "time budget: total_hours=%s, lessons=%d -> %d 课时/教案 = %d 分钟/教案",
            total_hours, total_lessons, self.per_lesson_hours, self.per_lesson_minutes,
        )

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

            try:
                lessons = self._generate_batch_with_fallback(
                    syllabus_content, expected_numbers, total_lessons
                )

                # 强制按期望课次赋值：LLM 输出的"课次"字段不可信（可能为 0/null/跳号），
                # 由调用方按批次顺序强制覆盖为 start, start+1, ..., end
                for i, lesson in enumerate(lessons):
                    lesson.课次 = expected_numbers[i]

                all_lessons.extend(lessons)
                logger.info("generated %d lessons (expected %d-%d)",
                            len(lessons), start, end)

                if progress_callback:
                    progress_callback(end, total_lessons, f"已完成 {end}/{total_lessons}")
            except Exception:
                logger.exception("生成第 %d-%d 个教案失败", start, end)
                raise

        return all_lessons

    # 核心字段：任一为空则教案无效（触发单课重试），根治空白教案
    _REQUIRED_FIELDS = (
        "授课内容", "课堂内容", "课堂导入", "课堂小结",
        "课后作业", "教学方法与设计",
    )
    _MIN_CONTENT_LEN = 200  # 课堂内容最小字数，低于此视为内容过短

    def _is_lesson_valid(self, lesson: LessonPlan) -> bool:
        """核心字段非空且课堂内容达到最小长度，才算有效。"""
        for field in self._REQUIRED_FIELDS:
            if not (getattr(lesson, field, "") or "").strip():
                return False
        if len((lesson.课堂内容 or "").strip()) < self._MIN_CONTENT_LEN:
            return False
        return True

    def _generate_batch_with_fallback(
        self,
        syllabus_content: str,
        expected_numbers: List[int],
        total_lessons: int,
    ) -> List[LessonPlan]:
        """整批生成 + 不足时单课兜底。绝不返回占位符。

        流程：
        1. 整批调用最多 MAX_BATCH_ATTEMPTS 次，每次给 LLM 反馈上次返回数量
        2. 仍不够 → 对缺失的每个课次单独调用 LLM（最多 MAX_SINGLE_ATTEMPTS 次）
        3. 单课仍失败 → 抛错让任务进 failed，而不是用空白教案骗用户
        """
        batch_size = len(expected_numbers)
        max_batch_attempts = 3
        max_single_attempts = 2

        lessons: List[LessonPlan] = []
        batch_ok = False
        for attempt in range(max_batch_attempts):
            lessons = self._call_batch_llm(
                syllabus_content,
                expected_numbers,
                total_lessons,
                attempt=attempt,
                prev_count=len(lessons) if attempt > 0 else None,
            )
            if len(lessons) >= batch_size:
                logger.info(
                    "整批成功 (第 %d 次): 返回 %d 个, 期望 %d 个",
                    attempt + 1, len(lessons), batch_size,
                )
                lessons = lessons[:batch_size]
                batch_ok = True
                break
            logger.warning(
                "整批第 %d 次返回 %d 个, 期望 %d 个",
                attempt + 1, len(lessons), batch_size,
            )

        # 整批仍不足 → 单课兜底补齐缺失位置
        if not batch_ok:
            missing = batch_size - len(lessons)
            logger.warning(
                "整批 %d 次仍缺 %d 个，进入单课兜底",
                max_batch_attempts, missing,
            )
            for idx in range(len(lessons), batch_size):
                expected_num = expected_numbers[idx]
                lessons.append(
                    self._call_single_lesson(
                        syllabus_content, expected_num, total_lessons, max_single_attempts
                    )
                )
                logger.info("单课兜底成功: 课次 %d", expected_num)

        # 核心字段校验：内容为空/过短的课次走单课重试，根治空白教案
        for idx in range(batch_size):
            if not self._is_lesson_valid(lessons[idx]):
                expected_num = expected_numbers[idx]
                logger.warning(
                    "课次 %d 核心字段为空或课堂内容过短（%d 字），单课重试",
                    expected_num, len((lessons[idx].课堂内容 or "").strip()),
                )
                lessons[idx] = self._call_single_lesson(
                    syllabus_content, expected_num, total_lessons, max_single_attempts
                )

        return lessons

    def _call_batch_llm(
        self,
        syllabus_content: str,
        expected_numbers: List[int],
        total_lessons: int,
        attempt: int,
        prev_count: Optional[int],
    ) -> List[LessonPlan]:
        start = expected_numbers[0]
        end = expected_numbers[-1]
        user_prompt = self._build_user_prompt(
            syllabus_content, start, end, total_lessons
        )
        if attempt > 0 and prev_count is not None:
            user_prompt += (
                f"\n\n[重要] 上次调用仅返回 {prev_count} 个对象，"
                f"期望 {len(expected_numbers)} 个，本次必须严格按数量输出。"
            )

        response_text = self.llm.generate_with_retry(
            system_prompt=self.system_prompt,
            user_prompt=user_prompt,
            json_output=True,
        )
        lessons_data = self._parse_json_response(response_text)
        return [LessonPlan.from_dict(d) for d in lessons_data]

    def _call_single_lesson(
        self,
        syllabus_content: str,
        lesson_num: int,
        total_lessons: int,
        max_attempts: int,
    ) -> LessonPlan:
        user_prompt = (
            f"用户要求生成 {total_lessons} 个教案，本次只需要生成第 {lesson_num} 个，"
            f"输出且仅输出 1 个对象，其「课次」字段必须为 {lesson_num}。\n"
            f"必须紧扣教学大纲，按章节顺序合理选择本课次的内容，不可编造。\n\n"
            f"字段列表：\n{LESSON_FIELDS}\n\n"
            f"教学大纲内容：\n{syllabus_content}\n"
        ) + self._build_time_constraint() + self._format_user_supplement()
        last_exc: Optional[Exception] = None
        for attempt in range(max_attempts):
            try:
                response_text = self.llm.generate_with_retry(
                    system_prompt=self.system_prompt,
                    user_prompt=user_prompt,
                    json_output=True,
                )
                lessons_data = self._parse_json_response(response_text)
                if lessons_data:
                    return LessonPlan.from_dict(lessons_data[0])
                logger.warning(
                    "单课调用返回空 (课次 %d, 第 %d 次)", lesson_num, attempt + 1
                )
            except Exception as e:
                last_exc = e
                logger.warning(
                    "单课调用异常 (课次 %d, 第 %d 次): %s",
                    lesson_num, attempt + 1, e,
                )

        raise RuntimeError(
            f"课次 {lesson_num} 经 {max_attempts} 次单课生成仍失败"
        ) from last_exc

    def _read_syllabus(self, file_path: str) -> str:
        doc = Document(file_path)
        parts = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    parts.append(cell.text)
        return "\n".join(parts)

    def _build_time_constraint(self) -> str:
        """动态时间约束：算每教案可用分钟，并扣掉导入/小结得到课堂内容主体时间。"""
        total_hours_desc = f"{self.total_hours}" if self.total_hours else "未指定（按默认）"
        import_minutes = 10
        summary_minutes = 10
        content_minutes = max(
            self.per_lesson_minutes - import_minutes - summary_minutes, 0
        )
        return (
            "\n本次时间分配要求：\n"
            f"- 本课程共 {total_hours_desc} 课时，分 {self.total_lessons} 次教案，"
            f"每次教案约 {self.per_lesson_hours} 课时，合计可用约 {self.per_lesson_minutes} 分钟。\n"
            f"- 时间拆分：课堂导入 {import_minutes} 分钟 + 课堂小结 {summary_minutes} 分钟，"
            f"「课堂内容」主体可用约 {content_minutes} 分钟"
            f"（= {self.per_lesson_minutes} − {import_minutes} − {summary_minutes}）。\n"
            f"- 「课堂内容」字段内各教学环节标注的分钟数之和应等于（或接近）{content_minutes} 分钟；"
            f"对应字段填：「课堂导入时间分配」={import_minutes} 分钟、"
            f"「课堂小结时间分配」={summary_minutes} 分钟、"
            f"「课堂内容时间分配」={content_minutes} 分钟。\n"
            "- 课堂导入与课堂小结是独立字段（各不超过 10 分钟），「课堂内容」是扣除二者后的主体环节，"
            "不要在「课堂内容」里再重复写\"导入\"\"小结\"环节。\n"
        )

    def _format_user_supplement(self) -> str:
        """用户补充要求区块：非空时作为额外约束追加到 user prompt 末尾。"""
        supp = (self.user_prompt or "").strip()
        if not supp:
            return ""
        return (
            "\n\n【本次生成的额外要求（请严格遵循）】\n"
            f"{supp}"
        )

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
            f"{self._build_time_constraint()}"
            f"{self._format_user_supplement()}"
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
