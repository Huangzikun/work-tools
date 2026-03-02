"""
教学大纲解析器 - 使用AI生成教案列表
"""
import json
import logging
from typing import List
from docx import Document

from .models import LessonPlan
from .ai_client import AIClient

logger = logging.getLogger(__name__)


# 默认系统提示词
DEFAULT_SYSTEM_PROMPT = '''
你是桂林学院信息工程学院的一名计算机专任教师，你拥有丰富的教学经验。
你的任务是按照读取教学大纲，根据教学大纲中的章节和实验安排，合理分布到用户提供的课时安排中。

遵守以下限制：
1. 每个课时为40分钟，如果每次课时为5个课时，则总共可使用的时间分配为5*40=200分钟。
2. 课堂导入、课后总结时间不得超过10分钟。
3. 课堂内容只描述内容，不要用括号举例的形式，不要使用代码。课堂内容应该包括知识内容、教学方法、教学工具或资源等。
4. **课堂内容字段必须详细、丰富，字数控制在300-500字之间**。要求：
   - 按时间顺序详细描述教学步骤，每个步骤包含：时间分配、教学活动、授课方法
   - 明确标注授课方法：如"讲授法""案例分析法""小组讨论法""情景模拟法""任务驱动法""项目教学法""翻转课堂"等
   - 详细说明教师活动：讲解什么内容、使用什么案例、如何引导、如何示范
   - 详细说明学生活动：讨论什么问题、完成什么任务、进行什么实践、如何展示
   - 每个知识点都要说明如何展开，使用什么教学资源或工具
   - **必须明确融入思政元素**：在适当环节标注思政融入点，如：
     * "融入工匠精神教育，强调精益求精"
     * "融入家国情怀，增强民族自豪感"
     * "融入法治意识，强调合规经营"
     * "融入职业道德，培养诚信品质"
     * "融入社会责任，树立正确价值观"
   - 语言要具体、准确，避免空洞和笼统的表述

请严格生成JSON格式的内容，每个字段都必须包含在JSON对象中。
输出格式应该是JSON数组，每个元素是一个教案对象。
'''


# 教案字段列表
LESSON_FIELDS = '''
{
    "课次": "应为一个整数，如1",
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
    "课堂内容": "必填！要求300-500字，按时间顺序详细描述教学步骤，每步包含时间分配、教学活动、授课方法；说明教师活动（讲解内容、案例、引导方式）和学生活动（讨论、任务、实践、展示）；明确融入思政元素（如工匠精神、家国情怀、法治意识、职业道德等），语言具体准确",
    "课堂内容时间分配": "如180分钟",
    "教学方法与设计": "教学方法和设计思路",
    "课堂小结": "课堂小结内容",
    "课堂小结时间分配": "如10分钟",
    "课后作业": "课后作业内容",
    "课后作业时间分配": "如60分钟",
    "教学反思": "教学反思（可选，可为空字符串）"
}
'''


class SyllabusParser:
    """教学大纲解析器"""

    def __init__(self, ai_client: AIClient, system_prompt: str = None):
        """
        初始化解析器

        Args:
            ai_client: AI客户端
            system_prompt: 系统提示词，如果为None则使用默认值
        """
        self.ai_client = ai_client
        self.system_prompt = system_prompt or DEFAULT_SYSTEM_PROMPT

    def parse(
        self,
        syllabus_path: str,
        total_lessons: int,
        batch_size: int = 2
    ) -> List[LessonPlan]:
        """
        解析大纲并生成教案列表

        Args:
            syllabus_path: 大纲文件路径
            total_lessons: 总教案数量
            batch_size: 每批生成的教案数量

        Returns:
            教案列表
        """
        logger.info(f"开始解析大纲: {syllabus_path}")

        # 1. 读取Word内容
        syllabus_content = self._read_syllabus(syllabus_path)
        logger.info(f"大纲内容读取完成，长度: {len(syllabus_content)} 字符")

        # 2. 分批调用AI生成教案
        all_lessons = []
        for start in range(1, total_lessons + 1, batch_size):
            end = min(start + batch_size - 1, total_lessons)
            logger.info(f"生成第{start}-{end}个教案...")

            # 构建用户提示词
            user_prompt = self._build_user_prompt(
                syllabus_content,
                start,
                end,
                total_lessons
            )

            # 调用AI
            try:
                response_text = self.ai_client.generate_response_with_retry(
                    self.system_prompt,
                    user_prompt,
                    response_format="json"
                )

                # 解析JSON响应
                lessons_data = self._parse_json_response(response_text)
                lessons = [LessonPlan.from_dict(data) for data in lessons_data]
                all_lessons.extend(lessons)

                logger.info(f"成功生成{len(lessons)}个教案")

            except Exception as e:
                logger.error(f"生成第{start}-{end}个教案失败: {e}")
                raise

        logger.info(f"大纲解析完成，共生成{len(all_lessons)}个教案")
        return all_lessons

    def _read_syllabus(self, file_path: str) -> str:
        """
        读取大纲文档内容

        Args:
            file_path: 文件路径

        Returns:
            文档文本内容
        """
        doc = Document(file_path)
        full_text = []

        # 读取段落
        for para in doc.paragraphs:
            full_text.append(para.text)

        # 读取表格
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    full_text.append(cell.text)

        return '\n'.join(full_text)

    def _build_user_prompt(
        self,
        syllabus_content: str,
        start_lesson: int,
        end_lesson: int,
        total_lessons: int
    ) -> str:
        """
        构建用户提示词

        Args:
            syllabus_content: 大纲内容
            start_lesson: 开始课次
            end_lesson: 结束课次
            total_lessons: 总教案数

        Returns:
            用户提示词
        """
        return f'''根据文档内容中的教学时数/5确定教案的数量，如教学时数为5，即你应该给我提供10/5=2个教案的内容。

总共应提供{total_lessons}个教案。
当前输出第{start_lesson}~{end_lesson}个教案。

字段列表如下：
{LESSON_FIELDS}

文档内容如下：
{syllabus_content}

如果文档未提及，则对象内的这一项设置为空字符串。
如果一项中包含一个列表，则每一个条目后面增加换行。'''

    def _parse_json_response(self, response_text: str) -> List[dict]:
        """
        解析JSON响应

        Args:
            response_text: AI返回的JSON文本

        Returns:
            教案数据列表

        Raises:
            ValueError: JSON解析失败
        """
        try:
            # 尝试直接解析
            data = json.loads(response_text)

            # 如果是字典，提取plans字段
            if isinstance(data, dict):
                if 'plans' in data:
                    return data['plans']
                elif 'data' in data:
                    return data['data']
                else:
                    # 如果没有plans或data字段，返回字典本身
                    return [data]

            # 如果是列表，直接返回
            if isinstance(data, list):
                return data

            raise ValueError(f"未知的JSON格式: {type(data)}")

        except json.JSONDecodeError as e:
            logger.error(f"JSON解析失败: {e}")
            logger.error(f"响应内容: {response_text[:500]}...")
            raise ValueError(f"AI返回的内容不是有效的JSON格式: {e}") from e
