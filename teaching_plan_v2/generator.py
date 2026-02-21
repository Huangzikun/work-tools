"""
教案生成器 - 主流程编排
"""
import logging
from pathlib import Path
from typing import Optional

from .ai_client import AIClient
from .parser import SyllabusParser
from .docx_builder import DocxBuilder
from .config import AppConfig

logger = logging.getLogger(__name__)


class LessonPlanGenerator:
    """教案生成器 - 主流程编排"""

    def __init__(self, config: AppConfig):
        """
        初始化生成器

        Args:
            config: 应用配置
        """
        self.config = config

        # 初始化AI客户端
        self.ai_client = AIClient(
            api_key=config.api_key,
            model=config.model,
            timeout=config.timeout,
            max_retries=config.max_retries
        )

        # 初始化大纲解析器
        self.parser = SyllabusParser(self.ai_client)

        # 初始化文档构建器
        self.builder = DocxBuilder(
            config.template_path,
            config.output_dir
        )

        logger.info("教案生成器初始化完成")

    def generate(
        self,
        syllabus_path: str,
        total_lessons: int,
        output_filename: str = None
    ) -> str:
        """
        生成合并的教案文档

        Args:
            syllabus_path: 大纲文件路径
            total_lessons: 总课时数
            output_filename: 输出文件名（可选，默认为"大纲名-教案.docx"）

        Returns:
            输出文件路径
        """
        logger.info(f"开始生成教案，总计: {total_lessons}个课时")
        logger.info(f"大纲文件: {syllabus_path}")

        # 1. 解析大纲
        lessons = self.parser.parse(
            syllabus_path,
            total_lessons,
            self.config.batch_size
        )
        logger.info(f"AI解析完成，生成{len(lessons)}个课时")

        # 2. 生成输出文件名
        if output_filename is None:
            syllabus_name = Path(syllabus_path).stem
            output_filename = f"{syllabus_name}-教案.docx"

        # 3. 生成合并的教案文档
        output_path = self.builder.build_all(lessons, output_filename)
        logger.info(f"教案生成完成: {output_path}")

        return output_path
