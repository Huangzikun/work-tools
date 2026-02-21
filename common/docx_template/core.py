"""
核心替换器 - DocxTemplateReplacer
"""
from docx import Document
from pathlib import Path
from typing import Dict, Any, Optional
import logging

from .config import ReplacerConfig
from .format_handler import FormatHandler
from .exceptions import TemplateLoadError, InvalidDataError
from .utils import validate_data

logger = logging.getLogger(__name__)


class DocxTemplateReplacer:
    """
    DOCX模板替换器

    支持全量替换和首次替换两种模式
    """

    def __init__(self, template_path: str, config: Optional[ReplacerConfig] = None):
        """
        初始化替换器

        Args:
            template_path: 模板文件路径
            config: 配置对象,默认使用DefaultConfig
        """
        self.config = config or ReplacerConfig()
        self.format_handler = FormatHandler(self.config.format)
        self.template_path = Path(template_path)

        # 统计信息
        self.stats = {
            'placeholders_found': 0,
            'placeholders_replaced': 0,
            'errors': []
        }

        # 首次替换模式：记录已替换的占位符
        self._replaced_placeholders = set()

        # 加载模板
        try:
            self.doc = Document(str(self.template_path))
            logger.info(f"成功加载模板: {template_path}")
        except Exception as e:
            raise TemplateLoadError(f"加载模板失败: {e}") from e

    def replace(self, data: Dict[str, Any], output_path: Optional[str] = None) -> Document:
        """
        执行批量替换

        Args:
            data: 数据字典 {key: value}
            output_path: 输出文件路径,如果为None则不保存

        Returns:
            Document对象
        """
        if not isinstance(data, dict):
            raise InvalidDataError("data参数必须是字典类型")

        # 重置已替换占位符记录
        self._replaced_placeholders = set()

        if not data:
            logger.warning("数据字典为空,无需替换")
            # 即使数据为空，如果有output_path也要保存文档
            if output_path:
                self.save(output_path)
            return self.doc

        # 验证数据
        if self.config.strict_mode:
            validate_data(data)

        logger.info(f"开始替换,数据键数量: {len(data)}, replace_all={self.config.replace_all}")

        # 构建占位符到值的映射
        placeholder_map = self._build_placeholder_map(data)

        # 替换表格中的占位符
        self._replace_in_tables(placeholder_map)

        # 替换段落中的占位符
        self._replace_in_paragraphs(placeholder_map)

        # 保存文档
        if output_path:
            self.save(output_path)

        logger.info(f"替换完成,统计: {self.stats}")
        return self.doc

    def _build_placeholder_map(self, data: Dict[str, Any]) -> Dict[str, str]:
        """
        构建占位符到值的映射

        Args:
            data: 数据字典

        Returns:
            占位符映射字典
        """
        placeholder_map = {}
        for key, value in data.items():
            placeholder = f"{self.config.placeholder.prefix}{key}{self.config.placeholder.suffix}"
            placeholder_map[placeholder] = str(value)
        return placeholder_map

    def _replace_in_paragraphs(self, placeholder_map: Dict[str, str]):
        """
        替换段落中的占位符

        Args:
            placeholder_map: 占位符映射字典
        """
        for paragraph in self.doc.paragraphs:
            if self._replace_paragraph(paragraph, placeholder_map):
                self.stats['placeholders_replaced'] += 1

    def _replace_in_tables(self, placeholder_map: Dict[str, str]):
        """
        替换表格中的占位符

        Args:
            placeholder_map: 占位符映射字典
        """
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if self._replace_paragraph(paragraph, placeholder_map):
                            self.stats['placeholders_replaced'] += 1

    def _replace_paragraph(self, paragraph, placeholder_map: Dict[str, str]) -> bool:
        """
        替换单个段落中的占位符

        Args:
            paragraph: docx的Paragraph对象
            placeholder_map: 占位符映射字典

        Returns:
            bool: 是否进行了替换
        """
        # 检查段落是否包含占位符
        paragraph_text = paragraph.text
        has_placeholder = any(ph in paragraph_text for ph in placeholder_map.keys())

        if not has_placeholder:
            return False

        # 提取所有run的文本和格式
        runs_data = []
        full_text = ""
        for run in paragraph.runs:
            run_text = run.text
            run_format = self.format_handler.extract_run_format(run)
            runs_data.append({
                'text': run_text,
                'format': run_format,
                'run': run
            })
            full_text += run_text

        # 提取段落格式
        paragraph_format = self.format_handler.extract_paragraph_format(paragraph)

        # 执行替换
        replaced = False
        for placeholder, replacement in placeholder_map.items():
            if placeholder in full_text:
                # 首次替换模式：检查是否已经替换过
                if not self.config.replace_all and placeholder in self._replaced_placeholders:
                    continue

                # 根据配置选择替换模式
                if self.config.replace_all:
                    # 全量替换：替换所有匹配项
                    full_text = full_text.replace(placeholder, replacement)
                else:
                    # 首次替换：仅替换第一次出现
                    full_text = full_text.replace(placeholder, replacement, 1)
                    # 记录已替换的占位符
                    self._replaced_placeholders.add(placeholder)

                replaced = True
                self.stats['placeholders_found'] += 1

        if not replaced:
            return False

        # 清空原有run
        for run in paragraph.runs:
            run.text = ""

        # 智能格式合并:合并所有run的格式
        merged_format = self.format_handler.merge_formats(
            [r['format'] for r in runs_data]
        )

        # 创建新run并应用合并后的格式
        if paragraph.runs:
            new_run = paragraph.runs[0]
        else:
            new_run = paragraph.add_run()

        new_run.text = full_text
        self.format_handler.apply_run_format(new_run, merged_format)

        # 应用段落格式
        self.format_handler.apply_paragraph_format(paragraph, paragraph_format)

        return True

    def save(self, output_path: str):
        """
        保存文档

        Args:
            output_path: 输出文件路径
        """
        try:
            output_path = Path(output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            self.doc.save(str(output_path))
            logger.info(f"文档已保存: {output_path}")
        except Exception as e:
            logger.error(f"保存文档失败: {e}")
            raise

    def get_stats(self) -> Dict[str, Any]:
        """
        获取统计信息

        Returns:
            统计信息字典
        """
        return self.stats.copy()
