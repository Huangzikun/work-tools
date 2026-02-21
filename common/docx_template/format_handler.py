"""
格式处理器 - 负责保留和复制文本格式
"""
from docx.text.run import Run
try:
    from docx.text.paragraph import Paragraph
except ImportError:
    from docx import Paragraph
from docx.oxml.ns import qn
from typing import Dict, Any
import logging

from .config import FormatConfig

logger = logging.getLogger(__name__)


class FormatHandler:
    """格式处理器 - 负责格式的提取、保留和应用"""

    def __init__(self, config: FormatConfig):
        self.config = config
        logger.debug(f"初始化FormatHandler, 配置: {config}")

    def extract_run_format(self, run: Run) -> Dict[str, Any]:
        """
        提取run的格式信息

        Args:
            run: docx的Run对象

        Returns:
            包含格式信息的字典
        """
        format_info = {}

        if not self.config.preserve_run_format:
            return format_info

        font = run.font

        # 基础格式属性
        for attr in self.config.run_format_attrs:
            try:
                if attr == 'bold' and font.bold is not None:
                    format_info['bold'] = font.bold
                elif attr == 'italic' and font.italic is not None:
                    format_info['italic'] = font.italic
                elif attr == 'underline' and font.underline is not None:
                    format_info['underline'] = font.underline
                elif attr == 'strike' and font.strike is not None:
                    format_info['strike'] = font.strike
                elif attr == 'all_caps' and font.all_caps is not None:
                    format_info['all_caps'] = font.all_caps
                elif attr == 'small_caps' and font.small_caps is not None:
                    format_info['small_caps'] = font.small_caps
                elif attr == 'superscript' and font.superscript is not None:
                    format_info['superscript'] = font.superscript
                elif attr == 'subscript' and font.subscript is not None:
                    format_info['subscript'] = font.subscript
                elif attr == 'font_name' and font.name:
                    format_info['font_name'] = font.name
                elif attr == 'font_size' and font.size:
                    format_info['font_size'] = font.size
                elif attr == 'color' and font.color and font.color.rgb:
                    format_info['color_rgb'] = font.color.rgb
                elif attr == 'highlight_color' and font.highlight_color and font.highlight_color.rgb:
                    format_info['highlight_color_rgb'] = font.highlight_color.rgb
            except Exception as e:
                logger.warning(f"提取格式属性 {attr} 失败: {e}")

        return format_info

    def extract_paragraph_format(self, paragraph: Paragraph) -> Dict[str, Any]:
        """
        提取段落格式信息

        Args:
            paragraph: docx的Paragraph对象

        Returns:
            包含段落格式信息的字典
        """
        format_info = {}

        if not self.config.preserve_paragraph_format:
            return format_info

        pf = paragraph.paragraph_format

        for attr in self.config.paragraph_format_attrs:
            try:
                value = getattr(pf, attr, None)
                if value is not None:
                    format_info[attr] = value
            except Exception as e:
                logger.warning(f"提取段落格式属性 {attr} 失败: {e}")

        return format_info

    def apply_run_format(self, run: Run, format_info: Dict[str, Any]):
        """
        将格式应用到run

        Args:
            run: docx的Run对象
            format_info: 格式信息字典
        """
        if not format_info:
            return

        font = run.font

        # 应用基础格式
        if 'bold' in format_info:
            font.bold = format_info['bold']
        if 'italic' in format_info:
            font.italic = format_info['italic']
        if 'underline' in format_info:
            font.underline = format_info['underline']
        if 'strike' in format_info:
            font.strike = format_info['strike']
        if 'all_caps' in format_info:
            font.all_caps = format_info['all_caps']
        if 'small_caps' in format_info:
            font.small_caps = format_info['small_caps']
        if 'superscript' in format_info:
            font.superscript = format_info['superscript']
        if 'subscript' in format_info:
            font.subscript = format_info['subscript']
        if 'font_name' in format_info:
            font.name = format_info['font_name']
            # 中文字体支持
            try:
                font._element.rPr.rFonts.set(qn('w:eastAsia'), format_info['font_name'])
            except Exception as e:
                logger.warning(f"设置中文字体失败: {e}")
        if 'font_size' in format_info:
            font.size = format_info['font_size']
        if 'color_rgb' in format_info:
            font.color.rgb = format_info['color_rgb']
        if 'highlight_color_rgb' in format_info:
            font.highlight_color.rgb = format_info['highlight_color_rgb']

    def apply_paragraph_format(self, paragraph: Paragraph, format_info: Dict[str, Any]):
        """
        将段落格式应用到段落

        Args:
            paragraph: docx的Paragraph对象
            format_info: 段落格式信息字典
        """
        if not format_info:
            return

        pf = paragraph.paragraph_format

        for attr, value in format_info.items():
            try:
                setattr(pf, attr, value)
            except Exception as e:
                logger.warning(f"应用段落格式属性 {attr} 失败: {e}")

    def merge_formats(self, format_list: list) -> Dict[str, Any]:
        """
        合并多个格式信息

        策略: 使用第一个非None的格式值(按优先级)

        Args:
            format_list: 格式信息列表

        Returns:
            合并后的格式信息字典
        """
        merged = {}

        for format_info in format_list:
            for key, value in format_info.items():
                if key not in merged and value is not None:
                    merged[key] = value

        return merged
