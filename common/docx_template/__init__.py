"""
DOCX模板替换通用组件

快速开始:
    from common.docx_template import DocxTemplateReplacer

    # 创建替换器
    replacer = DocxTemplateReplacer('template.docx')

    # 执行替换
    data = {
        '姓名': '张三',
        '学号': '202213008001',
        '课程': '数据结构'
    }
    replacer.replace(data, 'output.docx')
"""

from .core import DocxTemplateReplacer
from .config import ReplacerConfig, PlaceholderConfig, FormatConfig
from .exceptions import (
    DocxTemplateError,
    PlaceholderNotFoundError,
    TemplateLoadError,
    InvalidDataError
)

__version__ = '1.0.0'
__all__ = [
    'DocxTemplateReplacer',
    'ReplacerConfig',
    'PlaceholderConfig',
    'FormatConfig',
    'DocxTemplateError',
    'PlaceholderNotFoundError',
    'TemplateLoadError',
    'InvalidDataError'
]
