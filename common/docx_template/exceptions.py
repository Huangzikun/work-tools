"""
自定义异常类
"""


class DocxTemplateError(Exception):
    """基础异常类"""
    pass


class PlaceholderNotFoundError(DocxTemplateError):
    """占位符未找到异常"""
    pass


class TemplateLoadError(DocxTemplateError):
    """模板加载异常"""
    pass


class InvalidDataError(DocxTemplateError):
    """数据格式错误异常"""
    pass


class FormatPreservationError(DocxTemplateError):
    """格式保留异常"""
    pass
