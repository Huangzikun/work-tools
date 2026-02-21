"""
工具函数
"""
import re
import logging
from typing import Dict, Any

logger = logging.getLogger(__name__)


def build_placeholder_pattern(prefix: str, suffix: str) -> re.Pattern:
    """
    构建占位符正则表达式

    Args:
        prefix: 占位符前缀
        suffix: 占位符后缀

    Returns:
        编译后的正则表达式对象
    """
    pattern = re.escape(prefix) + r'([^' + re.escape(suffix + prefix) + r']+)' + re.escape(suffix)
    return re.compile(pattern)


def validate_data(data: Dict[str, Any]) -> bool:
    """
    验证数据字典

    Args:
        data: 数据字典

    Returns:
        bool: 验证是否通过

    Raises:
        ValueError: 验证失败时抛出异常
    """
    if not isinstance(data, dict):
        raise ValueError("数据必须是字典类型")

    for key, value in data.items():
        if not isinstance(key, str):
            raise ValueError(f"键必须是字符串类型: {key}")
        if value is None:
            logger.warning(f"键 {key} 的值为None,将被转换为空字符串")

    return True


def extract_placeholders(text: str, prefix: str = "{{", suffix: str = "}}") -> list:
    """
    从文本中提取所有占位符

    Args:
        text: 文本内容
        prefix: 占位符前缀
        suffix: 占位符后缀

    Returns:
        list: 占位符列表(不含前后缀)
    """
    pattern = re.escape(prefix) + r'(.*?)' + re.escape(suffix)
    matches = re.findall(pattern, text)
    return matches


def setup_logging(level: str = "INFO"):
    """
    配置日志

    Args:
        level: 日志级别
    """
    numeric_level = getattr(logging, level.upper(), logging.INFO)
    logging.basicConfig(
        level=numeric_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
