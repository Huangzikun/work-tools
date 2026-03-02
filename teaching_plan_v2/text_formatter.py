"""
文本格式化工具 - 用于美化教案内容格式
"""
import re
import logging

logger = logging.getLogger(__name__)


def format_lesson_content(text: str) -> str:
    """
    格式化课堂内容文本，在编号步骤之间添加换行

    Args:
        text: 原始文本

    Returns:
        格式化后的文本

    Examples:
        >>> format_lesson_content("1. 第一步2. 第二步3. 第三步")
        "1. 第一步\\n\\n2. 第二步\\n\\n3. 第三步"
    """
    if not text:
        return text

    # 匹配编号模式：数字加点+空格，或中文数字+点+空格
    # 例如: "1. "、"2. "、"一. "、"二. "、"（1）"、"（一）"
    patterns = [
        r'(\d+\. )',  # "1. "、"2. "
        r'([一二三四五六七八九十]+\. )',  # "一. "、"二. "
        r'(\([一二三四五六七八九十]+\)\s*)',  # "（一）"、"（二）"
        r'(\([0-9]+\)\s*)',  # "(1) "、"(2) "
    ]

    result = text

    # 对每个模式进行处理
    for pattern in patterns:
        # 在匹配的编号前添加两个换行符（保持美观的段落间距）
        # 使用正向回顾断言，确保不会在文本开头添加换行
        result = re.sub(r'(?<!\n)' + pattern, r'\n\n\1', result)

    # 清理多余的空行（超过2个连续换行的压缩为2个）
    result = re.sub(r'\n{3,}', '\n\n', result)

    # 清理开头的换行符
    result = result.lstrip('\n')

    logger.debug(f"格式化文本: 原长度={len(text)}, 格式化后长度={len(result)}")

    return result


def format_all_lesson_data(data: dict) -> dict:
    """
    格式化教案数据中的所有文本字段

    Args:
        data: 教案数据字典

    Returns:
        格式化后的数据字典
    """
    formatted_data = {}

    # 需要格式化的字段列表
    fields_to_format = [
        '课堂内容',
        '课堂导入',
        '教学方法与设计',
        '课堂小结',
        '课后作业',
    ]

    for key, value in data.items():
        if key in fields_to_format and isinstance(value, str):
            # 对指定字段进行格式化
            formatted_data[key] = format_lesson_content(value)
        else:
            # 其他字段保持不变
            formatted_data[key] = value

    return formatted_data
