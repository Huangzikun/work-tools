"""
工具函数
"""
import re
from pathlib import Path
from typing import List


def sanitize_filename(filename: str) -> str:
    """
    清理文件名中的非法字符

    Args:
        filename: 原始文件名

    Returns:
        清理后的文件名
    """
    # 移除Windows和Linux文件名中的非法字符
    illegal_chars = r'[<>:"/\\|?*]'
    cleaned = re.sub(illegal_chars, '_', filename)

    # 移除首尾空格
    cleaned = cleaned.strip()

    # 限制文件名长度
    max_length = 200
    if len(cleaned) > max_length:
        cleaned = cleaned[:max_length]

    return cleaned


def calculate_time_distribution(total_minutes: int) -> dict:
    """
    计算默认时间分配

    Args:
        total_minutes: 总分钟数

    Returns:
        时间分配字典
    """
    # 默认分配比例
    import_ratio = 0.05  # 课堂导入 5%
    content_ratio = 0.90  # 课堂内容 90%
    summary_ratio = 0.05  # 课堂小结 5%

    return {
        '课堂导入': int(total_minutes * import_ratio),
        '课堂内容': int(total_minutes * content_ratio),
        '课堂小结': int(total_minutes * summary_ratio),
    }
