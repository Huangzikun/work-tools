"""文本格式化（直接复制自 teaching_plan_v2/text_formatter.py）。"""

import re


def format_lesson_content(text: str) -> str:
    if not text:
        return text

    patterns = [
        r"(\d+\. )",
        r"([一二三四五六七八九十]+\. )",
        r"(\([一二三四五六七八九十]+\)\s*)",
        r"(\([0-9]+\)\s*)",
    ]

    result = text
    for pattern in patterns:
        result = re.sub(r"(?<!\n)" + pattern, r"\n\n\1", result)

    result = re.sub(r"\n{3,}", "\n\n", result)
    return result.lstrip("\n")


def format_all_lesson_data(data: dict) -> dict:
    formatted_data = {}
    fields_to_format = [
        "课堂内容",
        "课堂导入",
        "教学方法与设计",
        "课堂小结",
        "课后作业",
    ]
    for key, value in data.items():
        if key in fields_to_format and isinstance(value, str):
            formatted_data[key] = format_lesson_content(value)
        else:
            formatted_data[key] = value
    return formatted_data
