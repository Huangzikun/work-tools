"""
word_utils.py

此模块提供了一个 Word 文档操作类，用于替换文档中的指定文本并保留格式。
"""
from docx import Document

class WordReplacer:
    """
    用于处理 Word 文档文本替换的类，支持将文档中匹配指定键的文本替换为对应的值，同时保留原有的格式。
    """
    def __init__(self, file_path):
        """
        初始化 WordReplacer 类。

        :param file_path: 要处理的 Word 文档的文件路径。
        """
        self.doc = Document(file_path)

    def replace_text(self, replacement_map):
        """
        替换 Word 文档中的文本。

        :param replacement_map: 一个字典，键为要替换的文本，值为替换后的文本。
        """
        wrapped_replacement_map = {f"{{{key}}}": value for key, value in replacement_map.items()}

        for paragraph in self.doc.paragraphs:
            self._replace_in_paragraph(paragraph, wrapped_replacement_map)
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._replace_in_paragraph(paragraph, wrapped_replacement_map)

    def _replace_in_paragraph(self, paragraph, replacement_map):
        """
        在单个段落中执行文本替换操作。

        :param paragraph: 要处理的段落对象。
        :param replacement_map: 一个字典，键为要替换的文本，值为替换后的文本。
        """
        for key, value in replacement_map.items():
            if key in paragraph.text:
                inline = paragraph.runs
                for i in range(len(inline)):
                    if key in inline[i].text:
                        text = inline[i].text.replace(key, value)
                        inline[i].text = text

    def save(self, output_path):
        """
        将处理后的 Word 文档保存到指定路径。

        :param output_path: 保存处理后文档的文件路径。
        """
        self.doc.save(output_path)