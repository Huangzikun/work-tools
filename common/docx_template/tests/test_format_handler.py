"""
单元测试 - FormatHandler
"""
import unittest
import sys
from pathlib import Path

# 添加项目根目录到路径
project_root = Path(__file__).parent.parent.parent.parent
sys.path.insert(0, str(project_root))

from docx import Document
from common.docx_template.format_handler import FormatHandler
from common.docx_template.config import FormatConfig


class TestFormatHandler(unittest.TestCase):
    """测试FormatHandler类"""

    def setUp(self):
        """每个测试前的设置"""
        self.config = FormatConfig()
        self.handler = FormatHandler(self.config)

    def test_extract_format(self):
        """测试格式提取"""
        # 创建一个带格式的文档
        doc = Document()
        p = doc.add_paragraph()
        run = p.add_run('测试文本')
        run.bold = True
        run.italic = True
        run.font.size = 200000  # 10pt

        # 提取格式
        format_info = self.handler.extract_run_format(run)

        # 验证
        self.assertTrue(format_info.get('bold'))
        self.assertTrue(format_info.get('italic'))
        self.assertIsNotNone(format_info.get('font_size'))

    def test_apply_format(self):
        """测试格式应用"""
        doc = Document()
        p = doc.add_paragraph()
        run = p.add_run('原始文本')

        # 应用格式
        format_info = {
            'bold': True,
            'italic': True,
            'underline': True
        }
        self.handler.apply_run_format(run, format_info)

        # 验证
        self.assertTrue(run.bold)
        self.assertTrue(run.italic)
        self.assertTrue(run.underline)

    def test_merge_formats(self):
        """测试格式合并"""
        format_list = [
            {'bold': True},  # 第一个格式
            {'italic': True, 'underline': True},  # 第二个格式
            {'font_size': 200000}  # 第三个格式
        ]

        merged = self.handler.merge_formats(format_list)

        # 验证：应该使用第一个非None的值
        self.assertTrue(merged['bold'])  # 从第一个获取
        self.assertTrue(merged['italic'])  # 从第二个获取
        self.assertTrue(merged['underline'])  # 从第二个获取
        self.assertIsNotNone(merged['font_size'])  # 从第三个获取


if __name__ == '__main__':
    unittest.main()
