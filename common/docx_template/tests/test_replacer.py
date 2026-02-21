"""
单元测试 - DocxTemplateReplacer
"""
import unittest
import sys
from pathlib import Path

# 添加项目根目录到路径
project_root = Path(__file__).parent.parent.parent.parent
sys.path.insert(0, str(project_root))

from common.docx_template import DocxTemplateReplacer, ReplacerConfig
from common.docx_template.exceptions import TemplateLoadError, InvalidDataError


class TestDocxTemplateReplacer(unittest.TestCase):
    """测试DocxTemplateReplacer类"""

    def setUp(self):
        """每个测试前的设置"""
        self.fixture_dir = Path(__file__).parent / 'fixtures'
        self.output_dir = Path(__file__).parent / 'temp_output'
        self.output_dir.mkdir(exist_ok=True)

    def tearDown(self):
        """清理临时文件"""
        import shutil
        if self.output_dir.exists():
            shutil.rmtree(self.output_dir)

    def test_simple_replace(self):
        """测试基本替换功能"""
        template_path = self.fixture_dir / 'template_simple.docx'
        output_path = self.output_dir / 'test_simple_output.docx'

        replacer = DocxTemplateReplacer(str(template_path))
        data = {
            '姓名': '张三',
            '学号': '202213008001',
            '课程': '数据结构',
            '课时': '64'
        }
        replacer.replace(data, str(output_path))

        # 验证输出文件存在
        self.assertTrue(output_path.exists())

        # 验证替换成功
        from docx import Document
        doc = Document(str(output_path))
        text = '\n'.join([p.text for p in doc.paragraphs])
        self.assertIn('张三', text)
        self.assertIn('202213008001', text)
        self.assertIn('数据结构', text)
        self.assertIn('64', text)

    def test_replace_all_mode(self):
        """测试全量替换模式"""
        template_path = self.fixture_dir / 'template_multiple.docx'
        output_path = self.output_dir / 'test_all_output.docx'

        # 默认全量替换
        replacer = DocxTemplateReplacer(str(template_path))
        data = {'姓名': '李四', '学号': '202213008002', '课程': 'Python程序设计'}
        replacer.replace(data, str(output_path))

        # 验证所有出现的占位符都被替换
        from docx import Document
        doc = Document(str(output_path))
        text = '\n'.join([p.text for p in doc.paragraphs])

        # 检查李四出现3次（初始1次 + 授课教师1次 + 班主任1次）
        self.assertEqual(text.count('李四'), 3)

    def test_replace_first_mode(self):
        """测试首次替换模式"""
        template_path = self.fixture_dir / 'template_multiple.docx'
        output_path = self.output_dir / 'test_first_output.docx'

        # 配置为首次替换模式
        config = ReplacerConfig(replace_all=False)
        replacer = DocxTemplateReplacer(str(template_path), config)
        data = {'姓名': '王五', '学号': '202213008003', '课程': 'Java程序设计'}
        replacer.replace(data, str(output_path))

        # 验证只有第一次出现的占位符被替换
        from docx import Document
        doc = Document(str(output_path))
        text = '\n'.join([p.text for p in doc.paragraphs])

        # 检查王五只出现1次（只有第一次）
        self.assertEqual(text.count('王五'), 1)
        # 检查还有未替换的占位符
        self.assertIn('{{姓名}}', text)

    def test_table_replace(self):
        """测试表格替换"""
        template_path = self.fixture_dir / 'template_table.docx'
        output_path = self.output_dir / 'test_table_output.docx'

        replacer = DocxTemplateReplacer(str(template_path))
        data = {
            '课程': '算法设计与分析',
            '教师': '赵老师',
            '课时': '48',
            '学分': '3.0',
            '类型': '专业必修课'
        }
        replacer.replace(data, str(output_path))

        # 验证输出文件存在
        self.assertTrue(output_path.exists())

        # 验证表格中的替换成功
        from docx import Document
        doc = Document(str(output_path))
        table = doc.tables[0]
        self.assertEqual(table.rows[0].cells[1].text, '算法设计与分析')
        self.assertEqual(table.rows[1].cells[1].text, '赵老师')

    def test_format_preservation(self):
        """测试格式保留"""
        template_path = self.fixture_dir / 'template_simple.docx'
        output_path = self.output_dir / 'test_format_output.docx'

        replacer = DocxTemplateReplacer(str(template_path))
        data = {'姓名': '孙七', '学号': '202213008004', '课程': '操作系统', '课时': '56'}
        replacer.replace(data, str(output_path))

        # 验证格式保留（第一段的姓名应该是粗体）
        from docx import Document
        doc = Document(str(output_path))

        # 检查第一个段落（姓名）的格式
        first_para = doc.paragraphs[1]  # 跳过标题
        if first_para.runs:
            # 验证格式被保留
            self.assertIsNotNone(first_para.runs[0].font.size)

    def test_invalid_template(self):
        """测试无效模板"""
        with self.assertRaises(TemplateLoadError):
            DocxTemplateReplacer('/nonexistent/template.docx')

    def test_invalid_data_type(self):
        """测试无效数据类型"""
        template_path = self.fixture_dir / 'template_simple.docx'
        replacer = DocxTemplateReplacer(str(template_path))

        with self.assertRaises(InvalidDataError):
            replacer.replace("not a dict")

    def test_empty_data(self):
        """测试空数据"""
        template_path = self.fixture_dir / 'template_simple.docx'
        output_path = self.output_dir / 'test_empty_output.docx'

        replacer = DocxTemplateReplacer(str(template_path))
        replacer.replace({}, str(output_path))

        # 验证文件被创建但没有替换
        self.assertTrue(output_path.exists())
        stats = replacer.get_stats()
        self.assertEqual(stats['placeholders_replaced'], 0)

    def test_get_stats(self):
        """测试统计信息"""
        template_path = self.fixture_dir / 'template_simple.docx'
        output_path = self.output_dir / 'test_stats_output.docx'

        replacer = DocxTemplateReplacer(str(template_path))
        data = {'姓名': '周八', '学号': '202213008005', '课程': '计算机网络', '课时': '48'}
        replacer.replace(data, str(output_path))

        stats = replacer.get_stats()
        self.assertIn('placeholders_found', stats)
        self.assertIn('placeholders_replaced', stats)
        self.assertGreater(stats['placeholders_found'], 0)


if __name__ == '__main__':
    unittest.main()
