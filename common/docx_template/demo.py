"""
DOCX模板替换组件演示脚本

演示功能：
1. 全量替换模式
2. 首次替换模式
"""
import sys
from pathlib import Path

# 添加项目根目录到路径
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from common.docx_template import DocxTemplateReplacer, ReplacerConfig


def demo_replace_all():
    """演示全量替换模式"""
    print("=" * 50)
    print("演示1：全量替换模式（默认）")
    print("=" * 50)

    template_path = Path(__file__).parent / 'tests' / 'fixtures' / 'template_multiple.docx'
    output_path = Path(__file__).parent / 'demo_output_all.docx'

    # 创建替换器（默认全量替换）
    replacer = DocxTemplateReplacer(str(template_path))

    # 执行替换
    data = {
        '姓名': '张三',
        '学号': '202213008001',
        '课程': '数据结构'
    }

    print(f"模板文件: {template_path}")
    print(f"数据: {data}")
    print(f"替换模式: 全量替换")

    replacer.replace(data, str(output_path))

    stats = replacer.get_stats()
    print(f"替换完成！输出文件: {output_path}")
    print(f"统计信息: {stats}")
    print()


def demo_replace_first():
    """演示首次替换模式"""
    print("=" * 50)
    print("演示2：首次替换模式")
    print("=" * 50)

    template_path = Path(__file__).parent / 'tests' / 'fixtures' / 'template_multiple.docx'
    output_path = Path(__file__).parent / 'demo_output_first.docx'

    # 配置为首次替换模式
    config = ReplacerConfig(replace_all=False)
    replacer = DocxTemplateReplacer(str(template_path), config)

    # 执行替换
    data = {
        '姓名': '李四',
        '学号': '202213008002',
        '课程': 'Python程序设计'
    }

    print(f"模板文件: {template_path}")
    print(f"数据: {data}")
    print(f"替换模式: 首次替换（每个占位符只替换第一次出现）")

    replacer.replace(data, str(output_path))

    stats = replacer.get_stats()
    print(f"替换完成！输出文件: {output_path}")
    print(f"统计信息: {stats}")
    print()


def demo_table_replace():
    """演示表格替换"""
    print("=" * 50)
    print("演示3：表格替换")
    print("=" * 50)

    template_path = Path(__file__).parent / 'tests' / 'fixtures' / 'template_table.docx'
    output_path = Path(__file__).parent / 'demo_output_table.docx'

    # 创建替换器
    replacer = DocxTemplateReplacer(str(template_path))

    # 执行替换
    data = {
        '课程': '算法设计与分析',
        '教师': '王老师',
        '课时': '48',
        '学分': '3.0',
        '类型': '专业必修课'
    }

    print(f"模板文件: {template_path}")
    print(f"数据: {data}")
    print(f"替换模式: 全量替换")

    replacer.replace(data, str(output_path))

    stats = replacer.get_stats()
    print(f"替换完成！输出文件: {output_path}")
    print(f"统计信息: {stats}")
    print()


if __name__ == '__main__':
    print("\n" + "=" * 50)
    print("DOCX模板替换通用组件 - 演示脚本")
    print("=" * 50)
    print()

    # 运行演示
    demo_replace_all()
    demo_replace_first()
    demo_table_replace()

    print("=" * 50)
    print("演示完成！")
    print("=" * 50)
    print()
    print("生成的文件:")
    print("  - demo_output_all.docx (全量替换)")
    print("  - demo_output_first.docx (首次替换)")
    print("  - demo_output_table.docx (表格替换)")
