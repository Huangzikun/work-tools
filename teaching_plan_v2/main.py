"""
教案生成工具V2 - CLI入口
"""
import argparse
import sys
from pathlib import Path

from .config import AppConfig
from .generator import LessonPlanGenerator
from .logger import setup_logging


def main():
    """CLI入口"""
    parser = argparse.ArgumentParser(
        description="教案生成工具V2 - 使用AI从教学大纲自动生成教案",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 基本用法
  python -m teaching_plan_v2.main \\
    --syllabus="数据结构理论课教案-V1.docx" \\
    --total_lessons=9 \\
    --template="理论课教案-V1.docx"

  # 指定输出文件名
  python -m teaching_plan_v2.main \\
    --syllabus="数据结构理论课教案-V1.docx" \\
    --total_lessons=9 \\
    --output="数据结构-完整教案.docx"

  # 指定输出目录
  python -m teaching_plan_v2.main \\
    --syllabus="数据结构理论课教案-V1.docx" \\
    --total_lessons=9 \\
    --output_dir="output/教案"

环境变量:
  ARK_API_KEY  火山方舟API密钥（必需）
        """
    )

    # 必需参数
    parser.add_argument(
        '--syllabus',
        required=True,
        help='教学大纲Word文档路径'
    )
    parser.add_argument(
        '--total_lessons',
        type=int,
        required=True,
        help='总课时数'
    )

    # 可选参数
    parser.add_argument(
        '--template',
        help='Word模板路径（可选，使用配置文件中的默认值）'
    )
    parser.add_argument(
        '--output',
        help='输出文件名（可选，默认为"大纲名-教案.docx"）'
    )
    parser.add_argument(
        '--output_dir',
        help='输出目录（可选，默认为"generated_plans"）'
    )
    parser.add_argument(
        '--batch_size',
        type=int,
        default=2,
        help='每批生成的教案数（默认：2）'
    )
    parser.add_argument(
        '--model',
        default='doubao-seed-2-0-mini-260428',
        help='AI模型名称（默认：doubao-seed-2-0-mini-260428）'
    )
    parser.add_argument(
        '--log_level',
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
        default='INFO',
        help='日志级别（默认：INFO）'
    )

    args = parser.parse_args()

    # 加载配置
    try:
        config = AppConfig.from_env()

        # 命令行参数覆盖配置
        if args.template:
            config.template_path = args.template
        if args.output_dir:
            config.output_dir = args.output_dir
        config.batch_size = args.batch_size
        config.model = args.model
        config.log_level = args.log_level

        # 验证配置
        config.validate()

    except ValueError as e:
        print(f"配置错误: {e}", file=sys.stderr)
        sys.exit(1)

    # 初始化日志
    setup_logging(config.log_level, log_file="teaching_plan_v2.log")

    # 生成教案
    try:
        generator = LessonPlanGenerator(config)
        output_path = generator.generate(
            args.syllabus,
            args.total_lessons,
            args.output
        )

        print(f"\n✓ 教案生成完成！")
        print(f"输出文件: {output_path}")

    except Exception as e:
        print(f"\n✗ 生成失败: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
