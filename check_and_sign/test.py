from pathlib import Path
import subprocess
import os


def doc_to_docx(file_path):
    """
    将单个 .doc 文件转换为 .docx 格式，并删除原文件。
    :param file_path: .doc 文件路径
    :return: 转换后的 .docx 文件路径 或 None
    """
    try:
        # 规范化路径
        file_path = Path(file_path).resolve()
        if not file_path.exists():
            raise FileNotFoundError(f"文件不存在: {file_path}")

        # 检查扩展名是否正确
        if file_path.suffix.lower() != ".doc":
            raise ValueError(f"仅支持 .doc 文件: {file_path}")

        print(f"正在转换: {file_path}")

        # 构建输出目录和目标路径
        output_dir = file_path.parent
        new_file_path = file_path.with_suffix(".docx")

        # 构建命令
        command = [
            'libreoffice',
            '--headless',  # 无界面运行
            '--convert-to', 'docx',  # 转换为目标格式
            '--outdir', str(output_dir),  # 输出目录
            str(file_path)  # 输入文件
        ]

        # 执行转换
        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=True
        )

        # 检查转换结果
        if new_file_path.exists():
            file_path.unlink()
            print(f"成功转换并删除原文件: {new_file_path}")
            return str(new_file_path)
        else:
            raise Exception("转换失败，未生成 .docx 文件")

    except subprocess.CalledProcessError as e:
        print(f"LibreOffice 执行失败: {e.stderr}")
        return None
    except Exception as e:
        print(f"文件转换失败: {e}")
        return None


def batch_convert_doc_to_docx(folder_path):
    """
    批量转换指定文件夹下的所有 .doc 文件
    :param folder_path: 包含 .doc 文件的文件夹路径
    """
    folder = Path(folder_path).resolve()
    if not folder.is_dir():
        print(f"不是有效目录: {folder}")
        return

    doc_files = list(folder.rglob("*.doc"))  # 递归查找所有 .doc 文件
    print(f"找到 {len(doc_files)} 个 .doc 文件准备转换...")

    for f in doc_files:
        doc_to_docx(f)


if __name__ == "__main__":
    # 示例：单个文件转换
    test_file = "/Users/huangzikun/Desktop/桂林学院/算法设计与分析/期末材料/计科/2022计算机科学与技术《算法设计与分析》实验报告黄子坤55份/202213008205计算机科学与技术罗瑛子/202213008205罗瑛子1_1.doc"

    # 示例：批量转换某个文件夹下的所有 .doc 文件
    test_folder = "/Users/yourname/Desktop/文档集"

    # 单个文件测试
    result = doc_to_docx(test_file)
    if result:
        print(f"🎉 转换成功: {result}")
    else:
        print("❌ 转换失败，请检查错误信息")

    # 批量转换测试（取消注释即可）
    # batch_convert_doc_to_docx(test_folder)