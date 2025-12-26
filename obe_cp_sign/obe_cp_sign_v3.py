import os
import subprocess

import pandas as pd
import shutil
import argparse
import zipfile
import hashlib
from volcenginesdkarkruntime import Ark
import json
from docx import Document
from docx.shared import Cm
from pathlib import Path
from docx.text.paragraph import Paragraph

"""
将单个 .doc 文件转换为 .docx 格式，并删除原文件。
:param file_path: .doc 文件路径
:return: 转换后的 .docx 文件路径 或 None
"""
def doc_to_docx(file_path):
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
            '--writer',  # 强制使用文字处理模式
            '--nocrashreport',  # 禁止崩溃报告
            '--nodefault',  # 不加载默认模板
            '--norestore',  # 不恢复上次会话
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


def extract_non_table_text(doc):
    non_table_text = []

    # 遍历文档主体中的所有元素
    for element in doc.element.body:
        # 检查元素是否为段落（tag以'p'结尾，对应<w:p>标签）
        if element.tag.endswith('p'):
            # 直接从段落元素创建Paragraph对象
            paragraph = Paragraph(element, doc)
            # 只添加非空文本（可选，根据需求调整）
            if paragraph.text.strip():
                non_table_text.append(paragraph.text)

    return "\n".join(non_table_text)

def extract_table_text(doc):

    full_text = []

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    full_text.append(paragraph.text)

    return "\n".join(list(dict.fromkeys(full_text)))


def add_teacher_review_row(doc):

    for table in doc.tables:
        for row in table.rows:
            for i in range(len(row.cells)):

                if str.strip(row.cells[i].text) == '教师评阅':
                    return


    """在"结果分析与思考"后添加"教师评阅"行，并合并第二列单元格"""
    for table in doc.tables:
        for i, row in enumerate(table.rows):
            row_text = "".join([cell.text.strip() for cell in row.cells])

            #加在这后面
            if "结果分析与思考" in row_text:
                # 插入新行
                new_row = table.add_row()

                # 设置第一列内容
                new_row.cells[0].text = "教师评阅"

                # 合并第二列单元格（假设表格至少有两列）
                if len(table.columns) > 1:
                    # 合并当前行的第2列到最后一列
                    new_row.cells[1].merge(new_row.cells[-1])

                print("已添加并合并教师评阅行")
                return  # 只处理第一个匹配的表格
    print("未找到包含'结果分析与思考'的行")


def sign_by_picture(file_path, save_path, score, review):
    if not file_path.endswith(".docx"):
        print(f"不是docx文件无法签名. file_path={file_path}")
        return 0
    doc = Document(file_path)

    # 先添加教师评阅行并合并单元格
    add_teacher_review_row(doc)

    if len(doc.tables) <= 0:
        print(f"can't sign file. file = {file_path}")
        return 0

    for table in doc.tables:
        for row in table.rows:
            for i in range(len(row.cells)):

                if str.strip(row.cells[i].text) == '教师评阅':

                    # 确保有足够的单元格
                    if i + 1 < len(row.cells):
                        row.cells[i + 1].text = (f"{review}"
                                                 f"成绩：{score}分"
                                                 f"{os.linesep}"
                                                 f"{os.linesep}"
                                                 f"{os.linesep}"
                                                 f"{os.linesep}"
                                                 f"{os.linesep}"
                                                 f"                       教师签名：")

                        # 获取这个 cell 中的段落（此时已经有默认段落）
                        paragraph = row.cells[i + 1].paragraphs[-1]  # 使用最后一个段落，即刚设置 text 的那个段落
                        paragraph.add_run().add_picture(sign_picture, width=Cm(2))
                        paragraph.add_run(f"   {sign_date}")
                        doc.save(save_path)
                        return
                    else:
                        print(f"表格列数不足，无法在'教师评阅'后添加内容")

    print(f"sign fail. file = {file_path}")

def calculate_sha256(file_path):
    hash_object = hashlib.sha256()
    with open(file_path, 'rb') as f:
        for chunk in iter(lambda: f.read(4096), b''):
            hash_object.update(chunk)
    return hash_object.hexdigest()


def file_name_index(file, old_path, destination_path):
    file_list = os.listdir(destination_path)
    file_list = [f for f in file_list if not f.startswith('.')]

    file_md5_list = {}
    for old_file in file_list:
        file_md5_list[calculate_sha256(os.path.join(destination_path, old_file))] = old_file

    old_md5 = calculate_sha256(os.path.join(old_path, file))

    #覆盖
    if old_md5 in file_md5_list:
        return os.path.join(destination_path, file_md5_list[old_md5])
    else:
        file_count = len(file_list)+1
        return os.path.join(destination_path, add_suffix_before_extension(file, file_count))


def add_suffix_before_extension(file_name, suffix):
    """
    将后缀添加在原文件名与原后缀之间
    """
    base_name, ext = file_name.rsplit('.', 1)
    new_file_name = f"{base_name}_{suffix}.{ext}"
    return new_file_name

def is_archive(file_path):
    archive_extensions = (
        '.zip'
    )

    return file_path.lower().endswith(archive_extensions)

def score_and_sign(destination_path_file):
    if not destination_path_file.endswith("docx"):
        print(f"不是docx文件无法签名. file_path={destination_path_file}")
        return 0

    retry = 3
    while retry > 0:
        print(f"retry={retry}")
        try:
            document = Document(destination_path_file)
            document_text = extract_table_text(document)

            completion = client.chat.completions.create(
                # 指定您创建的方舟推理接入点 ID，此处已帮您修改为您的推理接入点 ID
                model="doubao-seed-1-6-250615",
                messages=[
                    {"role": "system",
                     "content": "你是桂林学院信息工程学院的一名计算机专任教师，你拥有丰富的教学经验。你的任务是针对学生提交的实验报告进行批改。你应该理解、使用用户提交的"教师要求"部分对"学生作答"部分进行批阅。你可以选择的分数为60,70,80,90和100分，并给出一个50字以内的批阅评语。生成json格式的内容，包含一个score和一个comment字段。"},
                    {"role": "user", "content": f"教师要求:{teacher};学生作答:{document_text}"},
                ],
                response_format={
                    "type": "json_object",
                },
            )

            print(completion.choices[0].message.content)
            result = json.loads(completion.choices[0].message.content)
            sign_by_picture(destination_path_file, destination_path_file, result["score"], result["comment"])
            return result["score"]
        except Exception as e:
            print(f"签名失败. file_path={destination_path_file}, error={e}")
            retry = retry - 1

    return 0

def copy_student_file(file, old_path, destination_path):

    old_full_path = os.path.join(old_path, file)
    if is_archive(old_full_path):
            # 创建临时解压目录
            temp_dir = os.path.join(old_path, f"{file}.extracted")
            os.makedirs(temp_dir, exist_ok=True)

            try:
                # 解压ZIP文件
                with zipfile.ZipFile(old_full_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)

                # 递归复制解压后的所有文件
                temp_score = 0
                for item in os.listdir(temp_dir):
                    temp_score = max(temp_score, copy_student_file(item, temp_dir, destination_path))
                return temp_score

            finally:
                # 清理临时目录
                shutil.rmtree(temp_dir)
                print(f"已清理临时目录: {temp_dir}")

    else:
        if os.path.isfile(old_full_path):

            destination_path_file = file_name_index(file, old_path, destination_path)
            #只要文档
            if not (destination_path_file.endswith("docx") or destination_path_file.endswith("doc")):
                return 0
            try:
                doc = Document(old_full_path)
                document_text = extract_non_table_text(doc)

                if "实验实训报告" not in document_text:
                    print("不是实验实训报告格式，跳过")
                    return 0

                shutil.copy(old_full_path, destination_path_file)
                print(f"新文件名：{destination_path_file}")

                if destination_path_file.endswith(".doc"):
                    destination_path_file = doc_to_docx(destination_path_file)

                return score_and_sign(destination_path_file)
            except Exception as e:
                print(f"签名失败. file_path={destination_path_file}, error={e}")
                return 0
        else:
            new_old_path =os.path.join(old_path, file)
            file_list = os.listdir(new_old_path)
            print("file_list=", file_list)
            max_score = 0
            for file in file_list:
                max_score = max(max_score, copy_student_file(file, new_old_path, destination_path))
            return max_score



##TEST

'''
执行内容
'''
parser = argparse.ArgumentParser()
parser.add_argument('--old_dir', type=str, help='旧路径', required=True)
parser.add_argument('--new_dir', type=str, help='新路径', required=True)
parser.add_argument('--directory', type=str, help='路径', required=True)
parser.add_argument('--sign_picture', type=str, help='签名图片地址', required=True)
parser.add_argument('--sign', type=str, help='签名字符串，如某某某', required=True)
parser.add_argument('--sign_date', type=str, help='签名时间', required=True)

#获取参数
args = parser.parse_args()
old_dir = args.old_dir
new_dir = args.new_dir
directory = args.directory
sign_picture = args.sign_picture
sign = args.sign
sign_date = args.sign_date

#teacher = "实验报告针对A*算法应用在8数码问题上的图遍历过程，目的应包含A*算法如何解决该问题；实验原理应包括A*算法的思想；结果分析应对A*算法进行总结。若实验目的、原理、结果分析均完善，应得100分；某一项有内容但不完整，应的90分；缺少某一项应得80分；缺少两项及以上应得70分。"
teacher = '''
实验报告基于Java，实验报告整体内容应包含通过实验实现英汉互译系统的开发，包括原理、实现和分析；
实验原理应包括使用Java基于HashMap实现英汉互译系统的开发；
结果分析应对代码进行总结和分析。
若实验目的、原理、结果分析均完善，应得100分；某一大项有内容但内容不完整，应得90分；
缺少实验目的、原理、结果的某一大项应得80分；缺少实验目的、原理、结果中的两项应得70分；
缺少实验目的、原理、结果应得50分。
在不影响判断要求的情况下尽可能匹配高分。
'''
# 请确保您已将 API Key 存储在环境变量 ARK_API_KEY 中
# 初始化Ark客户端，从环境变量中读取您的API Key
client = Ark(
    # 此为默认路径，您可根据业务所在地域进行配置
    base_url="https://ark.cn-beijing.volces.com/api/v3",
    # 从环境变量中获取您的 API Key。此为默认方式，您可根据需要进行修改
    api_key=os.environ.get("ARK_API_KEY"),
)

df = pd.read_excel(os.path.join(directory, '名单.xlsx'), names=['student_id','student_name', "1", "2"])
file_list = os.listdir(old_dir)

df = df.apply(lambda x: x.str.replace('\t', ''))
df = df[df['student_id'].str.match(r'^\d+$')]

file_count = 0
score_list = []

student_path_list = os.listdir(new_dir)
for index, row in df.iterrows():
    student_id = str(row['student_id'])
    student_name = row['student_name']

    use_student_path = ''
    score = 0
    for student_path in student_path_list:
        if student_id in student_path:
            use_student_path = os.path.join(new_dir, student_path)
            break

    for file in file_list:
        if student_id in file:
            score = copy_student_file(file, old_dir, use_student_path)
            print(f"复制{student_id}")
            file_count = file_count + 1
            break
        elif student_name in file:
            score = copy_student_file(file, old_dir, use_student_path)
            print(f"复制{student_id}")
            file_count = file_count + 1
            break

    print(f"score:{score}")
    score_list.append({
        'student_id': student_id,
        'student_name': student_name,
        "score": score,
    })

print(f"复制成功{file_count}")
pd.DataFrame(score_list).to_excel(os.path.join(directory, '实验报告3.xlsx'), index=False)
