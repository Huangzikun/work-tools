import os
import pandas as pd
import shutil
import argparse
import zipfile
import hashlib
from volcenginesdkarkruntime import Ark
import json
from docx import Document
from docx.shared import Cm

def extract_table_text(doc):

    full_text = []

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    full_text.append(paragraph.text)

    return "\n".join(list(dict.fromkeys(full_text)))

def sign_by_picture(file_path, save_path, score, review):
    if not file_path.endswith(".docx"):
        print(f"不是docx文件无法签名. file_path={file_path}")
        return 0
    doc = Document(file_path)
    if len(doc.tables) <= 0:
        print(f"can't sign file. file = {file_path}")
        return 0

    for table in doc.tables:
        for row in table.rows:
            for i in range(len(row.cells)):

                if str.strip(row.cells[i].text) == '教师评阅':

                    row.cells[i + 1].text = (f"{review}"
                                             f"成绩：{score}分"
                                             f"{os.linesep}"
                                             f"{os.linesep}"
                                             f"{os.linesep}"
                                             f"{os.linesep}"
                                             f"{os.linesep}"
                                             f"                       教师签名：{sign}   {sign_date}")
                    row.cells[i + 1].add_paragraph().add_run().add_picture(sign_picture, width=Cm(2))
                    doc.save(save_path)
                    return

    print(f"sign fail. file = {file_path}")

def calculate_sha256(file_path):
    hash_object = hashlib.sha256()
    with open(file_path, 'rb') as f:
        for chunk in iter(lambda: f.read(4096), b''):
            hash_object.update(chunk)
    return hash_object.hexdigest()


def file_name_index(file, old_path, destination_path):
    file_list = os.listdir(destination_path)

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
    try:
        document = Document(destination_path_file)
        document_text = extract_table_text(document)

        completion = client.chat.completions.create(
            # 指定您创建的方舟推理接入点 ID，此处已帮您修改为您的推理接入点 ID
            model="doubao-seed-1-6-250615",
            messages=[
                {"role": "system",
                 "content": "你是桂林学院信息工程学院的一名计算机专任教师，你拥有丰富的教学经验。你的任务是针对学生提交的实验报告进行批改。你应该理解、使用用户提交的“教师要求”部分对“学生作答”部分进行批阅。你可以选择的分数为70,80,90和100分，并给出一个50字以内的批阅评语。生成json格式的内容，包含一个score和一个comment字段。"},
                {"role": "user", "content": f"教师要求:{teacher};学生作答:{document_text}"},
            ],
            response_format={
                "type": "json_object",
            }
        )

        print(completion.choices[0].message.content)
        result = json.loads(completion.choices[0].message.content)
        sign_by_picture(destination_path_file, destination_path_file, result["score"], result["comment"])
        return result["score"]
    except Exception as e:
        print(f"签名失败. file_path={destination_path_file}")
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
                for item in os.listdir(temp_dir):
                    return copy_student_file(item, temp_dir, destination_path)
            finally:
                # 清理临时目录
                shutil.rmtree(temp_dir)
                print(f"已清理临时目录: {temp_dir}")

    else:
        if os.path.isfile(old_full_path):
            destination_path_file = file_name_index(file, old_path, destination_path)

            shutil.copy(old_full_path, destination_path_file)
            print(f"新文件名：{destination_path_file}")

            return score_and_sign(destination_path_file)
        else:
            new_old_path =os.path.join(old_path, file)
            file_list = os.listdir(new_old_path)
            for file in file_list:
                return copy_student_file(file, new_old_path, destination_path)

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

teacher = "实验报告针对归并与快速排序实验，目的应包含分治算法的基本思想；实验原理应包括分治、合并和解决；结果分析应对比两种排序。若每项内容均包含并符合算法思想，应得"


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
pd.DataFrame(score_list).to_excel(os.path.join(directory, '实验报告1.xlsx'), index=False)



