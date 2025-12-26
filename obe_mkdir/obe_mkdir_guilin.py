import os
import pandas as pd
import argparse
import re

'''
执行内容
'''
parser = argparse.ArgumentParser()
parser.add_argument('--class_name', type=str, help='班级名称,如：2022数据科学与大数据技术2班', required=True)
parser.add_argument('--course_name', type=str, help='课程名称', required=True)
parser.add_argument('--teacher_name', type=str, help='教师姓名', required=True)
parser.add_argument('--directory', type=str, help='路径', required=True)
parser.add_argument('--need_mkdir_str', type=str, help='需要创建的目录，如：课程考核、实验实训报告', required=True)
parser.add_argument('--roster_file', type=str, help='名单文件名，默认：桂林学院上课点名册.xls', default='桂林学院上课点名册.xls')

#获取参数
args = parser.parse_args()
class_name = args.class_name
# 去掉左边的数字作为专业名
major_name = re.sub(r"^\d*", "", class_name)
course_name = args.course_name
teacher_name = args.teacher_name
directory = args.directory
need_mkdir_str = args.need_mkdir_str
roster_file = args.roster_file

# 读取桂林学院点名册格式（HTML格式的XLS文件）
roster_path = os.path.join(directory, roster_file)
tables = pd.read_html(roster_path)
df = tables[1]  # 表格1是学生名单
# 处理 MultiIndex 列名
df.columns = df.columns.droplevel(1)
# 提取学号和姓名列，并重命名
df = df[['序号', '行政班级', '学号', '姓名']].copy()
df.columns = ['seq', 'student_class', 'student_id', 'student_name']

#名单处理：学号转字符串并筛选有效学生
df['student_id'] = df['student_id'].astype(str)
df = df[df['student_id'].str.match(r'^\d+$')]

#固定目录
must_mkdirs = [
    '教学课件',
    '教学教案',
]
for must_mkdir in must_mkdirs:
    folder_path = os.path.join(directory, f"{class_name}《{course_name}》{must_mkdir}{teacher_name}")
    os.makedirs(folder_path, exist_ok=True)

#切分目录
needMkdirs = []
splits = ['、', '/', ',', '，', ';', '；']
for split in splits:
    if split in need_mkdir_str:
        needMkdirs = need_mkdir_str.split(split)
        break
print(f"创建: {needMkdirs}")
os.makedirs(directory, exist_ok=True)

for need_mkdir in needMkdirs:
    folder_path = os.path.join(directory, f"{class_name}《{course_name}》{need_mkdir}{len(df)}份")
    os.makedirs(folder_path, exist_ok=True)

    # 遍历 DataFrame
    for index, row in df.iterrows():
        student_id = str(row['student_id'])
        student_name = row['student_name']
        student_path = os.path.join(folder_path, f"{student_id}{major_name}{student_name}")
        os.makedirs(student_path, exist_ok=True)
