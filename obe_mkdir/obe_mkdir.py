import os
import pandas as pd
import argparse

'''
执行内容
'''
parser = argparse.ArgumentParser()
parser.add_argument('--class_name', type=str, help='班级名称,如：2022数据科学与大数据技术2班', required=True)
parser.add_argument('--course_name', type=str, help='课程名称', required=True)
parser.add_argument('--teacher_name', type=str, help='教师姓名', required=True)
parser.add_argument('--directory', type=str, help='路径', required=True)
parser.add_argument('--need_mkdir_str', type=str, help='需要创建的目录，如：课程考核、实验实训报告', required=True)

#获取参数
args = parser.parse_args()
class_name = args.class_name
course_name = args.course_name
teacher_name = args.teacher_name
directory = args.directory
need_mkdir_str = args.need_mkdir_str

df = pd.read_excel(os.path.join(directory, '名单.xlsx'), names=['student_id','student_name', "1", "2"])

#名单处理
df = df.apply(lambda x: x.str.replace('\t', ''))
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
    if need_mkdir_str.find(split):
        needMkdirs = need_mkdir_str.split(split)
        break
print(f"创建: {needMkdirs}")
os.makedirs(directory, exist_ok=True)

for need_mkdir in needMkdirs:
    folder_path = os.path.join(directory, f"{class_name}《{course_name}》{need_mkdir}{teacher_name}{len(df)}份")
    os.makedirs(folder_path, exist_ok=True)

    # 遍历 DataFrame
    for index, row in df.iterrows():
        student_id = str(row['student_id'])
        student_name = row['student_name']
        student_path = os.path.join(folder_path, f"{student_id}{class_name}{student_name}")
        os.makedirs(student_path, exist_ok=True)
