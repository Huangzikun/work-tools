import os
import pandas as pd
import shutil
import argparse
import sys

def copy_student_file(file, old_path, destination_path):
    if os.path.isfile(os.path.join(old_path, file)):
        old = os.path.join(old_path, file)
        shutil.copy(old, destination_path)
        print(f"复制{old}->{destination_path}")
    else:
        new_old_path =os.path.join(old_path, file)
        file_list = os.listdir(new_old_path)
        for file in file_list:
            copy_student_file(file, new_old_path, destination_path)

'''
执行内容
'''
parser = argparse.ArgumentParser()
parser.add_argument('--old_dir', type=str, help='旧路径', required=True)
parser.add_argument('--new_dir', type=str, help='新路径', required=True)
parser.add_argument('--directory', type=str, help='路径', required=True)

#获取参数
args = parser.parse_args()
old_dir = args.old_dir
new_dir = args.new_dir
directory = args.directory

df = pd.read_excel(os.path.join(directory, '名单.xlsx'), names=['student_id','student_name', "1", "2"])
file_list = os.listdir(old_dir)

df = df.apply(lambda x: x.str.replace('\t', ''))
df = df[df['student_id'].str.match(r'^\d+$')]

# 遍历 DataFrame
for index, row in df.iterrows():
    student_id = str(row['student_id'])
    student_name = row['student_name']
    student_path = os.path.join(new_dir, f"{student_id}{student_name}")
    os.makedirs(student_path, exist_ok=True)



count = 0

for index, row in df.iterrows():
    student_id = str(row['student_id'])
    student_name = row['student_name']
    student_path = os.path.join(new_dir, f"{student_id}{student_name}")

    for file in file_list:
        if student_id in file:
            copy_student_file(file, old_dir, student_path)
            count = count + 1
            break
        elif student_name in file:
            copy_student_file(file, old_dir, student_path)
            print(f"复制{student_id}")
            count = count + 1
            break

print(f"复制成功{count}")


