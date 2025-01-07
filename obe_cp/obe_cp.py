import os
import pandas as pd
import shutil
import argparse
import sys

import hashlib

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


def copy_student_file(file, old_path, destination_path):
    if os.path.isfile(os.path.join(old_path, file)):
        old = os.path.join(old_path, file)

        destination_path_file = file_name_index(file, old_path, destination_path)

        shutil.copy(old, destination_path_file)
        print(f"复制{old}->{destination_path_file}")
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

file_count = 0

student_path_list = os.listdir(new_dir)

for index, row in df.iterrows():
    student_id = str(row['student_id'])
    student_name = row['student_name']

    use_student_path = ''
    for student_path in student_path_list:
        if student_id in student_path:
            use_student_path = os.path.join(new_dir, student_path)
            break

    for file in file_list:
        if student_id in file:
            copy_student_file(file, old_dir, use_student_path)
            print(f"复制{student_id}")
            file_count = file_count + 1
            break
        elif student_name in file:
            copy_student_file(file, old_dir, use_student_path)
            print(f"复制{student_id}")
            file_count = file_count + 1
            break

print(f"复制成功{file_count}")


