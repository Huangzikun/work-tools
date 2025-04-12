import logging
from flask import Blueprint, request
import os
import re
import pandas as pd
import shutil
import hashlib
import zipfile
import tarfile
import rarfile
import stat

from werkzeug.utils import secure_filename

BASE_PATH = '/tmp/obe_test'

# 创建蓝图对象
obe_bp = Blueprint('obe', __name__)

@obe_bp.route('/obe/mkdir', methods=['POST'])
def obe_mkdir():
    data = request.form
    class_name = data.get('class_name')
    if not class_name:
        return {'code': 400, 'msg': 'class_name is required'}
    # 去掉左边的数字作为专业名
    major_name = re.sub(r"^\d*", "", class_name)
    course_name = data.get('course_name')
    if not course_name:
        return {'code': 400, 'msg': 'course_name is required'}
    teacher_name = data.get('teacher_name')
    if not teacher_name:
        return {'code': 400, 'msg': 'teacher_name is required'}

    need_mkdir_str = data.get('need_mkdir_list')
    if not need_mkdir_str:
        return {'code': 400, 'msg': 'need_mkdir_str is required'}

    # 获取上传的文件
    file = request.files.get('file')
    if not file:
        return {'code': 400, 'msg': '名单文件未上传'}
    if file.filename.split('.')[-1] != 'xlsx':
        return {'code': 400, 'msg': '文件格式必须为 xlsx'}

    # 复制上传的文件
    return_directory = f"{class_name}《{course_name}》{teacher_name}"
    directory = os.path.join(BASE_PATH, return_directory)
    os.makedirs(directory, exist_ok=True)

    file_path = os.path.join(directory, file.filename)
    file.save(file_path)

    try:
        df = pd.read_excel(file_path, names=['student_id', 'student_name', "1", "2"])
        # 名单处理
        df = df.apply(lambda x: x.str.replace('\t', ''))
        df = df[df['student_id'].str.match(r'^\d+$')]

        # 固定目录
        must_mkdirs = [
            '教学课件',
            '教学教案',
        ]


        for must_mkdir in must_mkdirs:
            folder_path = os.path.join(directory, f"{class_name}《{course_name}》{must_mkdir}{teacher_name}")
            os.makedirs(folder_path, exist_ok=True)

        # 切分目录
        needMkdirs = need_mkdir_str.split(',')
        logging.info(f"创建: {needMkdirs}")
        os.makedirs(directory, exist_ok=True)

        for need_mkdir in needMkdirs:
            folder_path = os.path.join(directory, f"{class_name}《{course_name}》{need_mkdir}{teacher_name}{len(df)}份")
            os.makedirs(folder_path, exist_ok=True)

            # 遍历 DataFrame
            for index, row in df.iterrows():
                student_id = str(row['student_id'])
                student_name = row['student_name']
                student_path = os.path.join(folder_path, f"{student_id}{major_name}{student_name}")
                os.makedirs(student_path, exist_ok=True)



        # 获取 directory 下的目录列表
        dir_list = [d for d in os.listdir(directory) if os.path.isdir(os.path.join(directory, d))]

        return {
            'code': 0,
            'msg': '目录创建成功',
                'data': {
                    'directory_list': dir_list,
                    'base_dir': return_directory,
                    }
                }
    except Exception as e:
        return {'error': str(e)}, 500


@obe_bp.route('/obe/filelist', methods=['POST'])
def filelist():
    data = request.form
    path = data.get('path')
    if not path:
        return {'code': 400, 'msg': 'path is required'}

    directory = os.path.join(BASE_PATH, path)
    try:
        dir_list = [d for d in os.listdir(directory) if os.path.isdir(os.path.join(directory, d))]
        # 新增：统计每个目录的文件和子目录数量
        dir_list_value = []
        for dir_name in dir_list:
            dir_path = os.path.join(directory, dir_name)
            filtered_items = [item for item in os.listdir(dir_path) if not item.startswith('.')]
            file_count = len(filtered_items)
            dir_list_value.append({
                'directory': dir_name,
                'file_count': file_count,
            })

        return {
            'code': 0,
            'msg': '目录查询成功',
            'data': {
                'directory_list': dir_list_value,
                'base_dir': path,
            }
        }
    except Exception as e:
        return {'error': str(e)}, 500

@obe_bp.route('/obe/upload', methods=['POST'])
def upload():
    data = request.form
    path = data.get('path')
    if not path:
        return {'code': 400, 'msg': 'path is required'}
    base_dir = path.split('/')[0]
    path = "/".join(path.split('/')[1:])

    if base_dir == '':
        return {'code': 400,'msg': 'path is required'}
    # 获取上传的文件
    file = request.files.get('file')
    if not file:
        return {'code': 400, 'msg': '文件未上传'}


    logging.info(f"上传文件: base_dir: {base_dir}")
    logging.info(f"上传文件: path: {path}")

    base_dir = os.path.join(BASE_PATH, base_dir)
    new_dir = os.path.join(base_dir, path)

    old_dir = os.path.join(base_dir, 'temp')

    if os.path.exists(old_dir):
        return {'code': 400, 'msg': '有上传任务正在进行,请完成后重试'}

    logging.info(f"上传文件: {file.filename}")
    os.makedirs(old_dir, exist_ok=True)

    uploadFilePath = os.path.join(old_dir, file.filename)
    file.save(uploadFilePath)
    logging.info(f"已复制的上传文件: {uploadFilePath}")

    file_ext = file.filename.split('.')[-1]

    if file_ext == 'zip':
        with zipfile.ZipFile(uploadFilePath, 'r') as zip_ref:
            zip_ref.extractall(old_dir)

    save_dir = os.path.join(old_dir, file.filename.split('.')[0])
    # 路径中有数字判断为学生材料文件
    file_count = 0
    if re.search(r'\d', new_dir):
        df = pd.read_excel(os.path.join(base_dir, '名单.xlsx'), names=['student_id', 'student_name', "1", "2"])
        file_list = os.listdir(save_dir)
        logging.info(f"len(file_list): {len(file_list)}")
        df = df.apply(lambda x: x.str.replace('\t', ''))
        df = df[df['student_id'].str.match(r'^\d+$')]
        student_path_list = os.listdir(new_dir)

        for index, row in df.iterrows():
            student_id = str(row['student_id'])
            student_name = row['student_name']

            use_student_path = ''
            for student_path in student_path_list:
                if student_id in student_path:
                    use_student_path = os.path.join(new_dir, student_path)
                    break

            logging.info(f"匹配到的学生路径: {use_student_path}")
            for file in file_list:
                if student_id in file:
                    copy_student_file(file, save_dir, use_student_path)
                    logging.info(f"复制{student_id}")
                    file_count = file_count + 1
                    break
                elif student_name in file:
                    copy_student_file(file, save_dir, use_student_path)
                    logging.info(f"复制{student_id}")
                    file_count = file_count + 1
                    break

        logging.info(f"复制成功{file_count}")
    else:
        dir_list = [d for d in os.listdir(save_dir) if os.path.isdir(os.path.join(save_dir, d))]
        for dir in dir_list:
            shutil.copy(dir, new_dir)

    # 删除临时文件
    if os.path.exists(old_dir):
        shutil.rmtree(old_dir)

    return {'code': 0, 'msg': 'success', 'data': {
        'success_count': file_count,
    }}

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
        logging.info(f"复制{old}->{destination_path_file}")
    else:
        new_old_path =os.path.join(old_path, file)
        file_list = os.listdir(new_old_path)
        for file in file_list:
            copy_student_file(file, new_old_path, destination_path)

