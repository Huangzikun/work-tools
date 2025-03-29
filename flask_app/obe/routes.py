import logging

from flask import Blueprint, request
import os
import re
import pandas as pd

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

    try:
        df = pd.read_excel(file, names=['student_id', 'student_name', "1", "2"])
        # 名单处理
        df = df.apply(lambda x: x.str.replace('\t', ''))
        df = df[df['student_id'].str.match(r'^\d+$')]

        # 固定目录
        must_mkdirs = [
            '教学课件',
            '教学教案',
        ]

        directory = os.path.join(BASE_PATH, f"{class_name}《{course_name}》{teacher_name}{df.shape[0]}份汇总")

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

        #复制上传的文件
        file.save(os.path.join(directory, '名单.xlsx'))
        # 获取 directory 下的目录列表
        dir_list = [d for d in os.listdir(directory) if os.path.isdir(os.path.join(directory, d))]

        return {
            'code': 0,
            'msg': '目录创建成功',
                'data': {
                    'directory_list': dir_list,
                    'base_dir': directory,
                    }
                }
    except Exception as e:
        return {'error': str(e)}, 500