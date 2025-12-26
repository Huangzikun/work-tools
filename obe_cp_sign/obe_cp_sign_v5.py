import os
import subprocess
import time
from datetime import datetime, timedelta, timezone
import requests

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
from typing import Optional

# v5版本新增常量
FILE_EXPIRE_DAYS = 5  # 文件过期时间（天）

# 默认的 System 提示词
DEFAULT_SYSTEM_PROMPT = """你是桂林学院信息工程学院的一名计算机专任教师，你拥有丰富的教学经验。你的任务是针对学生提交的实验报告进行批改。你应该理解、使用用户提交的"教师要求"部分对"学生作答"部分进行批阅。你可以选择的分数为60,70,80,90和100分，并给出一个50字以内的批阅评语。生成json格式的内容，包含一个score和一个comment字段。"""


def load_prompt(file_path, default_prompt=None):
    """
    从文件加载提示词，如果文件不存在则使用默认值
    :param file_path: 提示词文件路径
    :param default_prompt: 默认提示词（可选）
    :return: 提示词字符串
    """
    if file_path and os.path.exists(file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read().strip()
    elif default_prompt:
        if file_path:
            print(f"提示: 提示词文件 {file_path} 不存在，使用默认提示词")
        return default_prompt
    else:
        raise FileNotFoundError(f"提示词文件 {file_path} 不存在且无默认值")


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


# ============ v5版本新增函数 ============

def docx_to_pdf(docx_path: str) -> Optional[str]:
    """
    将DOCX文件转换为PDF

    参数:
        docx_path: DOCX文件路径

    返回:
        PDF文件路径,失败返回None
    """
    try:
        docx_path = Path(docx_path).resolve()
        pdf_path = docx_path.with_suffix(".pdf")

        # 如果PDF已存在且较新,直接返回
        if pdf_path.exists():
            pdf_mtime = pdf_path.stat().st_mtime
            docx_mtime = docx_path.stat().st_mtime
            if pdf_mtime > docx_mtime:
                print(f"使用已存在的PDF: {pdf_path}")
                return str(pdf_path)

        print(f"正在转换DOCX到PDF: {docx_path}")

        # 使用LibreOffice转换
        command = [
            'libreoffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', str(docx_path.parent),
            str(docx_path)
        ]

        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=True
        )

        if pdf_path.exists():
            print(f"成功转换PDF: {pdf_path}")
            return str(pdf_path)
        else:
            raise Exception("PDF文件未生成")

    except subprocess.CalledProcessError as e:
        print(f"LibreOffice PDF转换失败: {e.stderr}")
        return None
    except Exception as e:
        print(f"PDF转换异常: {e}")
        return None


def upload_file_via_http(file_path: str, api_key: str, base_url: str = "https://ark.cn-beijing.volces.com/api/v3") -> Optional[str]:
    """
    使用HTTP请求上传文件到豆包Files API

    参数:
        file_path: 文件路径
        api_key: API密钥
        base_url: API基础URL

    返回:
        file_id,失败返回None
    """
    try:
        url = f"{base_url}/files"
        headers = {
            "Authorization": f"Bearer {api_key}"
        }

        # 计算过期时间（使用秒级时间戳，范围：当前时间+86400 到 当前时间+2592000）
        current_timestamp = int(time.time())  # 当前秒级时间戳
        min_expire_at = current_timestamp + 86400  # 最少1天
        max_expire_at = current_timestamp + 2592000  # 最多30天

        # 使用FILE_EXPIRE_DAYS，但限制在有效范围内
        expire_at = current_timestamp + (FILE_EXPIRE_DAYS * 86400)
        if expire_at < min_expire_at:
            expire_at = min_expire_at
        elif expire_at > max_expire_at:
            expire_at = max_expire_at

        expire_datetime = datetime.fromtimestamp(expire_at)
        print(f"当前时间戳(秒): {current_timestamp}")
        print(f"文件过期时间: {expire_datetime} (expire_at={expire_at})")

        # 准备multipart/form-data
        files = {
            'file': open(file_path, 'rb')
        }
        data = {
            'purpose': 'user_data',
            'expire_at': expire_at  # 直接传递整数，不转字符串
        }

        print(f"正在上传文件: {file_path}")
        response = requests.post(url, headers=headers, files=files, data=data)

        if response.status_code == 200:
            result = response.json()
            file_id = result.get('id')
            print(f"文件已上传, file_id={file_id}, status={result.get('status')}")
            return file_id
        else:
            print(f"上传失败, status_code={response.status_code}, response={response.text}")
            return None

    except Exception as e:
        print(f"HTTP上传失败: {e}")
        return None


def wait_for_file_processing_via_http(file_id: str, api_key: str, base_url: str = "https://ark.cn-beijing.volces.com/api/v3") -> bool:
    """
    使用HTTP请求等待文件处理完成

    参数:
        file_id: 文件ID
        api_key: API密钥
        base_url: API基础URL

    返回:
        True表示处理成功, False表示失败
    """
    try:
        url = f"{base_url}/files/{file_id}"
        headers = {
            "Authorization": f"Bearer {api_key}"
        }

        print("等待文件处理完成...")
        max_wait = 120  # 最多等待120秒
        waited = 0
        check_interval = 2

        while waited < max_wait:
            response = requests.get(url, headers=headers)

            if response.status_code == 200:
                result = response.json()
                status = result.get('status')

                if status == 'active':
                    print(f"文件处理完成: {file_id}")
                    return True
                elif status == 'failed':
                    print(f"文件处理失败: {file_id}")
                    return False
                else:
                    print(f"等待文件处理... ({waited}/{max_wait}s), status={status}")

            time.sleep(check_interval)
            waited += check_interval

        print(f"文件处理超时: {file_id}")
        return False

    except Exception as e:
        print(f"等待文件处理失败: {e}")
        return False


def upload_pdf_to_ark(pdf_path: str, api_key: str) -> Optional[str]:
    """
    上传PDF文件到豆包Files API，设置文件过期时间（默认1天）

    参数:
        pdf_path: PDF文件路径
        api_key: API密钥

    返回:
        file_id,失败返回None
    """
    try:
        # 1. 上传文件
        file_id = upload_file_via_http(pdf_path, api_key)
        if not file_id:
            return None

        # 2. 等待文件处理完成
        if wait_for_file_processing_via_http(file_id, api_key):
            return file_id
        else:
            return None

    except Exception as e:
        print(f"上传PDF失败: {e}")
        return None


def score_and_sign_fallback(docx_path: str, system_prompt: str, teacher_prompt: str) -> int:
    """
    回退方案:使用纯文本模式批阅(v4版本逻辑)

    当PDF转换或上传失败时,使用此函数

    参数:
        docx_path: DOCX文件路径
        system_prompt: 系统提示词
        teacher_prompt: 教师要求

    返回:
        分数(0表示失败)
    """
    try:
        print("使用回退方案(纯文本模式)")
        document = Document(docx_path)
        document_text = extract_table_text(document)

        completion = client.chat.completions.create(
            model="doubao-seed-1-6-flash-250828",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"教师要求:{teacher_prompt};学生作答:{document_text}"},
            ],
            response_format={"type": "json_object"},
        )

        print(completion.choices[0].message.content)
        result = json.loads(completion.choices[0].message.content)
        sign_by_picture(docx_path, docx_path, result["score"], result["comment"])
        return result["score"]

    except Exception as e:
        print(f"回退方案也失败: {e}")
        return 0


def score_and_sign_with_file(docx_path: str, system_prompt: str, teacher_prompt: str, ark_client) -> int:
    """
    使用Files API进行多模态批阅并签名（使用SDK同步调用）

    参数:
        docx_path: DOCX文件路径
        system_prompt: 系统提示词
        teacher_prompt: 教师要求
        ark_client: Ark客户端实例

    返回:
        分数(0表示失败)
    """
    pdf_path = None
    file_id = None

    try:
        # 1. 转换DOCX为PDF
        pdf_path = docx_to_pdf(docx_path)
        if not pdf_path:
            print("PDF转换失败,回退到文本模式")
            return score_and_sign_fallback(docx_path, system_prompt, teacher_prompt)

        # 2. 上传PDF到Files API
        file_id = upload_pdf_to_ark(pdf_path, ark_client.api_key)
        if not file_id:
            print("PDF上传失败,回退到文本模式")
            return score_and_sign_fallback(docx_path, system_prompt, teacher_prompt)

        # 3. 调用批阅API（使用HTTP请求调用Responses API）
        # 说明：
        # - 使用 Responses API 进行多模态批阅（而非 Chat API）
        # - 通过 caching={"type": "disabled"} 显式禁用缓存，关闭深度思考相关功能
        # - 使用 input_file 类型引用上传的 PDF 文件
        print("使用多模态API批阅")

        # 使用HTTP请求调用Responses API
        responses_url = "https://ark.cn-beijing.volces.com/api/v3/responses"
        headers = {
            "Authorization": f"Bearer {ark_client.api_key}",
            "Content-Type": "application/json"
        }

        request_body = {
            "model": "doubao-seed-1-6-flash-250828",
            "input": [
                {
                    "role": "system",
                    "content": system_prompt
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_file",
                            "file_id": file_id
                        },
                        {
                            "type": "input_text",
                            "text": f"教师要求:{teacher_prompt}\n请对这份实验报告进行批阅，生成json格式，包含score和comment字段。"
                        }
                    ]
                }
            ],
            "text": {
                "format": {
                    "type": "json_schema",
                    "name": "output",
                    "strict": True,
                    "schema": {
                        "properties": {
                            "score": {
                                "description": "学生的成绩",
                                "type": "integer",
                            },
                            "comment": {
                                "description": "对学生实验报告的评价",
                                "type": "string"
                            }
                        }
                    },
                },
            },
            'thinking': {
                'type': 'disabled'
            }
        }

        print(f"正在调用Responses API: {responses_url}")
        response = requests.post(responses_url, headers=headers, json=request_body)

        if response.status_code != 200:
            raise Exception(f"Responses API调用失败: status_code={response.status_code}, response={response.text}")

        response_data = response.json()

        # 提取响应内容（根据Responses API的响应格式）
        # 响应结构: output -> [message] -> content -> [output_text] -> text
        output_content = ""
        if "output" in response_data:
            for item in response_data["output"]:
                if item.get("type") == "message":
                    for content_item in item.get("content", []):
                        if content_item.get("type") == "output_text":
                            output_content += content_item.get("text", "")

        if not output_content:
            raise Exception(f"无法从响应中提取内容: {response_data}")

        print(f"API返回原始内容: {output_content}")
        result = json.loads(output_content)

        # 4. 签名并保存
        sign_by_picture(docx_path, docx_path, result["score"], result["comment"])

        return result["score"]

    except Exception as e:
        print(f"多模态批阅失败: {e}, 回退到文本模式")
        import traceback
        traceback.print_exc()
        return score_and_sign_fallback(docx_path, system_prompt, teacher_prompt)

    finally:
        # 清理临时PDF文件
        if pdf_path and os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
                print(f"已清理临时PDF: {pdf_path}")
            except Exception as e:
                print(f"清理PDF失败: {e}")


# ============ v5版本新增函数结束 ============


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

# v4版本保留用于兼容
def score_and_sign(destination_path_file, system_prompt, teacher_prompt):
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
                model="doubao-seed-1-6-flash-250828",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"教师要求:{teacher_prompt};学生作答:{document_text}"},
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

def copy_student_file(file, old_path, destination_path, system_prompt, teacher_prompt, ark_client):

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
                    temp_score = max(temp_score, copy_student_file(item, temp_dir, destination_path, system_prompt, teacher_prompt, ark_client))
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

                # v5版本: 使用新的多模态批阅函数
                return score_and_sign_with_file(destination_path_file, system_prompt, teacher_prompt, ark_client)
            except Exception as e:
                print(f"签名失败. file_path={destination_path_file}, error={e}")
                return 0
        else:
            new_old_path =os.path.join(old_path, file)
            file_list = os.listdir(new_old_path)
            print("file_list=", file_list)
            max_score = 0
            for file in file_list:
                max_score = max(max_score, copy_student_file(file, new_old_path, destination_path, system_prompt, teacher_prompt, ark_client))
            return max_score


def load_roster(directory, roster_file):
    """
    加载桂林学院格式学生名单（HTML格式的XLS文件）
    :param directory: 名单所在目录
    :param roster_file: 名单文件名
    :return: DataFrame，包含 student_id 和 student_name 列
    """
    roster_path = os.path.join(directory, roster_file)
    print(f"读取桂林学院格式名单: {roster_path}")

    tables = pd.read_html(roster_path)
    df = tables[1]  # 表格1是学生名单

    # 处理 MultiIndex 列名
    df.columns = df.columns.droplevel(1)

    # 提取学号和姓名列，并重命名
    df = df[['序号', '行政班级', '学号', '姓名']].copy()
    df.columns = ['seq', 'student_class', 'student_id', 'student_name']

    # 名单处理：学号转字符串并筛选有效学生
    df['student_id'] = df['student_id'].astype(str)
    df = df[df['student_id'].str.match(r'^\d+$')]

    print(f"成功加载 {len(df)} 名学生")
    return df


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
parser.add_argument('--roster_file', type=str, help='名单文件名，默认：桂林学院上课点名册.xls',
                    default='桂林学院上课点名册.xls')
parser.add_argument('--system_prompt_file', type=str,
                    help='System提示词文件路径（可选，不提供则使用默认值）',
                    default=None)
parser.add_argument('--teacher_prompt_file', type=str,
                    help='Teacher提示词文件路径（必填）',
                    required=True)

#获取参数
args = parser.parse_args()
old_dir = args.old_dir
new_dir = args.new_dir
directory = args.directory
sign_picture = args.sign_picture
sign = args.sign
sign_date = args.sign_date
roster_file = args.roster_file
system_prompt_file = args.system_prompt_file
teacher_prompt_file = args.teacher_prompt_file

# 请确保您已将 API Key 存储在环境变量 ARK_API_KEY 中
# 从环境变量中获取API Key
api_key = os.environ.get("ARK_API_KEY")

# 初始化Ark客户端
client = Ark(
    # 此为默认路径，您可根据业务所在地域进行配置
    base_url="https://ark.cn-beijing.volces.com/api/v3",
    # 从环境变量中获取您的 API Key。此为默认方式，您可根据需要进行修改
    api_key=api_key,
)

# 加载提示词
system_prompt = load_prompt(system_prompt_file, DEFAULT_SYSTEM_PROMPT)
teacher_prompt = load_prompt(teacher_prompt_file)
print(f"System 提示词已{'从文件加载' if system_prompt_file else '使用默认值'}")
print(f"Teacher 提示词已从文件加载: {teacher_prompt_file}")

# 读取桂林学院格式学生名单
df = load_roster(directory, roster_file)
file_list = os.listdir(old_dir)

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
            score = copy_student_file(file, old_dir, use_student_path, system_prompt, teacher_prompt, client)
            print(f"复制{student_id}")
            file_count = file_count + 1
            break
        elif student_name in file:
            score = copy_student_file(file, old_dir, use_student_path, system_prompt, teacher_prompt, client)
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
