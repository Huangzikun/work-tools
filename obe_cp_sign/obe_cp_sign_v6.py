# v6: 在 v5 基础上新增按分页符在中间页中心插入 ✓+评语+分数 浮动水印
# v6 完全保留 v5 的 AI 批阅、签名、Excel 汇总、ZIP 解压、DOC→DOCX、PDF 上传等全部能力
# 新增功能：识别 <w:br w:type="page"/> 分页符，跳过首页和末页，在每个中间页中心插入红色楷体批改记录

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
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from typing import Optional

# v5版本新增常量
FILE_EXPIRE_DAYS = 5  # 文件过期时间（天）

# v6版本新增全局变量（在 main 中通过 CLI 参数注入）
review_font = "楷体"          # 批改记录字体
review_color = "FF0000"       # 批改记录颜色（16进制）
review_size_half_pt = 144     # 批改记录字号（半磅，默认 72pt = 144 半磅，大号 ✓）

# docPr id 自增计数器（避免与文档已有 shape 冲突）
_review_docpr_counter = [1000]

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


# ============ v6版本新增函数开始 ============

# OOXML 命名空间（用于 parse_xml）
_V6_NS_MAP = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
)

# WordprocessingML 主命名空间 URI
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _find_first_paragraph_element(page_elements):
    """在页元素列表中找到第一个 <w:p> 元素，找不到返回 None"""
    for el in page_elements:
        if el.tag.endswith('}p') or el.tag == f'{{{_W_NS}}}p':
            return el
    return None


def _find_first_paragraph_in_tables(page_elements):
    """fallback：在页内表格的首单元格首段中返回段落元素"""
    for el in page_elements:
        if el.tag.endswith('}tbl') or el.tag == f'{{{_W_NS}}}tbl':
            # 在 <w:tbl> 内找第一个 <w:tc>，再找其内首个 <w:p>
            for tc in el.iter(f'{{{_W_NS}}}tc'):
                for p in tc.iter(f'{{{_W_NS}}}p'):
                    if p is not None:
                        return p
                break
            break
    return None


def split_doc_by_page_breaks(doc):
    """
    按 <w:br w:type="page"/> 和 <w:pageBreakBefore/> 切分文档，
    返回 list[list[Element]]，每个子列表代表一页的 body 子元素。

    检测两种分页信号：
      1. <w:pPr>/<w:pageBreakBefore/>：段落级前置分页（开新页后再追加该段落）
      2. <w:r>/<w:br w:type="page"/>：run 内显式分页（追加该段后开新页）
    忽略 <w:lastRenderedPageBreak/>（仅渲染提示，不算显式分页符）
    忽略 <w:sectPr>（节属性，不参与分页）
    """
    pages = [[]]
    body = doc.element.body
    page_break_tag = f'{{{_W_NS}}}br'
    page_break_before_tag = f'{{{_W_NS}}}pageBreakBefore'
    sect_pr_tag = f'{{{_W_NS}}}sectPr'

    for element in body:
        # sectPr 不参与分页，也不计入任何页（节属性属于整节元数据）
        if element.tag == sect_pr_tag:
            continue

        is_paragraph = element.tag == f'{{{_W_NS}}}p'

        # 检查 pageBreakBefore（段落属性，开新页再追加）
        if is_paragraph:
            pPr = element.find(f'{{{_W_NS}}}pPr')
            if pPr is not None and pPr.find(page_break_before_tag) is not None:
                # 当前段属于新页
                if pages[-1]:
                    pages.append([])
                # 追加本段
                pages[-1].append(element)
                # 继续扫描本段内部是否还有 run 级分页
                _scan_runs_and_append(element, pages, page_break_tag)
                continue

        # 追加到当前页
        pages[-1].append(element)

        # 检查段落内 run 级分页符
        if is_paragraph:
            _scan_runs_and_append(element, pages, page_break_tag)

    return pages


def _scan_runs_and_append(paragraph_element, pages, page_break_tag):
    """扫描段落内所有 <w:br>，命中 type=page 时开新页（lastRenderedPageBreak 由 tag 过滤忽略）"""
    # 遍历段落内所有 <w:br> 子元素
    for br in paragraph_element.iter(page_break_tag):
        # 检查 type 属性
        type_attr = br.get(f'{{{_W_NS}}}type')
        if type_attr == 'page':
            # 命中显式分页符，开新页
            pages.append([])


def _next_docpr_id():
    """生成全局唯一的 docPr id"""
    _review_docpr_counter[0] += 1
    return _review_docpr_counter[0]


def _xml_escape(text: str) -> str:
    """转义 XML 特殊字符"""
    if text is None:
        return ""
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&apos;")
    )


def build_review_anchor_xml(comment: str, score, idx: int,
                            pos_x: int, pos_y: int,
                            box_w: int, box_h: int,
                            font: str, color: str, sz_half_pt: int,
                            text_override: str = None) -> str:
    """
    生成浮动文本框 anchor 的 XML 字符串。

    使用 <wp:align w:val="center"/> 实现页面居中对齐（替代 posOffset 数值偏移），
    确保在 Word 和 LibreOffice 中均渲染到页面几何中心。

    参数:
        comment: AI 评语（仅在 text_override 为 None 时使用）
        score: 分数（仅在 text_override 为 None 时使用）
        idx: 序号（用于 docPr name）
        pos_x/pos_y: 已废弃，保留参数兼容旧调用
        box_w/box_h: 文本框尺寸（EMU）
        font: 字体名（中文也用此字体，ascii/eastAsia 都设）
        color: 16进制颜色（如 FF0000）
        sz_half_pt: 字号（半磅，14pt = 28）
        text_override: 自定义文本（如纯 "✓"），覆盖默认的 "✓ 评语 分数分"

    返回:
        <w:r><w:drawing>...</w:drawing></w:r> 字符串
    """
    if text_override is not None:
        text_content = _xml_escape(str(text_override))
    else:
        safe_comment = _xml_escape(str(comment))
        text_content = f"✓ {safe_comment}  {score}分"  # ✓ 评语  XX分
    docpr_id = _next_docpr_id()
    relative_height = 251659264 + idx

    xml = (
        f'<w:r {_V6_NS_MAP}>'
        f'<w:drawing>'
        f'<wp:anchor simplePos="0" relativeHeight="{relative_height}" '
        f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="margin"><wp:posOffset>{int(pos_x)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="margin"><wp:posOffset>{int(pos_y)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{int(box_w)}" cy="{int(box_h)}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:wrapNone/>'
        f'<wp:docPr id="{docpr_id}" name="ReviewBox{idx}"/>'
        f'<wp:cNvGraphicFramePr/>'
        f'<a:graphic>'
        f'<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f'<wps:wsp>'
        f'<wps:cNvSpPr txBox="1"/>'
        f'<wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(box_w)}" cy="{int(box_h)}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</wps:spPr>'
        f'<wps:txbx>'
        f'<w:txbxContent>'
        f'<w:p>'
        f'<w:pPr><w:jc w:val="center"/></w:pPr>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}"/>'
        f'<w:color w:val="{color}"/>'
        f'<w:sz w:val="{int(sz_half_pt)}"/>'
        f'<w:szCs w:val="{int(sz_half_pt)}"/>'
        f'</w:rPr>'
        f'<w:t>{text_content}</w:t>'
        f'</w:r>'
        f'</w:p>'
        f'</w:txbxContent>'
        f'</wps:txbx>'
        f'<wps:bodyPr rot="0" vert="horz" wrap="square" anchor="ctr"/>'
        f'</wps:wsp>'
        f'</a:graphicData>'
        f'</a:graphic>'
        f'</wp:anchor>'
        f'</w:drawing>'
        f'</w:r>'
    )
    return xml


def insert_review_to_page_center(doc, pages, comment, score,
                                 font=None, color=None, sz_half_pt=None) -> int:
    """
    [旧版兼容] 基于 split_doc_by_page_breaks 切分结果插入批改记录。
    仅适用于文档含显式分页符的场景。
    """
    if font is None:
        font = review_font
    if color is None:
        color = review_color
    if sz_half_pt is None:
        sz_half_pt = review_size_half_pt

    if len(pages) < 3:
        print(f"文档共 {len(pages)} 页，无中间页可插入批改记录，跳过")
        return 0

    if not doc.sections:
        print("文档无 sections，无法计算坐标，跳过")
        return 0
    section = doc.sections[0]
    page_w_emu = int(section.page_width)
    page_h_emu = int(section.page_height)
    box_w_emu = int(Cm(12))
    box_h_emu = int(Cm(2))
    pos_x = page_w_emu // 2 - box_w_emu // 2
    pos_y = page_h_emu // 2 - box_h_emu // 2

    middle_pages = pages[1:-1]
    success = 0
    for idx, page_elements in enumerate(middle_pages, start=1):
        target_p = _find_first_paragraph_element(page_elements)
        if target_p is None:
            target_p = _find_first_paragraph_in_tables(page_elements)
        if target_p is None:
            print(f"中间页 {idx} 无段落或表格可锚定，跳过")
            continue

        xml_str = build_review_anchor_xml(
            comment=comment, score=score, idx=idx,
            pos_x=pos_x, pos_y=pos_y,
            box_w=box_w_emu, box_h=box_h_emu,
            font=font, color=color, sz_half_pt=sz_half_pt,
        )
        try:
            anchor_run = parse_xml(xml_str)
            target_p.append(anchor_run)
            success += 1
            print(f"已在中间页 {idx} 中心插入批改记录: {comment} {score}分")
        except Exception as e:
            print(f"中间页 {idx} 插入批改记录失败: {e}")

    return success


# 默认首页/末页识别关键字
FIRST_PAGE_KEYWORDS = ("姓名", "学号", "课程名称", "学院", "专业", "指导教师", "开课学期")
LAST_PAGE_KEYWORDS = ("教师评阅", "结果分析与思考", "教师签名")


def _row_text(row):
    """拼接表格行所有单元格的文本"""
    return "".join(cell.text for cell in row.cells)


def _first_paragraph_element_of_cell(cell):
    """获取单元格首个段落元素（lxml element）"""
    for p in cell._tc.iter(f'{{{_W_NS}}}p'):
        return p
    return None


def insert_review_to_middle_pages_by_keywords(
        doc, comment, score,
        font=None, color=None, sz_half_pt=None,
        first_page_keywords=FIRST_PAGE_KEYWORDS,
        last_page_keywords=LAST_PAGE_KEYWORDS) -> int:
    """
    轻量化方案：基于关键字识别首页和末页，在剩余表格行/段落上锚定大号 ✓ 打钩。

    工作原理：
      - 锚定的浮动文本框用 positionH/V relativeFrom='page' 绝对定位到页面中心
      - 浮动元素显示在「锚定段落所在的那一页」
      - 因此只要在某页的任意段落上锚定 anchor，就会在该页中心显示打钩
      - 同一页多个段落都锚定时，多个 anchor 在页面中心重叠（视觉效果仍为一个）

    内容设计：
      - 中间页只显示大号 ✓（评语和分数不在此显示）
      - 评语和分数由 v5 原有的 sign_by_picture 写入末页"教师评阅"单元格

    跳过规则：
      - 含 first_page_keywords 的段落/表格行（识别为首页/封面）
      - 含 last_page_keywords 的表格行（识别为末页/签名页）
      - 空段落/空行

    参数:
        doc: Document 对象
        comment: AI 评语（不在中间页显示，仅日志输出）
        score: 分数（不在中间页显示，仅日志输出）
        font/color/sz_half_pt: 样式参数（None 时用全局变量）
        first_page_keywords: 首页识别关键字元组
        last_page_keywords: 末页识别关键字元组

    返回:
        成功插入的 anchor 数量
    """
    if font is None:
        font = review_font
    if color is None:
        color = review_color
    if sz_half_pt is None:
        sz_half_pt = review_size_half_pt

    # 读取页面尺寸和页边距
    if not doc.sections:
        print("文档无 sections，无法计算坐标，跳过")
        return 0
    section = doc.sections[0]
    page_w_emu = int(section.page_width)
    page_h_emu = int(section.page_height)
    margin_l_emu = int(section.left_margin) if section.left_margin else 0
    margin_t_emu = int(section.top_margin) if section.top_margin else 0
    margin_r_emu = int(section.right_margin) if section.right_margin else 0
    margin_b_emu = int(section.bottom_margin) if section.bottom_margin else 0

    # 大号 ✓ 使用方形文本框（3cm × 3cm），相对 margin 区域居中
    # （relativeFrom="margin" 下 posOffset 是相对内容区域左上角的偏移）
    box_w_emu = int(Cm(3))
    box_h_emu = int(Cm(3))
    content_w = page_w_emu - margin_l_emu - margin_r_emu
    content_h = page_h_emu - margin_t_emu - margin_b_emu
    pos_x = content_w // 2 - box_w_emu // 2
    pos_y = content_h // 2 - box_h_emu // 2

    # 收集可锚定的段落元素（按文档顺序）
    anchor_targets = []

    # 1. 遍历所有表格的行（实验报告主体通常是表格）
    for table in doc.tables:
        for row in table.rows:
            row_text = _row_text(row)
            if not row_text.strip():
                continue
            # 跳过首页行
            if any(kw in row_text for kw in first_page_keywords):
                continue
            # 跳过末页行
            if any(kw in row_text for kw in last_page_keywords):
                continue
            # 锚定到该行第一个单元格的首段
            if len(row.cells) > 0:
                p_el = _first_paragraph_element_of_cell(row.cells[0])
                if p_el is not None:
                    anchor_targets.append(("table_row", row_text[:30], p_el))

    # 2. 顶层段落（非表格内的标题/正文段落）
    #    仅当不含首页/末页关键字时才锚定（用于无表格的纯段落文档）
    for el in doc.element.body:
        if el.tag != f'{{{_W_NS}}}p':
            continue
        text = "".join(t.text or "" for t in el.iter(f'{{{_W_NS}}}t'))
        if not text.strip():
            continue
        if any(kw in text for kw in first_page_keywords):
            continue
        if any(kw in text for kw in last_page_keywords):
            continue
        # 跳过纯"实验实训报告"这种通用标题（避免在首页/页眉位置插入）
        if text.strip() in ("实验实训报告", "实验报告"):
            continue
        anchor_targets.append(("paragraph", text[:30], el))

    if not anchor_targets:
        print("未找到可锚定打钩的段落/表格行，跳过")
        return 0

    # 在每个目标段落上锚定 大号 ✓
    success = 0
    for idx, (kind, preview, p_el) in enumerate(anchor_targets, start=1):
        xml_str = build_review_anchor_xml(
            comment=comment, score=score, idx=idx,
            pos_x=pos_x, pos_y=pos_y,
            box_w=box_w_emu, box_h=box_h_emu,
            font=font, color=color, sz_half_pt=sz_half_pt,
            text_override="✓",
        )
        try:
            anchor_run = parse_xml(xml_str)
            p_el.append(anchor_run)
            success += 1
            print(f"已在{kind}锚定打钩 (预览='{preview}')")
        except Exception as e:
            print(f"插入打钩失败 ({kind} 预览='{preview}'): {e}")

    return success

def insert_review_to_middle_pages_by_pdf(
        doc, pdf_path, comment, score,
        font=None, color=None, sz_half_pt=None,
        first_page_keywords=FIRST_PAGE_KEYWORDS,
        last_page_keywords=LAST_PAGE_KEYWORDS) -> int:
    """
    基于 PDF 反向定位识别每页代表段落，每页只锚定一个 anchor。
    解决 Word 渲染多个 anchor 时错位显示的问题。

    工作原理：
      1. 用 PyMuPDF 打开 PDF，提取每页首个独特文字（去页眉）
      2. 在 docx 中查找这些文字所在的段落
      3. 跳过首页（PDF 第 1 页）和末页（PDF 最后一页）
      4. 在中间页的代表段落上锚定一个大号 ✓

    参数:
        doc: Document 对象
        pdf_path: PDF 文件路径（用于反向定位）
        comment/score: 仅日志用
        font/color/sz_half_pt: 样式参数
        first_page_keywords/last_page_keywords: 备用关键字（用于二次过滤）

    返回:
        成功插入的 anchor 数量
    """
    if font is None:
        font = review_font
    if color is None:
        color = review_color
    if sz_half_pt is None:
        sz_half_pt = review_size_half_pt

    try:
        import fitz  # PyMuPDF
    except ImportError:
        print("PyMuPDF 未安装，回退到关键字方案")
        return -1

    if not doc.sections:
        print("文档无 sections，无法计算坐标，跳过")
        return 0
    section = doc.sections[0]
    page_w_emu = int(section.page_width)
    page_h_emu = int(section.page_height)
    margin_l_emu = int(section.left_margin) if section.left_margin else 0
    margin_t_emu = int(section.top_margin) if section.top_margin else 0
    margin_r_emu = int(section.right_margin) if section.right_margin else 0
    margin_b_emu = int(section.bottom_margin) if section.bottom_margin else 0

    # 大号 ✓ 使用大方形文本框（5cm × 5cm），确保字符完整显示
    box_w_emu = int(Cm(5))
    box_h_emu = int(Cm(5))
    content_w = page_w_emu - margin_l_emu - margin_r_emu
    content_h = page_h_emu - margin_t_emu - margin_b_emu
    pos_x = content_w // 2 - box_w_emu // 2
    pos_y = content_h // 2 - box_h_emu // 2

    # 1. 用 PDF 提取每页首个独特文字（用正则找连续中文字符串）
    import re
    pdf_doc = fitz.open(pdf_path)
    page_count = len(pdf_doc)
    page_anchor_texts = []
    # 通用页眉/页脚（在所有页都出现，不能作为页特征）
    common_headers = ("实验实训报告", "实验报告")
    for page_idx in range(page_count):
        page = pdf_doc[page_idx]
        text = page.get_text()
        # 去掉空白和换行，拼接为连续字符串
        clean = re.sub(r'\s+', '', text)
        # 去掉通用页眉
        for h in common_headers:
            clean = clean.replace(h, '')
        # 找所有连续中文（≥4字符）作为候选
        candidates = re.findall(r'[一-龥]{4,}', clean)
        anchor_text = candidates[0] if candidates else None
        page_anchor_texts.append(anchor_text)
    pdf_doc.close()

    if page_count < 3:
        print(f"PDF 仅 {page_count} 页，无中间页可插入打钩")
        return 0

    # 2. 在 docx 中查找每页代表文字所在的段落
    # 维护一个搜索起点（避免后面页找到前面页的段落）
    all_paragraphs = list(doc.element.body.iter(f'{{{_W_NS}}}p'))
    search_start = 0
    paragraph_per_page = []

    for page_idx, anchor_text in enumerate(page_anchor_texts):
        found_p = None
        if anchor_text:
            # PDF 提取的文字可能拼接了多个单元格，先尝试完整匹配，
            # 失败则逐步缩短搜索串（取前 N 字符）直至找到
            search_lengths = sorted(set(
                [len(anchor_text), max(4, len(anchor_text) // 2), 6, 4]
            ), reverse=True)
            search_lengths = [n for n in search_lengths if 4 <= n <= len(anchor_text)]
            for slen in search_lengths:
                search_str = anchor_text[:slen]
                for i in range(search_start, len(all_paragraphs)):
                    p_el = all_paragraphs[i]
                    text = "".join(t.text or "" for t in p_el.iter(f'{{{_W_NS}}}t'))
                    if search_str in text:
                        found_p = p_el
                        search_start = i + 1
                        break
                if found_p is not None:
                    break
        paragraph_per_page.append(found_p)

    # 3. 跳过首页和末页，在中间页的代表段落锚定
    success = 0
    middle_pages = list(range(1, page_count - 1))  # 跳过 0 和 last

    for idx, page_idx in enumerate(middle_pages, start=1):
        target_p = paragraph_per_page[page_idx]
        page_anchor_text = page_anchor_texts[page_idx]
        if target_p is None:
            print(f"中间页 {idx} (PDF 页 {page_idx + 1}) 未找到代表段落，跳过")
            continue

        # 二次过滤：检查段落是否在首页/末页关键字
        text = "".join(t.text or "" for t in target_p.iter(f'{{{_W_NS}}}t'))
        if any(kw in text for kw in first_page_keywords):
            print(f"中间页 {idx} 段落含首页关键字，跳过")
            continue
        if any(kw in text for kw in last_page_keywords):
            print(f"中间页 {idx} 段落含末页关键字，跳过")
            continue

        xml_str = build_review_anchor_xml(
            comment=comment, score=score, idx=idx,
            pos_x=pos_x, pos_y=pos_y,
            box_w=box_w_emu, box_h=box_h_emu,
            font=font, color=color, sz_half_pt=sz_half_pt,
            text_override="✓",
        )
        try:
            anchor_run = parse_xml(xml_str)
            target_p.append(anchor_run)
            success += 1
            preview = (page_anchor_text[:30] if page_anchor_text else '')
            print(f"已在 PDF 页 {page_idx + 1} 锚定打钩 (代表文字='{preview}')")
        except Exception as e:
            print(f"插入打钩失败 (PDF 页 {page_idx + 1}): {e}")

    return success


# ============ v6版本新增函数结束 ============


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

                        # v6新增：在中间页中心插入 ✓+评语+分数 浮动水印
                        # 优先使用 PDF 反向定位（每页只锚定一次，避免 Word 渲染多个 anchor 错位）
                        # 回退到关键字方案（PyMuPDF 未安装或 PDF 转换失败时）
                        try:
                            # 先保存当前文档（带签名）
                            doc.save(save_path)
                            pdf_path = docx_to_pdf(save_path)
                            n = -1
                            if pdf_path:
                                try:
                                    doc2 = Document(save_path)
                                    n = insert_review_to_middle_pages_by_pdf(
                                        doc2, pdf_path, review, score,
                                        font=review_font,
                                        color=review_color,
                                        sz_half_pt=review_size_half_pt,
                                    )
                                    if n >= 0:
                                        doc2.save(save_path)
                                        print(f"批改记录插入完成（PDF 反向定位方案），共 {n} 个 anchor")
                                finally:
                                    if os.path.exists(pdf_path):
                                        try:
                                            os.remove(pdf_path)
                                        except Exception:
                                            pass
                            # 回退方案
                            if n < 0:
                                doc3 = Document(save_path)
                                n = insert_review_to_middle_pages_by_keywords(
                                    doc3, review, score,
                                    font=review_font,
                                    color=review_color,
                                    sz_half_pt=review_size_half_pt,
                                )
                                doc3.save(save_path)
                                print(f"批改记录插入完成（关键字方案），共 {n} 个 anchor")
                        except Exception as e:
                            print(f"批改记录插入异常: {e}")
                            try:
                                doc.save(save_path)
                            except Exception:
                                pass

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
def _main():
    global old_dir, new_dir, directory, sign_picture, sign, sign_date
    global roster_file, system_prompt_file, teacher_prompt_file
    global review_font, review_color, review_size_half_pt, client

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
    # v6新增 CLI 参数
    parser.add_argument('--review_font', type=str, default='楷体',
                        help='打钩字体，默认楷体')
    parser.add_argument('--review_color', type=str, default='FF0000',
                        help='打钩颜色（16进制），默认 FF0000 红色')
    parser.add_argument('--review_size_pt', type=int, default=72,
                        help='打钩字号（磅），默认 72（大号 ✓，内部转为半磅存储）')

    # 获取参数
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

    # v6新增参数注入全局变量
    review_font = args.review_font
    review_color = args.review_color
    review_size_half_pt = args.review_size_pt * 2  # 磅转半磅

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
    print(f"批改记录样式: font={review_font}, color={review_color}, size={args.review_size_pt}pt")

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


if __name__ == "__main__":
    _main()
