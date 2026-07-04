"""OBE 批阅核心函数（从 obe_cp_sign/obe_cp_sign_v6.py 抽取并参数化）。

与 v6 的差异：
- 所有签名参数打包到 SignContext，不再依赖模块级 global
- LLMClient 通过参数显式传入
- docPr 计数器用 itertools.count，线程安全
- doc_to_docx / docx_to_pdf 增加 lo_profile_dir 参数，给 LibreOffice 独立 user profile（并发隔离）

v6 CLI 保持不变，本模块是 Web 后端的参数化拷贝。
"""

from __future__ import annotations

import base64
import itertools
import json
import os
import re
import shutil
import subprocess
import tempfile
import time
import zipfile
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional

import requests
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm
from docx.table import Table
from docx.text.paragraph import Paragraph

# 容忍 docx 内部成员的 CRC 不匹配（学生上传的 docx 常有图片 CRC 损坏，LibreOffice 也打不开；
# 之前测试 convert-to docx/pdf/odt 全部 "source file could not be loaded"）。
# 跳过 CRC raise 后 python-docx 能读出文字/表格并保留原始图片数据。对正常文件无副作用。
def _zipextfile_update_crc_lenient(self, newdata):
    if self._expected_crc is None:
        return
    self._running_crc = zipfile.crc32(newdata, self._running_crc)


zipfile.ZipExtFile._update_crc = _zipextfile_update_crc_lenient

# PyMuPDF(fitz) 是否可用：缺失时批改会回退到关键字打钩方案（同一页叠加多个勾）。
# 这里在模块加载时检测并告警，避免悄无声息退化——这正是「每页多个批改痕迹」bug 的根因。
try:
    import fitz  # noqa: F401  PyMuPDF，用于 PDF 反向定位实现「每页一个勾」
    _FITZ_AVAILABLE = True
except ImportError:  # pragma: no cover
    _FITZ_AVAILABLE = False
    import warnings

    warnings.warn(
        "PyMuPDF(fitz) 未安装：批改将回退到关键字打钩方案，会在同一页叠加多个勾。"
        "请 `pip install PyMuPDF` 后重启服务。",
        RuntimeWarning,
        stacklevel=2,
    )

# 文件过期时间（豆包 Files API 要求 1-30 天）
FILE_EXPIRE_DAYS = 5

# OOXML 命名空间
_V6_NS_MAP = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
)
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 首页/末页识别关键字
FIRST_PAGE_KEYWORDS = ("姓名", "学号", "课程名称", "学院", "专业", "指导教师", "开课学期")
LAST_PAGE_KEYWORDS = ("教师评阅", "结果分析与思考", "教师签名")

# 默认 System 提示词（与 v6 一致，仅 legacy PDF 链路使用）
DEFAULT_SYSTEM_PROMPT = (
    "你是桂林学院信息工程学院的一名计算机专任教师，你拥有丰富的教学经验。"
    "你的任务是针对学生提交的实验报告进行批改。你应该理解、使用用户提交的"
    "「教师要求」部分对「学生作答」部分进行批阅。"
    "你可以选择的分数为60,70,80,90和100分，并给出一个50字以内的批阅评语。"
    "生成json格式的内容，包含一个score和一个comment字段。"
)

# ============ Rubric 分维度批改（新链路 v2） ============

# 新链路 System 提示词：执行者角色，完全遵循教师提供的评分标准
GRADING_SYSTEM_PROMPT = (
    "你是实验报告批改的执行者。严格按教师提供的「评分标准」批改学生实验报告，"
    "不得注入任何你自己的评分维度、评分原则、分数档位或评语风格。\n\n"
    "执行规则：\n"
    "1. 严格对照教师提供的评分维度与每维度 max_score 打分；每维度 score 为 0~max_score 整数。\n"
    "2. 每维度必须在 evidence 字段引用学生报告原文片段作为依据，禁止凭印象打分。\n"
    "3. total_score 决策优先级：① 教师在「总体要求」中给出的分数规则（如「完成 X 可得 100 分」"
    "「缺少 Y 扣 Z 分」）最高优先——命中满分条件时必须给 100，不得因细节保守降分；命中扣分条件时"
    "必须扣。② 教师未给明确规则时，total_score 按各维度汇总表现吸附到教师档位（默认 "
    "100/90/80/70/60/50/0）。\n"
    "4. 评语（comment 与每维度 reason）必须与 total_score 严格一致，描述得分或扣分原因："
    "100 分→评语只说明学生如何满足教师满分条件，禁止提及不足；非满分→评语必须明确列出导致扣分的"
    "具体原因，且不得使用「完全达到」「优秀」「完美」「出色」等表示无不足的措辞。"
    "禁止「评语夸赞但分数偏低」或「评语贬低但分数偏高」的脱钩。\n"
    "5. 学生报告以 markdown 为主（表格/代码/标题层级已保留），辅以 PDF 关键页截图。"
    "公式占位 $$...$$ 与图片占位 ![image-N] 表示该处有内容；图片截图在 input_image 中提供。\n"
    "6. 教师标准与报告冲突时以教师标准为准；报告明显残缺（只有标题没正文），对应维度按教师标准扣分，total_score 取最低档或 0。\n"
    "输出严格遵守 JSON schema。"
)

# Rubric 评分输出 schema：豆包 Ark strict 模式要求 object 显式 additionalProperties=False、字段全 required
GRADING_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "dimensions": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "name": {"type": "string", "description": "维度名称"},
                    "score": {"type": "integer", "description": "该维度得分"},
                    "max_score": {"type": "integer", "description": "该维度满分"},
                    "reason": {
                        "type": "string",
                        "description": "该维度评语 30~60 字（风格遵循教师要求）",
                    },
                    "evidence": {"type": "string", "description": "引用学生报告原文片段"},
                },
                "required": ["name", "score", "max_score", "reason", "evidence"],
            },
        },
        "total_score": {
            "type": "integer",
            "description": "总分，必须是教师指定档位之一（默认 100/90/80/70/60/50/0）",
        },
        "comment": {
            "type": "string",
            "description": "总评 80~150 字，必须与 total_score 一致：100 分写得分原因、非满分写扣分原因",
        },
    },
    "required": ["dimensions", "total_score", "comment"],
}

# 默认 5 维度评分量表（教师未自定义时使用）
DEFAULT_RUBRIC = [
    {
        "name": "实验目的与原理",
        "max_score": 20,
        "weight": 0.20,
        "criteria": "是否清晰阐述实验目的、涉及的原理与知识点",
    },
    {
        "name": "实验步骤与过程",
        "max_score": 20,
        "weight": 0.20,
        "criteria": "步骤是否完整、逻辑清晰、可复现",
    },
    {
        "name": "数据与结果",
        "max_score": 25,
        "weight": 0.25,
        "criteria": "数据/截图/运行结果是否真实、完整、与步骤对应",
    },
    {
        "name": "分析与讨论",
        "max_score": 25,
        "weight": 0.25,
        "criteria": "是否对结果深入分析、有独立思考与问题反思",
    },
    {
        "name": "报告规范性",
        "max_score": 10,
        "weight": 0.10,
        "criteria": "格式、图表、语言表达是否规范",
    },
]

# 文本过短阈值：低于此且无图片，疑似扫描件/损坏 docx，回退旧 PDF 链路
MIN_CHARS_FALLBACK = 200
# 视觉采样页上限：PDF 关键页转图最多渲染多少张（避免 mini 模型过载）
MAX_VISION_PAGES = 8
# markdown 主输入截断阈值（字符数），超长尾部截断
MAX_MARKDOWN_CHARS = 30000

# 允许的总分档位（七档，total_score 必须是其中之一；后端 snap 作硬保证）
SCORE_LEVELS = (100, 90, 80, 70, 60, 50, 0)

# docPr id 计数器：itertools.count 的 next() 在 CPython 由 GIL 保护，线程安全
_docpr_counter = itertools.count(1001)


@dataclass
class SignContext:
    """签名相关参数（替代 v6 的全局变量）。"""

    sign_picture_path: str
    sign_date_str: str
    font: str = "楷体"
    color: str = "FF0000"
    sz_half_pt: int = 144  # 72pt


# ============ LibreOffice 文档转换 ============


def _build_libreoffice_cmd(
    template: list[str], lo_profile_dir: Optional[str]
) -> list[str]:
    """给 libreoffice 命令插入独立 user profile 参数（并发隔离）。"""
    cmd = list(template)
    if lo_profile_dir:
        profile_url = f"file://{Path(lo_profile_dir).resolve()}"
        cmd.insert(1, f"-env:UserInstallation={profile_url}")
    return cmd


def doc_to_docx(file_path: str, lo_profile_dir: Optional[str] = None) -> Optional[str]:
    """将单个 .doc 转 .docx，并删除原文件。失败返回 None。"""
    try:
        file_path_obj = Path(file_path).resolve()
        if not file_path_obj.exists():
            raise FileNotFoundError(f"文件不存在: {file_path_obj}")
        if file_path_obj.suffix.lower() != ".doc":
            raise ValueError(f"仅支持 .doc 文件: {file_path_obj}")

        output_dir = file_path_obj.parent
        new_file_path = file_path_obj.with_suffix(".docx")

        command = _build_libreoffice_cmd(
            [
                "libreoffice",
                "--headless",
                "--writer",
                "--nocrashreport",
                "--nodefault",
                "--norestore",
                "--convert-to",
                "docx",
                "--outdir",
                str(output_dir),
                str(file_path_obj),
            ],
            lo_profile_dir,
        )

        subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=True,
        )

        if new_file_path.exists():
            file_path_obj.unlink()
            return str(new_file_path)
        raise Exception("转换失败，未生成 .docx 文件")
    except subprocess.CalledProcessError as e:
        print(f"[doc_to_docx] LibreOffice 失败: {e.stderr}")
        return None
    except Exception as e:
        print(f"[doc_to_docx] 失败: {e}")
        return None


def docx_to_pdf(
    docx_path: str, lo_profile_dir: Optional[str] = None
) -> Optional[str]:
    """DOCX 转 PDF。失败返回 None。已存在的 PDF 若较新则直接复用。"""
    try:
        docx_path_obj = Path(docx_path).resolve()
        pdf_path = docx_path_obj.with_suffix(".pdf")

        if pdf_path.exists():
            if pdf_path.stat().st_mtime > docx_path_obj.stat().st_mtime:
                return str(pdf_path)

        command = _build_libreoffice_cmd(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(docx_path_obj.parent),
                str(docx_path_obj),
            ],
            lo_profile_dir,
        )

        subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=True,
        )

        if pdf_path.exists():
            return str(pdf_path)
        raise Exception("PDF 文件未生成")
    except subprocess.CalledProcessError as e:
        print(f"[docx_to_pdf] LibreOffice 失败: {e.stderr}")
        return None
    except Exception as e:
        print(f"[docx_to_pdf] 失败: {e}")
        return None


# ============ 豆包 Files API 上传 ============


def upload_file_via_http(
    file_path: str, api_key: str, base_url: str = "https://ark.cn-beijing.volces.com/api/v3"
) -> Optional[str]:
    """上传文件到豆包 Files API，返回 file_id。"""
    try:
        url = f"{base_url}/files"
        headers = {"Authorization": f"Bearer {api_key}"}

        current_timestamp = int(time.time())
        expire_at = current_timestamp + FILE_EXPIRE_DAYS * 86400
        # API 要求 1-30 天范围
        expire_at = max(current_timestamp + 86400, min(expire_at, current_timestamp + 2592000))

        with open(file_path, "rb") as fp:
            files = {"file": fp}
            data = {"purpose": "user_data", "expire_at": expire_at}
            response = requests.post(url, headers=headers, files=files, data=data)

        if response.status_code == 200:
            result = response.json()
            file_id = result.get("id")
            print(f"[upload_file_via_http] file_id={file_id}, status={result.get('status')}")
            return file_id
        print(
            f"[upload_file_via_http] 失败 status={response.status_code}, body={response.text}"
        )
        return None
    except Exception as e:
        print(f"[upload_file_via_http] 异常: {e}")
        return None


def wait_for_file_processing_via_http(
    file_id: str, api_key: str, base_url: str = "https://ark.cn-beijing.volces.com/api/v3"
) -> bool:
    """轮询文件处理状态，active 返回 True。"""
    url = f"{base_url}/files/{file_id}"
    headers = {"Authorization": f"Bearer {api_key}"}

    max_wait = 120
    waited = 0
    check_interval = 2

    while waited < max_wait:
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                status = response.json().get("status")
                if status == "active":
                    return True
                if status == "failed":
                    return False
        except Exception as e:
            print(f"[wait_for_file] 查询异常: {e}")
        time.sleep(check_interval)
        waited += check_interval

    return False


def upload_pdf_to_ark(pdf_path: str, api_key: str) -> Optional[str]:
    """上传 PDF 到 Files API 并等待处理完成。"""
    file_id = upload_file_via_http(pdf_path, api_key)
    if not file_id:
        return None
    if wait_for_file_processing_via_http(file_id, api_key):
        return file_id
    return None


# ============ 文本提取 ============


def extract_non_table_text(doc) -> str:
    """提取文档主体中所有非表格段落文本。"""
    result = []
    for element in doc.element.body:
        if element.tag.endswith("}p") or element.tag == f"{{{_W_NS}}}p":
            paragraph = Paragraph(element, doc)
            if paragraph.text.strip():
                result.append(paragraph.text)
    return "\n".join(result)


def extract_table_text(doc) -> str:
    """提取文档中所有表格的文本，按出现顺序去重。"""
    full_text = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    full_text.append(paragraph.text)
    return "\n".join(list(dict.fromkeys(full_text)))


# ============ docx 结构化提取（v2 主输入） ============


def _paragraph_has_image(p_el) -> bool:
    """段落是否含图片（w:drawing 或 v:pic/w:pict）。"""
    for child in p_el.iter():
        tag = child.tag
        if tag.endswith("}drawing") or tag.endswith("}pict"):
            return True
    return False


def _paragraph_has_formula(p_el) -> bool:
    """段落是否含 OMML 公式（m:oMath / m:oMathPara）。"""
    for child in p_el.iter():
        if child.tag.endswith("}oMath") or child.tag.endswith("}oMathPara"):
            return True
    return False


def _paragraph_heading_level(paragraph) -> int:
    """返回标题层级 1-6，非标题返回 0。兼容中文「标题 1」与英文「Heading 1」。"""
    try:
        style_name = paragraph.style.name or ""
    except Exception:
        return 0
    if "标题" in style_name or "Heading" in style_name or "heading" in style_name:
        m = re.search(r"(\d)", style_name)
        if m:
            lvl = int(m.group(1))
            if 1 <= lvl <= 6:
                return lvl
    return 0


def _paragraph_is_list_item(paragraph) -> bool:
    """段落是否是编号/项目符号列表项（pPr 下有 numPr）。"""
    try:
        p_pr = paragraph._p.find(qn("w:pPr"))
        if p_pr is None:
            return False
        return p_pr.find(qn("w:numPr")) is not None
    except Exception:
        return False


def _table_to_markdown(table) -> str:
    """把 docx 表格转成 markdown table 字符串，保留行列结构。

    合并单元格（python-docx 会重复返回同一 cell）按 tc 元素 id 去重；
    单元格内多段落用 <br> 连接；转义 | 防止破坏表格结构。
    """
    rows_md = []
    for row in table.rows:
        cell_texts = []
        seen_tc_ids = set()
        for cell in row.cells:
            tc_id = id(cell._tc)
            if tc_id in seen_tc_ids:
                continue  # 合并单元格重复出现，跳过
            seen_tc_ids.add(tc_id)
            parts = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
            cell_text = "<br>".join(parts)
            cell_text = cell_text.replace("|", "\\|").replace("\n", " ")
            cell_texts.append(cell_text)
        if not cell_texts:
            continue
        rows_md.append("| " + " | ".join(cell_texts) + " |")
    if not rows_md:
        return ""
    col_count = rows_md[0].count("|") - 1
    separator = "| " + " | ".join(["---"] * max(col_count, 1)) + " |"
    rows_md.insert(1, separator)
    return "\n".join(rows_md)


def extract_docx_to_markdown(doc) -> tuple[str, dict]:
    """把 docx 提取成结构化 markdown，保留标题层级/表格/列表/公式占位/图片占位。

    顺序遍历 doc.element.body 一级元素（段落 / 表格），保证文档原始顺序。
    mini 模型读 markdown 文本远比读几十张 PDF 图准确，这是 v2 信息完整性的核心。

    Returns:
        (markdown_text, stats)
        stats = {char_count, table_count, image_count, formula_count, heading_count}
    """
    blocks: list[str] = []
    stats = {
        "char_count": 0,
        "table_count": 0,
        "image_count": 0,
        "formula_count": 0,
        "heading_count": 0,
    }
    image_idx = 0

    for element in doc.element.body:
        tag = element.tag

        if tag.endswith("}p"):
            paragraph = Paragraph(element, doc)
            text = paragraph.text.strip()
            has_image = _paragraph_has_image(element)
            has_formula = _paragraph_has_formula(element)

            if not text and not has_image and not has_formula:
                continue

            lvl = _paragraph_heading_level(paragraph)
            if lvl > 0:
                stats["heading_count"] += 1
                parts = [f"{'#' * lvl} {text}"] if text else [f"{'#' * lvl}"]
                if has_image:
                    image_idx += 1
                    stats["image_count"] += 1
                    parts.append(f"![image-{image_idx}]")
                if has_formula:
                    stats["formula_count"] += 1
                    parts.append("$$公式$$")
                blocks.append("\n".join(parts))
                continue

            prefix = "- " if _paragraph_is_list_item(paragraph) else ""
            line_parts: list[str] = []
            if has_image:
                image_idx += 1
                stats["image_count"] += 1
                line_parts.append(f"![image-{image_idx}]")
            if has_formula:
                stats["formula_count"] += 1
                line_parts.append("$$公式$$")
            if text:
                line_parts.append(text)
            if line_parts:
                blocks.append(prefix + " ".join(line_parts))

        elif tag.endswith("}tbl"):
            stats["table_count"] += 1
            tbl = Table(element, doc)
            md = _table_to_markdown(tbl)
            if md:
                blocks.append(md)

    markdown = "\n\n".join(blocks)
    stats["char_count"] = len(markdown)
    return markdown, stats


# ============ 教师评阅行 ============


def add_teacher_review_row(doc) -> None:
    """在「结果分析与思考」行后插入「教师评阅」行并合并第二列起的单元格。"""
    # 已存在则跳过
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip() == "教师评阅":
                    return

    for table in doc.tables:
        for _ in table.rows:
            row_text = "".join(cell.text.strip() for cell in table.rows[0].cells)
            # 上面这段在 v6 里有点奇怪，按 v6 原逻辑保留
            break
        # 找「结果分析与思考」行
        target_row = None
        for row in table.rows:
            row_text = "".join(cell.text.strip() for cell in row.cells)
            if "结果分析与思考" in row_text:
                target_row = row
                break
        if target_row is None:
            continue
        new_row = table.add_row()
        new_row.cells[0].text = "教师评阅"
        if len(table.columns) > 1:
            new_row.cells[1].merge(new_row.cells[-1])
        return


# ============ XML 工具 ============


def _next_docpr_id() -> int:
    return next(_docpr_counter)


def _xml_escape(text: str) -> str:
    if text is None:
        return ""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def build_review_anchor_xml(
    comment: str,
    score,
    idx: int,
    pos_x: int,
    pos_y: int,
    box_w: int,
    box_h: int,
    font: str,
    color: str,
    sz_half_pt: int,
    text_override: Optional[str] = None,
) -> str:
    """生成浮动文本框 anchor 的 XML 字符串。"""
    if text_override is not None:
        text_content = _xml_escape(str(text_override))
    else:
        safe_comment = _xml_escape(str(comment))
        text_content = f"✓ {safe_comment}  {score}分"
    docpr_id = _next_docpr_id()
    relative_height = 251659264 + idx

    return (
        f"<w:r {_V6_NS_MAP}>"
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
        f"<a:graphic>"
        f'<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f"<wps:wsp>"
        f'<wps:cNvSpPr txBox="1"/>'
        f"<wps:spPr>"
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(box_w)}" cy="{int(box_h)}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f"<a:noFill/>"
        f"<a:ln><a:noFill/></a:ln>"
        f"</wps:spPr>"
        f"<wps:txbx>"
        f"<w:txbxContent>"
        f"<w:p>"
        f'<w:pPr><w:jc w:val="center"/></w:pPr>'
        f"<w:r>"
        f"<w:rPr>"
        f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}"/>'
        f'<w:color w:val="{color}"/>'
        f'<w:sz w:val="{int(sz_half_pt)}"/>'
        f'<w:szCs w:val="{int(sz_half_pt)}"/>'
        f"</w:rPr>"
        f"<w:t>{text_content}</w:t>"
        f"</w:r>"
        f"</w:p>"
        f"</w:txbxContent>"
        f"</wps:txbx>"
        f'<wps:bodyPr rot="0" vert="horz" wrap="square" anchor="ctr"/>'
        f"</wps:wsp>"
        f"</a:graphicData>"
        f"</a:graphic>"
        f"</wp:anchor>"
        f"</w:drawing>"
        f"</w:r>"
    )


# ============ 中间页打钩 ============


def _find_first_paragraph_element(page_elements):
    for el in page_elements:
        if el.tag.endswith("}p") or el.tag == f"{{{_W_NS}}}p":
            return el
    return None


def _find_first_paragraph_in_tables(page_elements):
    for el in page_elements:
        if el.tag.endswith("}tbl") or el.tag == f"{{{_W_NS}}}tbl":
            for tc in el.iter(f"{{{_W_NS}}}tc"):
                for p in tc.iter(f"{{{_W_NS}}}p"):
                    if p is not None:
                        return p
                break
            break
    return None


def split_doc_by_page_breaks(doc):
    """按显式分页符切分文档 body，返回 list[list[Element]]。"""
    pages = [[]]
    body = doc.element.body
    page_break_tag = f"{{{_W_NS}}}br"
    page_break_before_tag = f"{{{_W_NS}}}pageBreakBefore"
    sect_pr_tag = f"{{{_W_NS}}}sectPr"

    for element in body:
        if element.tag == sect_pr_tag:
            continue

        is_paragraph = element.tag == f"{{{_W_NS}}}p"

        if is_paragraph:
            pPr = element.find(f"{{{_W_NS}}}pPr")
            if pPr is not None and pPr.find(page_break_before_tag) is not None:
                if pages[-1]:
                    pages.append([])
                pages[-1].append(element)
                _scan_runs_and_append(element, pages, page_break_tag)
                continue

        pages[-1].append(element)

        if is_paragraph:
            _scan_runs_and_append(element, pages, page_break_tag)

    return pages


def _scan_runs_and_append(paragraph_element, pages, page_break_tag):
    for br in paragraph_element.iter(page_break_tag):
        type_attr = br.get(f"{{{_W_NS}}}type")
        if type_attr == "page":
            pages.append([])


def _row_text(row) -> str:
    return "".join(cell.text for cell in row.cells)


def _first_paragraph_element_of_cell(cell):
    for p in cell._tc.iter(f"{{{_W_NS}}}p"):
        return p
    return None


def insert_review_to_middle_pages_by_keywords(
    doc,
    comment: str,
    score,
    font: str,
    color: str,
    sz_half_pt: int,
    first_page_keywords=FIRST_PAGE_KEYWORDS,
    last_page_keywords=LAST_PAGE_KEYWORDS,
) -> int:
    """关键字方案：在中间页锚定大号 ✓（PyMuPDF 未装时的回退）。"""
    if not doc.sections:
        return 0
    section = doc.sections[0]
    page_w_emu = int(section.page_width)
    page_h_emu = int(section.page_height)
    margin_l = int(section.left_margin) if section.left_margin else 0
    margin_t = int(section.top_margin) if section.top_margin else 0
    margin_r = int(section.right_margin) if section.right_margin else 0
    margin_b = int(section.bottom_margin) if section.bottom_margin else 0

    box_w_emu = int(Cm(3))
    box_h_emu = int(Cm(3))
    content_w = page_w_emu - margin_l - margin_r
    content_h = page_h_emu - margin_t - margin_b
    pos_x = content_w // 2 - box_w_emu // 2
    pos_y = content_h // 2 - box_h_emu // 2

    anchor_targets = []

    for table in doc.tables:
        for row in table.rows:
            row_text = _row_text(row)
            if not row_text.strip():
                continue
            if any(kw in row_text for kw in first_page_keywords):
                continue
            if any(kw in row_text for kw in last_page_keywords):
                continue
            if len(row.cells) > 0:
                p_el = _first_paragraph_element_of_cell(row.cells[0])
                if p_el is not None:
                    anchor_targets.append(("table_row", row_text[:30], p_el))

    for el in doc.element.body:
        if el.tag != f"{{{_W_NS}}}p":
            continue
        text = "".join(t.text or "" for t in el.iter(f"{{{_W_NS}}}t"))
        if not text.strip():
            continue
        if any(kw in text for kw in first_page_keywords):
            continue
        if any(kw in text for kw in last_page_keywords):
            continue
        if text.strip() in ("实验实训报告", "实验报告"):
            continue
        anchor_targets.append(("paragraph", text[:30], el))

    if not anchor_targets:
        return 0

    success = 0
    for idx, (kind, preview, p_el) in enumerate(anchor_targets, start=1):
        xml_str = build_review_anchor_xml(
            comment=comment,
            score=score,
            idx=idx,
            pos_x=pos_x,
            pos_y=pos_y,
            box_w=box_w_emu,
            box_h=box_h_emu,
            font=font,
            color=color,
            sz_half_pt=sz_half_pt,
            text_override="✓",
        )
        try:
            anchor_run = parse_xml(xml_str)
            p_el.append(anchor_run)
            success += 1
            print(f"[keywords] 已在 {kind} 锚定 ✓ (预览='{preview}')")
        except Exception as e:
            print(f"[keywords] 插入失败 ({kind} '{preview}'): {e}")

    return success


def insert_review_to_middle_pages_by_pdf(
    doc,
    pdf_path: str,
    comment: str,
    score,
    font: str,
    color: str,
    sz_half_pt: int,
    first_page_keywords=FIRST_PAGE_KEYWORDS,
    last_page_keywords=LAST_PAGE_KEYWORDS,
) -> int:
    """PDF 反向定位方案：每页只锚定一个 ✓，避免 Word 渲染多个 anchor 错位。"""
    if not _FITZ_AVAILABLE:
        print("[pdf] PyMuPDF 未安装，回退关键字方案（会在同一页叠加多个勾！请 pip install PyMuPDF）")
        return -1
    import fitz  # PyMuPDF

    if not doc.sections:
        return 0
    section = doc.sections[0]
    page_w_emu = int(section.page_width)
    page_h_emu = int(section.page_height)
    margin_l = int(section.left_margin) if section.left_margin else 0
    margin_t = int(section.top_margin) if section.top_margin else 0
    margin_r = int(section.right_margin) if section.right_margin else 0
    margin_b = int(section.bottom_margin) if section.bottom_margin else 0

    box_w_emu = int(Cm(5))
    box_h_emu = int(Cm(5))
    content_w = page_w_emu - margin_l - margin_r
    content_h = page_h_emu - margin_t - margin_b
    pos_x = content_w // 2 - box_w_emu // 2
    pos_y = content_h // 2 - box_h_emu // 2

    pdf_doc = fitz.open(pdf_path)
    page_count = len(pdf_doc)
    page_anchor_texts: list[Optional[str]] = []
    common_headers = ("实验实训报告", "实验报告")
    for page_idx in range(page_count):
        page = pdf_doc[page_idx]
        text = page.get_text()
        clean = re.sub(r"\s+", "", text)
        for h in common_headers:
            clean = clean.replace(h, "")
        candidates = re.findall(r"[一-龥]{4,}", clean)
        page_anchor_texts.append(candidates[0] if candidates else None)
    pdf_doc.close()

    if page_count < 3:
        print(f"[pdf] PDF 仅 {page_count} 页，无中间页")
        return 0

    all_paragraphs = list(doc.element.body.iter(f"{{{_W_NS}}}p"))
    search_start = 0
    paragraph_per_page: list[Optional[object]] = []

    for anchor_text in page_anchor_texts:
        found_p = None
        if anchor_text:
            search_lengths = sorted(
                set([len(anchor_text), max(4, len(anchor_text) // 2), 6, 4]),
                reverse=True,
            )
            search_lengths = [n for n in search_lengths if 4 <= n <= len(anchor_text)]
            for slen in search_lengths:
                search_str = anchor_text[:slen]
                for i in range(search_start, len(all_paragraphs)):
                    p_el = all_paragraphs[i]
                    text = "".join(t.text or "" for t in p_el.iter(f"{{{_W_NS}}}t"))
                    if search_str in text:
                        found_p = p_el
                        search_start = i + 1
                        break
                if found_p is not None:
                    break
        paragraph_per_page.append(found_p)

    success = 0
    middle_pages = list(range(1, page_count - 1))

    for idx, page_idx in enumerate(middle_pages, start=1):
        target_p = paragraph_per_page[page_idx]
        page_anchor_text = page_anchor_texts[page_idx]
        if target_p is None:
            continue
        text = "".join(t.text or "" for t in target_p.iter(f"{{{_W_NS}}}t"))
        if any(kw in text for kw in first_page_keywords):
            continue
        if any(kw in text for kw in last_page_keywords):
            continue
        xml_str = build_review_anchor_xml(
            comment=comment,
            score=score,
            idx=idx,
            pos_x=pos_x,
            pos_y=pos_y,
            box_w=box_w_emu,
            box_h=box_h_emu,
            font=font,
            color=color,
            sz_half_pt=sz_half_pt,
            text_override="✓",
        )
        try:
            anchor_run = parse_xml(xml_str)
            target_p.append(anchor_run)
            success += 1
            preview = page_anchor_text[:30] if page_anchor_text else ""
            print(f"[pdf] 已在 PDF 页 {page_idx + 1} 锚定 ✓ (代表='{preview}')")
        except Exception as e:
            print(f"[pdf] 插入失败 (PDF 页 {page_idx + 1}): {e}")

    return success


# ============ 签名 + 中间页打钩 ============


def sign_by_picture(
    file_path: str,
    save_path: str,
    score,
    review: str,
    sign_ctx: SignContext,
    lo_profile_dir: Optional[str] = None,
) -> int:
    """在「教师评阅」单元格写评语/分数/签名图，并在中间页中心打大号 ✓。"""
    if not file_path.endswith(".docx"):
        print(f"[sign_by_picture] 非 docx，跳过: {file_path}")
        return 0
    doc = Document(file_path)

    add_teacher_review_row(doc)

    if not doc.tables:
        print(f"[sign_by_picture] 无表格，跳过: {file_path}")
        return 0

    for table in doc.tables:
        for row in table.rows:
            for i in range(len(row.cells)):
                if row.cells[i].text.strip() == "教师评阅":
                    if i + 1 >= len(row.cells):
                        continue
                    sep_block = os.linesep * 5
                    row.cells[i + 1].text = (
                        f"{review}成绩：{score}分{sep_block}"
                        "                       教师签名："
                    )
                    paragraph = row.cells[i + 1].paragraphs[-1]
                    paragraph.add_run().add_picture(sign_ctx.sign_picture_path, width=Cm(2))
                    paragraph.add_run(f"   {sign_ctx.sign_date_str}")

                    # 中间页打钩（PDF 反向定位优先，关键字方案回退）
                    try:
                        doc.save(save_path)
                        pdf_path = docx_to_pdf(save_path, lo_profile_dir)
                        n = -1
                        if pdf_path:
                            try:
                                doc2 = Document(save_path)
                                n = insert_review_to_middle_pages_by_pdf(
                                    doc2,
                                    pdf_path,
                                    review,
                                    score,
                                    font=sign_ctx.font,
                                    color=sign_ctx.color,
                                    sz_half_pt=sign_ctx.sz_half_pt,
                                )
                                if n >= 0:
                                    doc2.save(save_path)
                                    print(f"[sign] 打钩完成（PDF 方案），{n} 个 anchor")
                            finally:
                                if os.path.exists(pdf_path):
                                    try:
                                        os.remove(pdf_path)
                                    except Exception:
                                        pass
                        if n < 0:
                            doc3 = Document(save_path)
                            n = insert_review_to_middle_pages_by_keywords(
                                doc3,
                                review,
                                score,
                                font=sign_ctx.font,
                                color=sign_ctx.color,
                                sz_half_pt=sign_ctx.sz_half_pt,
                            )
                            doc3.save(save_path)
                            print(f"[sign] 打钩完成（关键字方案），{n} 个 anchor")
                    except Exception as e:
                        print(f"[sign] 打钩异常: {e}")
                        try:
                            doc.save(save_path)
                        except Exception:
                            pass
                    return 1

    print(f"[sign_by_picture] 未找到'教师评阅'单元格: {file_path}")
    return 0


# ============ Rubric 批改（v2 新链路） ============


def pdf_pages_to_images(
    pdf_path: str, out_dir: str, max_pages: int = MAX_VISION_PAGES
) -> list[str]:
    """用 PyMuPDF 把 PDF 关键页渲染成 PNG 到 out_dir，返回文件路径列表。

    采样策略：≤max_pages 页全转；否则首页 + 末页 + 中间均匀采样到 max_pages 张。
    out_dir 生命周期由调用方管理。matrix(2,2)≈144DPI，mini 模型看得清。
    """
    if not _FITZ_AVAILABLE:
        print("[pdf_pages_to_images] PyMuPDF 未安装，跳过视觉采样")
        return []
    import fitz

    try:
        pdf_doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"[pdf_pages_to_images] 打开 PDF 失败: {e}")
        return []

    page_count = len(pdf_doc)
    if page_count == 0:
        pdf_doc.close()
        return []

    if page_count <= max_pages:
        page_indices = list(range(page_count))
    else:
        page_indices = [0, page_count - 1]
        mid_count = max_pages - 2
        if mid_count > 0:
            step = (page_count - 1) / (mid_count + 1)
            for i in range(1, mid_count + 1):
                idx = int(round(i * step))
                if idx not in page_indices:
                    page_indices.append(idx)
        page_indices = sorted(set(page_indices))[:max_pages]

    out_paths: list[str] = []
    try:
        for page_idx in page_indices:
            try:
                page = pdf_doc[page_idx]
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                out_path = os.path.join(out_dir, f"page_{page_idx + 1}.png")
                pix.save(out_path)
                out_paths.append(out_path)
            except Exception as e:
                print(f"[pdf_pages_to_images] 页 {page_idx + 1} 渲染失败: {e}")
    finally:
        pdf_doc.close()

    return out_paths


def _image_to_data_url(img_path: str) -> str:
    """读图片转 data URL（base64），用于 input_image。"""
    with open(img_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    ext = os.path.splitext(img_path)[1].lower().lstrip(".")
    mime = {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "webp": "image/webp",
    }.get(ext, "image/png")
    return f"data:{mime};base64,{b64}"


def _parse_custom_rubric(text: str) -> list[dict]:
    """从教师文本解析自定义维度。支持「名称（XX分）：说明」/「名称 XX分：说明」/「名称: XX分 说明」。

    匹配失败返回空列表，调用方回退到默认量表。
    """
    result: list[dict] = []
    # 行内同时含「名称 + 数字 + 分 + 冒号 + 说明」
    pattern = re.compile(
        r"([一-龥A-Za-z][一-龥A-Za-z0-9_·\- ]{1,14}?)\s*[（(]?\s*(\d{1,3})\s*分?\s*[）)]?\s*[:：]\s*([^\n；;]+)"
    )
    seen_names = set()
    for m in pattern.finditer(text):
        name = m.group(1).strip()
        try:
            max_score = int(m.group(2))
        except ValueError:
            continue
        criteria = m.group(3).strip()
        if (
            1 <= max_score <= 100
            and name not in seen_names
            and name not in ("评分维度", "评分量表", "rubric", "Rubric")
        ):
            seen_names.add(name)
            result.append(
                {"name": name, "max_score": max_score, "weight": 0.0, "criteria": criteria}
            )
    return result


def build_rubric(
    teacher_prompt: str, dimensions: Optional[list[dict]] = None
) -> list[dict]:
    """构建评分量表。完全由教师控制，优先级：
    1) 显式传入 dimensions（结构化表单）—— 直接用，归一 weight
    2) teacher_prompt 文本能解析出 rubric —— 用解析的（兼容旧调用）
    3) 兜底单维度 [{name:"总体评分", max_score:100, criteria: teacher_prompt}]
    不再注入 DEFAULT_RUBRIC。每项 {name, max_score, weight, criteria}。
    """
    if dimensions:
        clean: list[dict] = []
        total = 0
        for d in dimensions:
            name = (d.get("name") or "").strip()
            try:
                ms = int(d.get("max_score", 0))
            except (TypeError, ValueError):
                continue
            if not name or ms <= 0:
                continue
            clean.append(
                {
                    "name": name,
                    "max_score": ms,
                    "weight": 0.0,
                    "criteria": (d.get("criteria") or "").strip(),
                }
            )
            total += ms
        if clean:
            total = total or 1
            for d in clean:
                d["weight"] = round(d["max_score"] / total, 4)
            return clean

    # 兼容旧文本解析
    if teacher_prompt:
        custom = _parse_custom_rubric(teacher_prompt)
        if len(custom) >= 2:
            total_max = sum(d["max_score"] for d in custom) or 1
            for d in custom:
                d["weight"] = round(d["max_score"] / total_max, 4)
            return custom

    # 极简兜底：让 rubric 流程能跑，标准完全来自自由文本
    return [
        {
            "name": "总体评分",
            "max_score": 100,
            "weight": 1.0,
            "criteria": teacher_prompt.strip() or "按实验报告整体质量评分",
        }
    ]


def parse_score_levels(free_text: str) -> Optional[tuple[int, ...]]:
    """从自由文本解析教师指定档位。

    匹配「分数按 90/80/70/60/50」「档位：90 80 70 60 50」「分档 90、80、70」等。
    解析不到返回 None（调用方用默认七档）。
    """
    if not free_text:
        return None
    if not re.search(r"(分数按|档位|分档|采用.{0,6}档|按.{0,4}档)", free_text):
        return None
    nums = re.findall(r"\b([0-9]{1,3})\b", free_text)
    levels: list[int] = []
    for n in nums:
        v = int(n)
        if 0 <= v <= 100 and v not in levels:
            levels.append(v)
    if len(levels) >= 2:
        return tuple(sorted(levels, reverse=True))
    return None


# ============ AI 生成评分标准（积极评分 + 严谨扣分） ============

RUBRIC_GEN_SYSTEM_PROMPT = (
    "你是教学批改助手。教师会提供实验内容/要求，请据此生成一份结构化批改评分提示词。\n"
    "生成原则：\n"
    "1. 积极评分：学生满足核心实验要求即应给满分（100），不因细枝末节扣分。"
    "把「完成核心任务」明确写成「可得 100 分」的满分条件。\n"
    "2. 严谨扣分：扣分必须有明确、可量化的依据（如「完全未涉及 X 扣 Y 分」），"
    "避免主观吹毛求疵；能用「建议改进」表达的就不扣分。扣分条件用「仅当…才扣」限定。\n"
    "3. 输出 4-6 个评分维度，每维度含 name、max_score（满分，合计 100）、criteria（评分说明）。\n"
    "4. freeText 为总体要求，须包含：「完成 X 可得 100 分」「仅当缺少 Y 才扣 Z 分」"
    "「评语与分值强挂钩，描述得分/扣分原因」等明确规则。\n"
    "输出严格遵守 JSON schema。"
)

RUBRIC_GEN_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "dimensions": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "name": {"type": "string", "description": "维度名"},
                    "max_score": {"type": "integer", "description": "该维度满分（1-100）"},
                    "criteria": {"type": "string", "description": "该维度评分说明"},
                },
                "required": ["name", "max_score", "criteria"],
            },
        },
        "freeText": {"type": "string", "description": "总体要求（满分条件+扣分规则+评语要求）"},
    },
    "required": ["dimensions", "freeText"],
}


def generate_rubric(experiment_content: str, client=None) -> dict:
    """根据教师提供的实验内容，AI 生成「积极评分 + 严谨扣分」的评分标准。

    返回 {dimensions: [{name, max_score, criteria}], freeText: str}。
    """
    if client is None:
        client = get_grading_client()
    user_content = [
        {
            "type": "input_text",
            "text": (
                f"# 教师提供的实验内容/要求\n{experiment_content}\n\n"
                "请据此生成批改评分提示词：4-6 个维度（合计 100 分）+ 总体要求"
                "（积极评分、严谨扣分、评语与分值挂钩）。"
            ),
        }
    ]
    output = client.generate_with_retry(
        system_prompt=RUBRIC_GEN_SYSTEM_PROMPT,
        user_content=user_content,
        json_schema=RUBRIC_GEN_SCHEMA,
        thinking_disabled=True,
    )
    return json.loads(output)


_grading_client = None


def get_grading_client():
    """批改专用 LLMClient（模块级缓存）。

    读 OBE_GRADING_MODEL 环境变量：有则建独立实例（强模型，未来升级即插即用）；
    无则 fallback 到全局 get_default_client()（当前 mini）。不污染全局单例。
    """
    global _grading_client
    if _grading_client is None:
        model = os.environ.get("OBE_GRADING_MODEL")
        if model:
            try:
                from common.llm_client import LLMClient

                _grading_client = LLMClient(model=model)
                print(f"[grading] 批改使用独立模型: {model}")
            except Exception as e:
                print(f"[grading] 模型 {model} 初始化失败，回退默认 client: {e}")
                from common.llm_client import get_default_client

                _grading_client = get_default_client()
        else:
            from common.llm_client import get_default_client

            _grading_client = get_default_client()
    return _grading_client


def _build_grading_user_content(
    markdown: str,
    rubric: list[dict],
    teacher_prompt: str,
    image_paths: list[str],
    score_levels: Optional[tuple[int, ...]] = None,
) -> list[dict]:
    """拼装 Responses API user_content：教师评分维度 + 总体要求 + 学生报告 + 视觉图。"""
    if len(markdown) > MAX_MARKDOWN_CHARS:
        markdown = markdown[:MAX_MARKDOWN_CHARS] + "\n\n...（报告过长，已截断）"

    rubric_text = json.dumps(rubric, ensure_ascii=False, indent=2)
    levels_hint = (
        f"总分必须是以下档位之一：{'/'.join(str(x) for x in score_levels)}。"
        if score_levels
        else "总分必须是 100/90/80/70/60/50/0 七档之一。"
    )
    text_block = (
        "# 教师评分维度（严格遵循）\n"
        f"{rubric_text}\n\n"
        "# 教师总体要求（评语风格 / 档位 / 其他指示，严格遵循）\n"
        f"{teacher_prompt or '（教师未额外说明）'}\n\n"
        "# 评分约束\n"
        f"{levels_hint} 每维度按 max_score 打分；evidence 引用学生报告原文。\n\n"
        "# 学生实验报告（结构化提取，markdown 格式）\n"
        f"{markdown}\n\n"
        "# 任务\n"
        "1. 先判断教师「总体要求」的分数规则是否命中（如「完成多次 X 可得 100」是否满足、「缺少 Y 扣 Z」是否成立），据此定 total_score；\n"
        "2. 各维度按 max_score 打分，evidence 引用学生报告原文；\n"
        "3. comment 必须与 total_score 一致——100 分描述得分原因（满足哪些条件），非满分描述扣分原因（缺什么导致扣分），禁止评语与分数脱钩。\n"
        "输出 dimensions / total_score（教师档位之一）/ comment。"
    )
    content: list[dict] = [{"type": "input_text", "text": text_block}]
    for img_path in image_paths:
        try:
            data_url = _image_to_data_url(img_path)
            content.append({"type": "input_image", "image_url": data_url})
        except Exception as e:
            print(f"[build_content] 图片读取失败 {img_path}: {e}")
    return content


def grade_by_rubric(
    markdown: str,
    rubric: list[dict],
    teacher_prompt: str,
    image_paths: list[str],
    client,
    score_levels: Optional[tuple[int, ...]] = None,
) -> dict:
    """调 LLM 走 rubric 评分，返回 {dimensions, total_score, comment}。

    异常向上抛，由 grade_student_docx_v2 决定是否回退。
    """
    user_content = _build_grading_user_content(
        markdown, rubric, teacher_prompt, image_paths, score_levels
    )
    output = client.generate_with_retry(
        system_prompt=GRADING_SYSTEM_PROMPT,
        user_content=user_content,
        json_schema=GRADING_SCHEMA,
        thinking_disabled=True,
    )
    result = json.loads(output)
    if "dimensions" not in result or "total_score" not in result:
        raise ValueError(f"返回缺少必要字段: {list(result.keys())}")
    return result


# ============ 批改主流程 ============


def score_and_sign_fallback(
    docx_path: str,
    system_prompt: str,
    teacher_prompt: str,
    llm_client,
    sign_ctx: SignContext,
) -> tuple[int, Optional[str]]:
    """回退方案：纯文本批阅（无 PDF 多模态）。

    返回 (score, comment)。score=0 表示失败。
    """
    try:
        print("[fallback] 使用纯文本模式")
        document = Document(docx_path)
        document_text = extract_table_text(document)

        output_text = llm_client.generate(
            system_prompt=system_prompt,
            user_prompt=f"教师要求:{teacher_prompt};学生作答:{document_text}",
            json_output=True,
        )
        print(f"[fallback] LLM 输出: {output_text}")
        import json

        result = json.loads(output_text)
        score = _snap_to_score_level(int(result["score"]))
        comment = result.get("comment")
        sign_by_picture(docx_path, docx_path, score, comment, sign_ctx)
        return score, comment
    except Exception as e:
        print(f"[fallback] 失败: {e}")
        return 0, None


def _legacy_pdf_grading(
    docx_path: str,
    system_prompt: str,
    teacher_prompt: str,
    llm_client,
    sign_ctx: SignContext,
    lo_profile_dir: Optional[str] = None,
) -> tuple[int, Optional[str]]:
    """Legacy PDF 视觉批阅（v2 主链路的 fallback）：DOCX→PDF→上传 Files API→Responses API→签名+打钩。

    整份 PDF 上传让模型自己分页视觉理解，mini 模型对长报告易漏读，仅作 v2 的回退。
    返回 (score, comment)。score=0 表示失败。
    """
    import json

    pdf_path = None
    file_id = None

    try:
        pdf_path = docx_to_pdf(docx_path, lo_profile_dir)
        if not pdf_path:
            print("[score] PDF 转换失败，回退文本模式")
            return score_and_sign_fallback(
                docx_path, system_prompt, teacher_prompt, llm_client, sign_ctx
            )

        file_id = upload_pdf_to_ark(pdf_path, llm_client.api_key)
        if not file_id:
            print("[score] PDF 上传失败，回退文本模式")
            return score_and_sign_fallback(
                docx_path, system_prompt, teacher_prompt, llm_client, sign_ctx
            )

        output_content = llm_client.generate(
            system_prompt=system_prompt,
            user_content=[
                {"type": "input_file", "file_id": file_id},
                {
                    "type": "input_text",
                    "text": (
                        f"教师要求:{teacher_prompt}\n"
                        "请对这份实验报告进行批阅，生成json格式，包含score和comment字段。"
                    ),
                },
            ],
            json_schema={
                "properties": {
                    "score": {"description": "学生的成绩", "type": "integer"},
                    "comment": {
                        "description": "对学生实验报告的评价",
                        "type": "string",
                    },
                }
            },
            thinking_disabled=True,
        )

        print(f"[score] LLM 输出: {output_content}")
        result = json.loads(output_content)
        score = _snap_to_score_level(int(result["score"]))
        comment = result.get("comment")
        sign_by_picture(docx_path, docx_path, score, comment, sign_ctx, lo_profile_dir)
        return score, comment

    except Exception as e:
        print(f"[score] 多模态批阅失败: {e}，回退文本模式")
        import traceback

        traceback.print_exc()
        return score_and_sign_fallback(
            docx_path, system_prompt, teacher_prompt, llm_client, sign_ctx
        )

    finally:
        if pdf_path and os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except Exception as e:
                print(f"[score] 清理 PDF 失败: {e}")


def _snap_to_score_level(raw: int, levels: tuple[int, ...] = SCORE_LEVELS) -> int:
    """把 0-100 任意整数吸附到最近的允许档位。平局取高档。levels 默认七档。"""
    raw = max(0, min(100, int(raw)))
    return min(levels, key=lambda x: (abs(x - raw), -x))


def _try_open_docx(path: str) -> bool:
    """python-docx 能否解析该文件（触发 body 解析）。

    依赖模块顶部的 zipfile CRC 宽容补丁——CRC 不匹配的成员（如损坏图片）不会抛 BadZipFile。
    """
    try:
        Document(path).element.body
        return True
    except Exception:
        return False


def _libreoffice_repair_docx(docx_path: str, lo_profile_dir: Optional[str] = None) -> bool:
    """LibreOffice 重转 docx 标准化（按内容识别格式、容错强）。成功覆盖原文件。

    用于 python-docx 仍读不了的病态 docx（如 relationship target 为 NULL 等非 CRC 问题，
    这类 LibreOffice 能修；CRC 损坏 LibreOffice 也打不开，靠顶部补丁直接读）。
    """
    tmp_dir = tempfile.mkdtemp(prefix="obe_repair_")
    try:
        cmd = _build_libreoffice_cmd(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "docx",
                "--outdir",
                tmp_dir,
                docx_path,
            ],
            lo_profile_dir,
        )
        subprocess.run(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=False
        )
        repaired = os.path.join(tmp_dir, os.path.basename(docx_path))
        if not os.path.exists(repaired) or not _try_open_docx(repaired):
            return False
        shutil.move(repaired, docx_path)
        print("[ensure_docx] LibreOffice 修复成功，已覆盖原文件")
        return True
    except Exception as e:
        print(f"[ensure_docx] LibreOffice 修复异常: {e}")
        return False
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _ensure_docx_readable(docx_path: str, lo_profile_dir: Optional[str] = None) -> bool:
    """确保 docx 可被 python-docx 打开。损坏则修复并覆盖原文件。

    依赖模块顶部的 zipfile CRC 宽容补丁：图片 CRC 损坏时直接读、保留原图（不替换占位）。
    补丁救不了的（relationship NULL 等非 CRC 病态）再走 LibreOffice 重转标准化。
    """
    if _try_open_docx(docx_path):
        return True
    print(f"[ensure_docx] python-docx 打开失败，尝试 LibreOffice 标准化修复")
    return _libreoffice_repair_docx(docx_path, lo_profile_dir)


def grade_student_docx_v2(
    docx_path: str,
    teacher_prompt: str,
    sign_ctx: SignContext,
    lo_profile_dir: Optional[str] = None,
    rubric_dimensions: Optional[list[dict]] = None,
    score_levels: Optional[list[int]] = None,
) -> tuple[int, Optional[str]]:
    """v2 主链路：docx→markdown 结构化提取（主）+ PDF 关键页采样图（辅）+ Rubric 分维度评分。

    三级 fallback：内容过短/疑似损坏 → _legacy_pdf_grading；rubric 调用异常 →
    _legacy_pdf_grading；_legacy_pdf_grading 内部仍会再降到 score_and_sign_fallback。
    返回 (score, comment)。score=0 表示彻底失败。
    """
    pdf_path = None
    tmp_dir = None
    try:
        doc = Document(docx_path)
        markdown, stats = extract_docx_to_markdown(doc)
        print(
            f"[v2] 提取完成: {stats['char_count']}字符 / "
            f"{stats['table_count']}表 / {stats['image_count']}图 / "
            f"{stats['formula_count']}公式 / {stats['heading_count']}标题"
        )

        client = get_grading_client()

        # fallback 判定：内容过短且无图，疑似扫描件/损坏 docx
        if stats["char_count"] < MIN_CHARS_FALLBACK and stats["image_count"] == 0:
            print(f"[v2] 内容过短且无图，回退 legacy PDF 链路")
            return _legacy_pdf_grading(
                docx_path, DEFAULT_SYSTEM_PROMPT, teacher_prompt, client, sign_ctx, lo_profile_dir
            )

        rubric = build_rubric(teacher_prompt, rubric_dimensions)

        # 分数档位：自由文本解析优先，其次教师显式传入，最后默认七档
        levels = parse_score_levels(teacher_prompt)
        if levels is None and score_levels:
            levels = tuple(score_levels)
        if levels is None:
            levels = SCORE_LEVELS

        # 视觉采样：始终尝试转 PDF 渲染关键页（用户要求"始终加 PDF 视觉"）
        pdf_path = docx_to_pdf(docx_path, lo_profile_dir)
        image_paths: list[str] = []
        if pdf_path:
            tmp_dir = tempfile.mkdtemp(prefix="obe_vision_")
            image_paths = pdf_pages_to_images(pdf_path, tmp_dir)
            # 图片已落盘，PDF 本身不再需要，立刻删避免与 sign_by_picture 的 PDF 混淆
            try:
                os.remove(pdf_path)
            except Exception:
                pass
            pdf_path = None

        try:
            result = grade_by_rubric(
                markdown, rubric, teacher_prompt, image_paths, client, score_levels=levels
            )
        except Exception as e:
            print(f"[v2] rubric 评分失败({e})，回退 legacy PDF 链路")
            return _legacy_pdf_grading(
                docx_path, DEFAULT_SYSTEM_PROMPT, teacher_prompt, client, sign_ctx, lo_profile_dir
            )

        score = _snap_to_score_level(int(result.get("total_score", 0)), levels)
        comment = result.get("comment") or ""
        print(f"[v2] 评分完成 total={score} 维度数={len(result.get('dimensions', []))}")

        sign_by_picture(docx_path, docx_path, score, comment, sign_ctx, lo_profile_dir)
        return score, comment

    except Exception as e:
        print(f"[v2] 主链路异常({e})，回退 legacy PDF 链路")
        import traceback

        traceback.print_exc()
        return _legacy_pdf_grading(
            docx_path,
            DEFAULT_SYSTEM_PROMPT,
            teacher_prompt,
            get_grading_client(),
            sign_ctx,
            lo_profile_dir,
        )
    finally:
        if tmp_dir and os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir, ignore_errors=True)
        if pdf_path and os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except Exception:
                pass


def grade_student_docx(
    docx_path: str,
    system_prompt: str,
    teacher_prompt: str,
    llm_client,
    sign_ctx: SignContext,
    lo_profile_dir: Optional[str] = None,
    rubric_dimensions: Optional[list[dict]] = None,
    score_levels: Optional[list[int]] = None,
) -> tuple[int, Optional[str], Optional[str]]:
    """原地批改已经在学生目录下的 docx（不再 shutil.copy）。

    流程：
    1. 如果是 .doc 先转 .docx
    2. 调 grade_student_docx_v2（markdown 提取 + 采样图 + Rubric 评分 + 签名/打钩）
       - system_prompt / llm_client 入参保留以兼容旧调用方，v2 内部使用
         GRADING_SYSTEM_PROMPT 与 get_grading_client()，忽略这两个入参。
       - rubric_dimensions / score_levels 透传给 v2，让教师完全控制评分标准。
    3. 批改后的 docx 直接保存在原地

    返回 (score, comment, graded_file_abs_path)。
    """
    if docx_path.lower().endswith(".doc"):
        converted = doc_to_docx(docx_path, lo_profile_dir)
        if converted:
            # 原地 .doc → .docx 转换并删除原文件
            docx_path = converted

    # 确保是可读的 docx：不相信后缀，损坏则用 LibreOffice 重转修复（覆盖原文件）
    _ensure_docx_readable(docx_path, lo_profile_dir)

    score, comment = grade_student_docx_v2(
        docx_path,
        teacher_prompt,
        sign_ctx,
        lo_profile_dir,
        rubric_dimensions=rubric_dimensions,
        score_levels=score_levels,
    )
    return score, comment, docx_path


# ============ 文件复制 + 批改 ============


def is_archive(file_path: str) -> bool:
    return file_path.lower().endswith(".zip")


def copy_student_file(
    file: str,
    old_path: str,
    destination_path: str,
    system_prompt: str,
    teacher_prompt: str,
    llm_client,
    sign_ctx: SignContext,
    lo_profile_dir: Optional[str] = None,
) -> tuple[int, Optional[str], Optional[str]]:
    """把学生文件复制到对应学生目录（已由上层匹配好），按需转 docx，然后批改签名。

    返回 (score, comment, graded_file_path)。score=0 表示失败。
    """
    import shutil

    old_full_path = os.path.join(old_path, file)

    if is_archive(old_full_path):
        temp_dir = os.path.join(old_path, f"{file}.extracted")
        os.makedirs(temp_dir, exist_ok=True)
        try:
            with zipfile.ZipFile(old_full_path, "r") as zip_ref:
                zip_ref.extractall(temp_dir)
            best = (0, None, None)
            for item in os.listdir(temp_dir):
                score, comment, gfile = copy_student_file(
                    item,
                    temp_dir,
                    destination_path,
                    system_prompt,
                    teacher_prompt,
                    llm_client,
                    sign_ctx,
                    lo_profile_dir,
                )
                if score > best[0]:
                    best = (score, comment, gfile)
            return best
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    if not os.path.isfile(old_full_path):
        return 0, None, None

    destination_file = os.path.join(destination_path, file)
    if not (destination_file.endswith("docx") or destination_file.endswith("doc")):
        return 0, None, None

    try:
        doc = Document(old_full_path)
        document_text = extract_non_table_text(doc)
        if "实验实训报告" not in document_text:
            print(f"[copy] 不是实验实训报告格式，跳过: {file}")
            return 0, None, None
    except Exception as e:
        print(f"[copy] 无法读取 {file}: {e}")
        return 0, None, None

    shutil.copy(old_full_path, destination_file)

    if destination_file.endswith(".doc"):
        converted = doc_to_docx(destination_file, lo_profile_dir)
        if converted:
            destination_file = converted
            # 删除原 .doc（doc_to_docx 已删）

    score, comment = score_and_sign_with_file(
        destination_file,
        system_prompt,
        teacher_prompt,
        llm_client,
        sign_ctx,
        lo_profile_dir,
    )
    return score, comment, destination_file
