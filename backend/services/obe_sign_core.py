"""OBE 批阅核心函数（从 obe_cp_sign/obe_cp_sign_v6.py 抽取并参数化）。

与 v6 的差异：
- 所有签名参数打包到 SignContext，不再依赖模块级 global
- LLMClient 通过参数显式传入
- docPr 计数器用 itertools.count，线程安全
- doc_to_docx / docx_to_pdf 增加 lo_profile_dir 参数，给 LibreOffice 独立 user profile（并发隔离）

v6 CLI 保持不变，本模块是 Web 后端的参数化拷贝。
"""

from __future__ import annotations

import itertools
import os
import re
import subprocess
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
from docx.text.paragraph import Paragraph

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

# 默认 System 提示词（与 v6 一致）
DEFAULT_SYSTEM_PROMPT = (
    "你是桂林学院信息工程学院的一名计算机专任教师，你拥有丰富的教学经验。"
    "你的任务是针对学生提交的实验报告进行批改。你应该理解、使用用户提交的"
    "「教师要求」部分对「学生作答」部分进行批阅。"
    "你可以选择的分数为60,70,80,90和100分，并给出一个50字以内的批阅评语。"
    "生成json格式的内容，包含一个score和一个comment字段。"
)

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
        sign_by_picture(docx_path, docx_path, result["score"], result["comment"], sign_ctx)
        return int(result["score"]), result.get("comment")
    except Exception as e:
        print(f"[fallback] 失败: {e}")
        return 0, None


def score_and_sign_with_file(
    docx_path: str,
    system_prompt: str,
    teacher_prompt: str,
    llm_client,
    sign_ctx: SignContext,
    lo_profile_dir: Optional[str] = None,
) -> tuple[int, Optional[str]]:
    """多模态批阅主流程：DOCX→PDF→上传 Files API→Responses API→签名+打钩。

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
        sign_by_picture(
            docx_path, docx_path, result["score"], result["comment"], sign_ctx, lo_profile_dir
        )
        return int(result["score"]), result.get("comment")

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


def grade_student_docx(
    docx_path: str,
    system_prompt: str,
    teacher_prompt: str,
    llm_client,
    sign_ctx: SignContext,
    lo_profile_dir: Optional[str] = None,
) -> tuple[int, Optional[str], Optional[str]]:
    """原地批改已经在学生目录下的 docx（不再 shutil.copy）。

    流程：
    1. 如果是 .doc 先转 .docx
    2. 调 score_and_sign_with_file（含 PDF 转换、上传、签名、中间页打钩）
    3. 批改后的 docx 直接保存在原地

    返回 (score, comment, graded_file_abs_path)。
    """
    if docx_path.lower().endswith(".doc"):
        converted = doc_to_docx(docx_path, lo_profile_dir)
        if converted:
            # 原地 .doc → .docx 转换并删除原文件
            docx_path = converted

    score, comment = score_and_sign_with_file(
        docx_path,
        system_prompt,
        teacher_prompt,
        llm_client,
        sign_ctx,
        lo_profile_dir,
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
