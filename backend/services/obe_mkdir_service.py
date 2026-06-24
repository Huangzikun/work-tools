import io
import os
import re
import zipfile
from dataclasses import dataclass
from typing import List, Optional, Tuple

import pandas as pd
from werkzeug.datastructures import FileStorage

MAX_UPLOAD_BYTES = 10 * 1024 * 1024  # 10 MB
ALLOWED_EXTENSIONS = {".xls", ".xlsx", ".html", ".htm"}


class ObeMkdirError(Exception):
    """业务错误，route 层 catch 后转 fail()"""

    def __init__(self, message: str):
        super().__init__(message)
        self.message = message


@dataclass
class Student:
    seq: str
    student_class: str
    student_id: str
    student_name: str


# 名单所需的中文列名 → 内部英文字段
REQUIRED_COLUMNS = {
    "序号": "seq",
    "行政班级": "student_class",
    "学号": "student_id",
    "姓名": "student_name",
}


def _try_decode_html(file_bytes: bytes) -> Optional[str]:
    """按常见中文编码顺序尝试 decode，返回字符串或 None。"""
    for enc in ("utf-8", "gb18030", "gbk", "latin-1"):
        try:
            return file_bytes.decode(enc)
        except UnicodeDecodeError:
            continue
    return None


def parse_roster(file_bytes: bytes, filename: str = "") -> List[Student]:
    """
    解析桂林学院上课点名册（HTML 伪装成 .xls），1:1 复刻 obe_mkdir_guilin.py 逻辑。

    主路径：pd.read_html 取 tables[1]
    兜底：read_html 失败或只有一个表时，尝试 pd.read_excel
    """
    df: Optional[pd.DataFrame] = None
    html_err: Optional[Exception] = None

    # 主路径：先 decode 再 read_html（避免 BytesIO 编码误识别）
    text = _try_decode_html(file_bytes)
    if text is not None:
        try:
            tables = pd.read_html(io.StringIO(text))
            if len(tables) >= 2:
                df = tables[1]
        except ValueError as exc:
            html_err = exc

    # 兜底：read_excel（真 .xlsx/.xls 二进制）
    if df is None:
        try:
            df = pd.read_excel(io.BytesIO(file_bytes))
        except Exception as exc:
            raise ObeMkdirError(
                f"无法解析名单文件（read_html/read_excel 均失败）：{html_err or exc}"
            ) from exc

    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.droplevel(1)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ObeMkdirError(f"名单缺少必要列：{', '.join(missing)}")

    df = df[list(REQUIRED_COLUMNS.keys())].copy()
    df.columns = list(REQUIRED_COLUMNS.values())

    df["student_id"] = df["student_id"].astype(str).str.strip()
    df = df[df["student_id"].str.match(r"^\d+$")]

    if df.empty:
        raise ObeMkdirError("名单解析后无有效学生（学号均为非纯数字）")

    df["student_name"] = df["student_name"].astype(str).str.strip()

    return [
        Student(
            seq=str(row.seq),
            student_class=str(row.student_class),
            student_id=str(row.student_id),
            student_name=str(row.student_name),
        )
        for row in df.itertuples(index=False)
    ]


def derive_major_name(class_name: str) -> str:
    """从班级名提取专业名：去除开头的'数字+级'。"""
    return re.sub(r"^\d+级", "", class_name)


def build_dir_tree(
    root: str,
    class_name: str,
    course_name: str,
    teacher_name: str,
    fixed_dir_types: List[str],
    student_dir_types: List[str],
    students: List[Student],
) -> None:
    """
    在 root 下按 CLI 规则创建完整目录树。

    - fixed_dir_types：每项创建一个空目录（{班级}《{课程}》{类型}{教师}），不带人数、不建学生子目录
    - student_dir_types：每项创建 {班级}《{课程}》{类型}{教师}{人数}份，并在其下为每个学生建 {学号}{专业名}{姓名} 子目录
    """
    major_name = derive_major_name(class_name)
    count = len(students)

    for dir_type in fixed_dir_types:
        path = os.path.join(root, f"{class_name}《{course_name}》{dir_type}{teacher_name}")
        os.makedirs(path, exist_ok=True)

    for dir_type in student_dir_types:
        folder = os.path.join(
            root,
            f"{class_name}《{course_name}》{dir_type}{teacher_name}{count}份",
        )
        os.makedirs(folder, exist_ok=True)
        for s in students:
            # student_name 兜底，防止包含路径分隔符造成穿越
            safe_name = os.path.basename(s.student_name)
            student_path = os.path.join(folder, f"{s.student_id}{major_name}{safe_name}")
            os.makedirs(student_path, exist_ok=True)


def _pack_dir_tree_to_zip(root: str) -> bytes:
    """把 root 下的目录树打包成 ZIP 字节流（所有条目均为空目录/无文件）。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for dirpath, dirnames, _ in os.walk(root):
            for d in dirnames:
                full = os.path.join(dirpath, d)
                arcname = os.path.relpath(full, root).replace(os.sep, "/") + "/"
                zf.writestr(arcname, b"")
    return buf.getvalue()


def generate_zip(
    file_storage: FileStorage,
    class_name: str,
    course_name: str,
    teacher_name: str,
    fixed_dir_types: List[str],
    student_dir_types: List[str],
) -> Tuple[bytes, str]:
    """主流程：校验文件 → 解析名单 → 生成目录树 → 打包 ZIP。"""
    import tempfile

    filename = file_storage.filename or ""
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise ObeMkdirError(
            f"不支持的文件类型：{ext or '无后缀'}，仅支持 {', '.join(sorted(ALLOWED_EXTENSIONS))}"
        )

    if not fixed_dir_types and not student_dir_types:
        raise ObeMkdirError("至少需要指定一个固定目录或考核目录")

    file_bytes = file_storage.read()
    if len(file_bytes) > MAX_UPLOAD_BYTES:
        raise ObeMkdirError(
            f"文件大小超过 {MAX_UPLOAD_BYTES // (1024 * 1024)}MB 限制"
        )

    students = parse_roster(file_bytes, filename)

    with tempfile.TemporaryDirectory(prefix="obe_mkdir_") as tmpdir:
        build_dir_tree(
            tmpdir,
            class_name,
            course_name,
            teacher_name,
            fixed_dir_types,
            student_dir_types,
            students,
        )
        data = _pack_dir_tree_to_zip(tmpdir)

    zip_name = f"{class_name}《{course_name}》OBE目录.zip"
    return data, zip_name
