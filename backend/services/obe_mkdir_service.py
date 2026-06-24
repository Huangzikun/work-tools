"""OBE 目录生成服务（持久化版）。

v1（已废弃）：临时目录 + 内存 ZIP 流式返回。
v2（当前）：写 obe_task / obe_student 表，目录持久化到 storage/obe/{taskId}/root/，
           返回 {taskId, tree}。

复用：parse_roster / derive_major_name / build_dir_tree 与 v1 一致。
"""

from __future__ import annotations

import io
import os
import re
from dataclasses import dataclass
from typing import List, Optional

import pandas as pd
from werkzeug.datastructures import FileStorage

MAX_UPLOAD_BYTES = 10 * 1024 * 1024  # 10 MB（名单文件）
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


REQUIRED_COLUMNS = {
    "序号": "seq",
    "行政班级": "student_class",
    "学号": "student_id",
    "姓名": "student_name",
}


def _try_decode_html(file_bytes: bytes) -> Optional[str]:
    for enc in ("utf-8", "gb18030", "gbk", "latin-1"):
        try:
            return file_bytes.decode(enc)
        except UnicodeDecodeError:
            continue
    return None


def parse_roster(file_bytes: bytes, filename: str = "") -> List[Student]:
    """解析桂林学院上课点名册（HTML 伪装成 .xls）或标准 .xlsx。"""
    df: Optional[pd.DataFrame] = None
    html_err: Optional[Exception] = None

    text = _try_decode_html(file_bytes)
    if text is not None:
        try:
            tables = pd.read_html(io.StringIO(text))
            if len(tables) >= 2:
                df = tables[1]
        except ValueError as exc:
            html_err = exc

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


def derive_fixed_dir_name(class_name: str, course_name: str, dir_type: str, teacher_name: str) -> str:
    return f"{class_name}《{course_name}》{dir_type}{teacher_name}"


def derive_student_dir_type_name(
    class_name: str, course_name: str, dir_type: str, teacher_name: str, count: int
) -> str:
    return f"{class_name}《{course_name}》{dir_type}{teacher_name}{count}份"


def derive_student_dir_name(student: Student, major_name: str) -> str:
    """单个学生子目录名（与 v6 CLI 保持一致）。"""
    safe_name = os.path.basename(student.student_name)
    return f"{student.student_id}{major_name}{safe_name}"


def build_dir_tree(
    root: str,
    class_name: str,
    course_name: str,
    teacher_name: str,
    fixed_dir_types: List[str],
    student_dir_types: List[str],
    students: List[Student],
) -> None:
    """在 root 下按规则创建完整目录树（与 v6 CLI 输出一致）。"""
    major_name = derive_major_name(class_name)
    count = len(students)

    for dir_type in fixed_dir_types:
        os.makedirs(
            os.path.join(
                root, derive_fixed_dir_name(class_name, course_name, dir_type, teacher_name)
            ),
            exist_ok=True,
        )

    for dir_type in student_dir_types:
        folder = os.path.join(
            root,
            derive_student_dir_type_name(
                class_name, course_name, dir_type, teacher_name, count
            ),
        )
        os.makedirs(folder, exist_ok=True)
        for s in students:
            student_path = os.path.join(folder, derive_student_dir_name(s, major_name))
            os.makedirs(student_path, exist_ok=True)


def create_task(
    user_id: str,
    file_storage: FileStorage,
    class_name: str,
    course_name: str,
    teacher_name: str,
    fixed_dir_types: List[str],
    student_dir_types: List[str],
) -> dict:
    """主流程：校验 → 解析名单 → 创建任务（DB + 磁盘）→ 返回 {taskId, tree, ...}。

    事务边界：先建表+磁盘（成功后 commit），失败则回滚 DB + 清理磁盘。
    """
    from extensions import db
    from models.obe import ObeStudent, ObeTask
    from services.obe_storage import (
        build_tree,
        ensure_task_dirs,
        remove_task_storage,
        root_dir,
    )

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
            f"名单文件大小超过 {MAX_UPLOAD_BYTES // (1024 * 1024)}MB 限制"
        )

    students = parse_roster(file_bytes, filename)
    fixed_dir_types = [t.strip() for t in fixed_dir_types if t and t.strip()]
    student_dir_types = [t.strip() for t in student_dir_types if t and t.strip()]

    # 1. 先 INSERT 拿 task_id（事务内）
    task = ObeTask(
        user_id=user_id,
        class_name=class_name,
        course_name=course_name,
        teacher_name=teacher_name,
        fixed_dir_types=fixed_dir_types,
        student_dir_types=student_dir_types,
        roster_file_name=filename,
        student_count=len(students),
        storage_path="",  # 占位，commit 后更新
        status="created",
    )
    db.session.add(task)
    db.session.flush()  # 拿 task.id 但不 commit
    task_id = task.id

    # 2. 建磁盘目录
    try:
        ensure_task_dirs(task_id)
        build_dir_tree(
            str(root_dir(task_id)),
            class_name,
            course_name,
            teacher_name,
            fixed_dir_types,
            student_dir_types,
            students,
        )
    except Exception as e:
        db.session.rollback()
        remove_task_storage(task_id)
        raise ObeMkdirError(f"创建目录失败：{e}") from e

    # 3. 写 obe_student（每个 student_dir_types × student 一行，experiment_label="default" 占位）
    major_name = derive_major_name(class_name)
    count = len(students)
    try:
        for dir_type in student_dir_types:
            dir_type_name = derive_student_dir_type_name(
                class_name, course_name, dir_type, teacher_name, count
            )
            for s in students:
                student_dir_name = derive_student_dir_name(s, major_name)
                db.session.add(
                    ObeStudent(
                        task_id=task_id,
                        dir_type=dir_type,
                        experiment_label="default",
                        student_id=s.student_id,
                        student_name=s.student_name,
                        student_class=s.student_class,
                        student_dir_name=f"{dir_type_name}/{student_dir_name}",
                        matched=False,
                        grade_status="pending",
                    )
                )
        task.storage_path = str(task_id)
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        remove_task_storage(task_id)
        raise ObeMkdirError(f"写入学生记录失败：{e}") from e

    # 4. 构建目录树
    tree = build_tree(task_id)

    return {
        "taskId": task_id,
        "tree": tree,
        "studentCount": len(students),
        "className": class_name,
        "courseName": course_name,
        "teacherName": teacher_name,
        "fixedDirTypes": fixed_dir_types,
        "studentDirTypes": student_dir_types,
    }
