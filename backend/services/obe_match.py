"""OBE 上传文件 → 学生目录的匹配服务（experiment 维度，按内层 ZIP 名匹配）。

工作流：
1. 上传的外层 ZIP 自动成为一个 experiment（label = ZIP stem）
2. 解外层 ZIP 拿到 N 个内层 ZIP（每个学生一个 ZIP，文件名格式 `{学号}-{姓名}.zip`）
3. 用**内层 ZIP 的文件名**匹配学生（不是最里面的 docx 名，因为 docx 名可能不规范）
4. 匹配成功后，把内层 ZIP 的所有文件（docx + 其他附件）解压到学生目录
   `root/{dirType_folder}/{学号}{专业}{姓名}/`
5. 记录主 docx 路径到 ObeStudent.uploaded_file（批改时用）

兼容场景：
- 嵌套 ZIP（外层 → 内层 ZIP → docx）：用内层 ZIP 名匹配
- 平铺 ZIP（外层 → 直接是 docx）：用 docx 名匹配（fallback）
- 多个独立 docx 上传：用 docx 名匹配
"""

from __future__ import annotations

import re
import uuid
from pathlib import Path
from typing import List

from extensions import db
from models.obe import ObeStudent
from services.obe_storage import (
    ObeStorageError,
    _safe_segment,
    root_dir,
    safe_extract_zip,
    task_root,
    uploads_dir,
)


class ObeMatchError(Exception):
    def __init__(self, message: str):
        super().__init__(message)
        self.message = message


SUPPORTED_DOC_EXTS = {".doc", ".docx"}

_ZIP_NAME_SUFFIXES = [
    "-实验报告(附件)",
    "-实验报告",
    "_实验报告(附件)",
    "_实验报告",
    "（附件）",
    "(附件)",
]


def _experiment_label_from_zip_stem(stem: str) -> str:
    """从 ZIP 文件名 stem 提取实验名。"""
    parts = stem.split("_", 1)
    name = parts[1] if len(parts) == 2 and len(parts[0]) == 8 else stem
    for suffix in _ZIP_NAME_SUFFIXES:
        if name.endswith(suffix):
            name = name[: -len(suffix)]
            break
    name = name.strip()
    return name or "default"


def _student_id_pattern(sid: str) -> re.Pattern:
    return re.compile(rf"(?<!\d){re.escape(sid)}(?!\d)")


def _ensure_experiment_students(
    task_id: int, dir_type: str, experiment_label: str
) -> List[ObeStudent]:
    """确保该 experiment 下有完整学生列表；不存在则从 default 复制。"""
    existing = (
        ObeStudent.query.filter_by(
            task_id=task_id, dir_type=dir_type, experiment_label=experiment_label
        )
        .order_by(ObeStudent.student_id)
        .all()
    )
    if existing:
        return existing

    defaults = (
        ObeStudent.query.filter_by(
            task_id=task_id, dir_type=dir_type, experiment_label="default"
        )
        .order_by(ObeStudent.student_id)
        .all()
    )
    for s in defaults:
        db.session.add(
            ObeStudent(
                task_id=task_id,
                dir_type=dir_type,
                experiment_label=experiment_label,
                student_id=s.student_id,
                student_name=s.student_name,
                student_class=s.student_class,
                student_dir_name=s.student_dir_name,
                matched=False,
                grade_status="pending",
            )
        )
    db.session.flush()
    return list(
        ObeStudent.query.filter_by(
            task_id=task_id, dir_type=dir_type, experiment_label=experiment_label
        ).all()
    )


def _match_one_name_to_student(name_str: str, students, id_patterns):
    """用「学号优先 + 姓名兜底」从一个字符串里匹配学生。返回 (student, reason_or_None)。"""
    id_hits = [s for s, p in id_patterns if p.search(name_str)]
    if len(id_hits) == 1:
        return id_hits[0], None
    if len(id_hits) > 1:
        return None, ("multiple_id_matches", id_hits)

    name_hits = [s for s in students if s.student_name and s.student_name in name_str]
    if len(name_hits) == 1:
        return name_hits[0], None
    if len(name_hits) > 1:
        return None, ("multiple_name_matches", name_hits)

    return None, ("no_match", None)


def _rel_to_task_root(file_path: Path, task_id: int) -> str:
    """相对任务根的路径。"""
    try:
        rel = file_path.resolve().relative_to(task_root(task_id).resolve())
        return str(rel).replace("\\", "/")
    except ValueError:
        return file_path.name


def _save_inner_zip_to_student_dir(
    inner_zip_path: Path, student: ObeStudent, task_id: int
) -> List[Path]:
    """把内层 ZIP 的所有文件解压到学生目录。

    学生目录 = root_dir / student.student_dir_name
    解压后所有文件（docx + 其他附件）都在学生目录下。
    返回解压出来的所有文件绝对路径。
    """
    student_dir = root_dir(task_id) / student.student_dir_name
    student_dir.mkdir(parents=True, exist_ok=True)
    try:
        return safe_extract_zip(inner_zip_path, student_dir)
    except ObeStorageError as e:
        print(f"[match] 内层 ZIP 解压到学生目录失败 {inner_zip_path.name}: {e}")
        return []


def _pick_main_docx(files: List[Path]) -> Path | None:
    """从解压出来的文件里选一个主 docx（用于批改）。"""
    docx_files = [f for f in files if f.suffix.lower() in SUPPORTED_DOC_EXTS]
    if not docx_files:
        return None
    # 优先选文件名含「实验报告」或「报告」的；否则取第一个
    for f in docx_files:
        if "实验报告" in f.name or "报告" in f.name:
            return f
    return docx_files[0]


def _process_inner_zip(
    inner_zip_path: Path,
    students,
    id_patterns,
    task_id: int,
) -> dict:
    """处理单个内层 ZIP：用 ZIP 文件名匹配学生 → 解压到学生目录 → 选主 docx。"""
    zip_stem = inner_zip_path.stem
    student, reason = _match_one_name_to_student(zip_stem, students, id_patterns)

    if student is None:
        reason_type, candidates = reason
        if candidates:
            return {
                "kind": "ambiguous",
                "fileName": inner_zip_path.name,
                "reason": reason_type,
                "candidates": [
                    {"studentId": c.student_id, "studentName": c.student_name}
                    for c in candidates
                ],
            }
        return {"kind": "unmatched", "fileName": inner_zip_path.name}

    # 匹配成功：解压到学生目录
    extracted = _save_inner_zip_to_student_dir(inner_zip_path, student, task_id)
    if not extracted:
        return {
            "kind": "unmatched",
            "fileName": inner_zip_path.name,
            "reason": "extract_failed",
        }

    main_docx = _pick_main_docx(extracted)
    if main_docx is None:
        return {
            "kind": "unmatched",
            "fileName": inner_zip_path.name,
            "reason": "no_docx_inside",
        }

    # 记录到 ObeStudent（覆盖；多次上传同一实验会覆盖）
    rel = _rel_to_task_root(main_docx, task_id)
    student.uploaded_file = rel
    student.matched = True

    return {
        "kind": "matched",
        "studentId": student.student_id,
        "studentName": student.student_name,
        "studentClass": student.student_class,
        "fileName": main_docx.name,
        "filePath": rel,
        "allFiles": [f.name for f in extracted],
    }


def _process_flat_docx(
    docx_path: Path,
    students,
    id_patterns,
    task_id: int,
) -> dict:
    """处理平铺上传的 docx 文件（非嵌套 ZIP 场景）。"""
    stem = docx_path.stem
    student, reason = _match_one_name_to_student(stem, students, id_patterns)

    if student is None:
        reason_type, candidates = reason
        if candidates:
            return {
                "kind": "ambiguous",
                "fileName": docx_path.name,
                "reason": reason_type,
                "candidates": [
                    {"studentId": c.student_id, "studentName": c.student_name}
                    for c in candidates
                ],
            }
        return {"kind": "unmatched", "fileName": docx_path.name}

    rel = _rel_to_task_root(docx_path, task_id)
    student.uploaded_file = rel
    student.matched = True

    return {
        "kind": "matched",
        "studentId": student.student_id,
        "studentName": student.student_name,
        "studentClass": student.student_class,
        "fileName": docx_path.name,
        "filePath": rel,
    }


def _match_one_experiment(
    task_id: int,
    dir_type: str,
    experiment_label: str,
    inner_items: List[tuple[str, Path]],
) -> dict:
    """对单个 experiment 的内层 ZIP / docx 列表做匹配。

    inner_items: [(kind, path), ...]  kind ∈ {"inner_zip", "docx"}
    """
    students = _ensure_experiment_students(task_id, dir_type, experiment_label)
    id_patterns = [(s, _student_id_pattern(s.student_id)) for s in students]

    matched = []
    ambiguous = []
    unmatched = []

    for kind, path in inner_items:
        if kind == "inner_zip":
            result = _process_inner_zip(path, students, id_patterns, task_id)
        else:
            result = _process_flat_docx(path, students, id_patterns, task_id)

        k = result.pop("kind")
        if k == "matched":
            matched.append(result)
        elif k == "ambiguous":
            ambiguous.append(result)
        else:
            unmatched.append(result)

    db.session.commit()
    return {"matched": matched, "ambiguous": ambiguous, "unmatched": unmatched}


# ============ 主入口 ============


def match_uploads(
    task_id: int, dir_type: str, file_storages: list
) -> dict:
    """主入口：处理多个上传文件。

    每个外层 ZIP 自动成为一个 experiment。解外层 ZIP 后：
    - 内层 ZIP：用 ZIP 文件名匹配学生，解压所有文件到学生目录
    - 平铺 docx：用 docx 文件名匹配学生（fallback）
    """
    upload_root = uploads_dir(task_id, dir_type)
    upload_root.mkdir(parents=True, exist_ok=True)
    extracted_root = upload_root / ".extracted"
    extracted_root.mkdir(parents=True, exist_ok=True)

    experiments_data: dict[str, List[tuple[str, Path]]] = {}

    for fs in file_storages:
        original_name = fs.filename or "unknown"
        safe_name = _safe_segment(original_name)
        uuid_prefix = uuid.uuid4().hex[:8]
        target = upload_root / f"{uuid_prefix}_{safe_name}"
        fs.save(target)

        if target.suffix.lower() == ".zip":
            label = _experiment_label_from_zip_stem(target.stem)
            extract_dir = extracted_root / f"{uuid_prefix}_{target.stem}"
            extract_dir.mkdir(parents=True, exist_ok=True)
            try:
                outer_files = safe_extract_zip(target, extract_dir)
            except ObeStorageError as e:
                print(f"[match] 外层 ZIP 解压失败 {target.name}: {e}")
                continue

            # 区分内层 ZIP 和平铺 docx
            for f in outer_files:
                if f.suffix.lower() == ".zip":
                    experiments_data.setdefault(label, []).append(("inner_zip", f))
                elif f.suffix.lower() in SUPPORTED_DOC_EXTS:
                    experiments_data.setdefault(label, []).append(("docx", f))
        elif target.suffix.lower() in SUPPORTED_DOC_EXTS:
            experiments_data.setdefault("独立文件", []).append(("docx", target))

    if not experiments_data:
        return {"experiments": []}

    result_experiments = []
    for label, items in experiments_data.items():
        res = _match_one_experiment(task_id, dir_type, label, items)
        res["experimentLabel"] = label
        res["fileCount"] = len(items)
        result_experiments.append(res)

    return {"experiments": result_experiments}


def resolve_ambiguous(
    task_id: int,
    dir_type: str,
    experiment_label: str,
    resolutions: list,
) -> dict:
    """用户解决 ambiguous：把指定文件绑定到指定学生。

    resolutions: [{filePath(内层 ZIP 路径或 docx 路径), studentId}]
    """
    students = (
        ObeStudent.query.filter_by(
            task_id=task_id, dir_type=dir_type, experiment_label=experiment_label
        ).all()
    )
    sid_to_students = {s.student_id: s for s in students}

    resolved = []
    failed = []

    for r in resolutions:
        student_id = r.get("studentId")
        file_path_str = r.get("filePath") or r.get("fileName")
        s = sid_to_students.get(student_id)
        if not s or not file_path_str:
            failed.append(r)
            continue

        abs_path = (task_root(task_id) / file_path_str).resolve()
        base = task_root(task_id).resolve()
        if abs_path != base and base not in abs_path.parents:
            failed.append(r)
            continue
        if not abs_path.exists():
            failed.append(r)
            continue

        # 如果是内层 ZIP，解压到学生目录
        if abs_path.suffix.lower() == ".zip":
            extracted = _save_inner_zip_to_student_dir(abs_path, s, task_id)
            main_docx = _pick_main_docx(extracted)
            if main_docx is None:
                failed.append(r)
                continue
            rel = _rel_to_task_root(main_docx, task_id)
        else:
            rel = file_path_str

        s.uploaded_file = rel
        s.matched = True
        resolved.append(
            {
                "studentId": s.student_id,
                "studentName": s.student_name,
                "fileName": abs_path.name,
                "filePath": rel,
            }
        )

    db.session.commit()
    return {"resolved": resolved, "failed": failed}
