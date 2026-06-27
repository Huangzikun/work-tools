"""教案生成文件系统存储层（参考 obe_storage.py）。

布局（相对 LESSON_PLAN_STORAGE_ROOT）：
    {taskId}/
      syllabus/{原文件名}            # 用户上传的教学大纲，永久保留
      output/{课程名}-教案.docx        # 最终输出
      retry/{timestamp}/             # 重新生成时旧 output 备份
    _templates/通用模板.docx          # 通用首页模板（仓库随代码分发）
"""

from __future__ import annotations

import os
import shutil
from datetime import datetime
from pathlib import Path

from flask import current_app


class LessonPlanStorageError(Exception):
    def __init__(self, message: str):
        super().__init__(message)
        self.message = message


def _storage_root() -> Path:
    root = current_app.config.get("LESSON_PLAN_STORAGE_ROOT")
    if not root:
        raise LessonPlanStorageError("LESSON_PLAN_STORAGE_ROOT 未配置")
    return Path(root)


def _safe_segment(name: str) -> str:
    safe = os.path.basename((name or "").strip())
    for ch in '<>:"/\\|?*':
        safe = safe.replace(ch, "_")
    return safe or "default"


def task_root(task_id: int) -> Path:
    return _storage_root() / str(task_id)


def syllabus_dir(task_id: int) -> Path:
    return task_root(task_id) / "syllabus"


def output_dir(task_id: int) -> Path:
    return task_root(task_id) / "output"


def retry_dir(task_id: int) -> Path:
    return task_root(task_id) / "retry"


def template_path() -> Path:
    """通用模板路径。"""
    p = current_app.config.get("LESSON_PLAN_TEMPLATE_PATH")
    if not p:
        raise LessonPlanStorageError("LESSON_PLAN_TEMPLATE_PATH 未配置")
    return Path(p)


def ensure_task_dirs(task_id: int) -> None:
    base = task_root(task_id)
    for sub in ("syllabus", "output"):
        (base / sub).mkdir(parents=True, exist_ok=True)


def save_syllabus(task_id: int, file_bytes: bytes, original_name: str) -> str:
    """保存上传的教学大纲到 syllabus/{原文件名}，返回保存的绝对路径。"""
    target_dir = syllabus_dir(task_id)
    target_dir.mkdir(parents=True, exist_ok=True)
    safe_name = _safe_segment(original_name) or "syllabus.docx"
    target = target_dir / safe_name
    target.write_bytes(file_bytes)
    return str(target)


def save_syllabus_from_storage(file_storage, task_id: int) -> str:
    """从 Flask FileStorage 保存。"""
    original_name = file_storage.filename or "syllabus.docx"
    file_storage.save(syllabus_dir(task_id) / _safe_segment(original_name))
    return str(syllabus_dir(task_id) / _safe_segment(original_name))


def syllabus_abs_path(task_id: int, syllabus_name: str) -> Path:
    """根据数据库存的 syllabus_name 还原绝对路径。"""
    return syllabus_dir(task_id) / _safe_segment(syllabus_name)


def output_rel_to_abs(task_id: int, rel: str) -> Path:
    """数据库存的相对路径（相对 task_root） → 绝对路径。"""
    abs_path = (task_root(task_id) / rel).resolve()
    base = task_root(task_id).resolve()
    if abs_path != base and base not in abs_path.parents:
        raise LessonPlanStorageError(f"输出路径越权: {rel}")
    return abs_path


def backup_output_for_retry(task_id: int, current_output_rel: str | None) -> str | None:
    """把当前 output 备份到 retry/{timestamp}/。返回备份子目录名（相对路径）。"""
    if not current_output_rel:
        return None
    src = output_rel_to_abs(task_id, current_output_rel)
    if not src.exists():
        return None
    ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    backup_subdir = retry_dir(task_id) / ts
    backup_subdir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, backup_subdir / src.name)
    return f"retry/{ts}/{src.name}"


def remove_task_storage(task_id: int) -> None:
    path = task_root(task_id)
    if path.exists():
        shutil.rmtree(path, ignore_errors=True)
