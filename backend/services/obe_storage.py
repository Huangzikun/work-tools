"""OBE 任务文件系统存储层。

布局（相对 OBE_STORAGE_ROOT）：
    {taskId}/
      root/                                 # 实际目录树（学生目录，批改后 docx 写这）
      uploads/{dirType}/                    # 用户上传原始文件，永不修改
      signatures/{sha256}.png               # 签名图按内容去重
      excel/{dirType}_{jobId}_成绩.xlsx      # 每个 job 一份
      jobs/{jobId}/
        logs/batch.log                      # 批改日志
        lo_profile/                         # LibreOffice 独立 user profile
"""

from __future__ import annotations

import hashlib
import os
import shutil
import zipfile
from pathlib import Path
from typing import Optional, Tuple

from flask import current_app


class ObeStorageError(Exception):
    def __init__(self, message: str):
        super().__init__(message)
        self.message = message


# ============ 路径管理 ============


def _storage_root() -> Path:
    """从 current_app.config 读 OBE_STORAGE_ROOT（启动时一次，避免热切）。"""
    root = current_app.config.get("OBE_STORAGE_ROOT")
    if not root:
        raise ObeStorageError("OBE_STORAGE_ROOT 未配置")
    return Path(root)


def task_root(task_id: int) -> Path:
    return _storage_root() / str(task_id)


def task_root_relative(task_id: int) -> str:
    """相对 storage_root 的路径（存数据库）。"""
    return str(task_id)


def root_dir(task_id: int) -> Path:
    """实际目录树根（学生目录在这里）。"""
    return task_root(task_id) / "root"


def uploads_dir(task_id: int, dir_type: str, experiment_label: str | None = None) -> Path:
    """uploads 子目录。传 experiment_label 时返回 uploads/{dirType}/{experiment}/。"""
    base = task_root(task_id) / "uploads" / _safe_segment(dir_type)
    if experiment_label is None:
        return base
    return base / _safe_segment(experiment_label)


def signatures_dir(task_id: int) -> Path:
    return task_root(task_id) / "signatures"


def excel_dir(task_id: int) -> Path:
    return task_root(task_id) / "excel"


def job_dir(task_id: int, job_id: int) -> Path:
    return task_root(task_id) / "jobs" / str(job_id)


def job_log_path(task_id: int, job_id: int) -> Path:
    return job_dir(task_id, job_id) / "logs" / "batch.log"


def job_lo_profile_dir(task_id: int, job_id: int) -> Path:
    return job_dir(task_id, job_id) / "lo_profile"


def _safe_segment(name: str) -> str:
    """把目录段名净化为安全文件名（避免 / 和 ..）。"""
    safe = os.path.basename(name.strip())
    # 去掉 Windows 不允许的字符
    for ch in '<>:"/\\|?*':
        safe = safe.replace(ch, "_")
    return safe or "default"


def ensure_task_dirs(task_id: int) -> None:
    """创建任务全部子目录。mkdir_p 行为。"""
    base = task_root(task_id)
    for sub in ("root", "uploads", "signatures", "excel", "jobs"):
        (base / sub).mkdir(parents=True, exist_ok=True)


# ============ 签名图 sha256 去重 ============


def save_signature(task_id: int, file_bytes: bytes) -> Tuple[str, str]:
    """保存签名图到 signatures/{sha256}.png（按内容去重）。

    返回 (相对任务根的路径, 文件名)。
    """
    digest = hashlib.sha256(file_bytes).hexdigest()
    sig_dir = signatures_dir(task_id)
    sig_dir.mkdir(parents=True, exist_ok=True)
    filename = f"{digest}.png"
    sig_path = sig_dir / filename
    if not sig_path.exists():
        sig_path.write_bytes(file_bytes)
    rel = f"signatures/{filename}"
    return rel, filename


def resolve_signature_path(task_id: int, rel: str) -> Path:
    """把数据库存的相对路径解析为绝对路径（防越权访问）。"""
    abs_path = task_root(task_id) / rel
    abs_path = abs_path.resolve()
    base = task_root(task_id).resolve()
    if abs_path != base and base not in abs_path.parents:
        raise ObeStorageError(f"签名路径越权: {rel}")
    return abs_path


# ============ ZIP 安全解压 ============


def safe_extract_zip(zip_path: Path, dest_dir: Path) -> list[Path]:
    """安全解压 ZIP：拒绝绝对路径和 .. 路径（防 Zip Slip）。

    返回解压出来的所有文件的绝对路径列表。
    """
    dest_dir = dest_dir.resolve()
    dest_dir.mkdir(parents=True, exist_ok=True)
    extracted: list[Path] = []

    with zipfile.ZipFile(zip_path, "r") as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            member = info.filename.replace("\\", "/")
            if member.startswith("/") or member.startswith("\\"):
                raise ObeStorageError(f"非法路径（绝对）: {info.filename}")
            target = (dest_dir / member).resolve()
            if target != dest_dir and dest_dir not in target.parents:
                raise ObeStorageError(f"非法路径（越权）: {info.filename}")
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, open(target, "wb") as dst:
                shutil.copyfileobj(src, dst)
            extracted.append(target)

    return extracted


def extract_nested_zip(outer_zip: Path, dest_dir: Path) -> list[Path]:
    """递归解压嵌套 ZIP：外层 ZIP → 内层 ZIP → 拿到所有最终文件。

    典型场景：用户上传「2024...班-图书管理系统-管理员端-实验报告.zip」，
    里面是每个学生一个 ZIP（202413008771-梁庆胜.zip），最里面是 docx。

    返回所有最深层文件（含 docx 和其他类型）的绝对路径。
    """
    dest_dir = dest_dir.resolve()
    dest_dir.mkdir(parents=True, exist_ok=True)

    # 第一步：解压外层 ZIP
    outer_files = safe_extract_zip(outer_zip, dest_dir)

    # 第二步：对每个内层 ZIP 递归解压
    final_files: list[Path] = []
    inner_zips = [f for f in outer_files if f.suffix.lower() == ".zip"]
    non_zip_files = [f for f in outer_files if f.suffix.lower() != ".zip"]

    final_files.extend(non_zip_files)

    for inner in inner_zips:
        # 内层 ZIP 解压到以自身 stem 命名的子目录
        inner_dest = dest_dir / f"{inner.stem}_unpacked"
        try:
            inner_files = safe_extract_zip(inner, inner_dest)
            # 如果内层还有 zip，继续递归（最多 3 层防无限递归）
            depth = 0
            while any(f.suffix.lower() == ".zip" for f in inner_files) and depth < 3:
                next_round: list[Path] = []
                for f in inner_files:
                    if f.suffix.lower() == ".zip":
                        sub_dest = dest_dir / f"_nested_{depth}_{f.stem}"
                        next_round.extend(safe_extract_zip(f, sub_dest))
                    else:
                        next_round.append(f)
                inner_files = next_round
                depth += 1
            final_files.extend(inner_files)
        except ObeStorageError:
            # 内层解压失败，跳过（外层已经记录了 ZIP 路径，不致命）
            continue

    return final_files


# ============ 目录树构建 ============


def build_tree(task_id: int) -> dict:
    """遍历 root/ 目录构建 NTree 数据结构。

    返回:
        {
          "key": "root",
          "label": "root",
          "type": "dir",
          "children": [
            {"key": "root/实验报告50份", "label": "...", "type": "dir_type",
             "children": [
               {"key": ".../学生目录", "label": "...", "type": "student_dir",
                "children": [{"key": ".../xxx.docx", "label": "xxx.docx", "type": "file"}]}
             ]}
          ]
        }
    """
    root = root_dir(task_id)
    if not root.exists():
        return {"key": "root", "label": "root", "type": "dir", "children": []}
    return _walk_dir(root, "root", "dir")


def _walk_dir(path: Path, key: str, node_type: str) -> dict:
    """递归构建目录树节点。"""
    node: dict = {
        "key": key,
        "label": path.name or "root",
        "type": node_type,
    }

    if path.is_dir():
        children = []
        try:
            entries = sorted(path.iterdir(), key=lambda p: (not p.is_dir(), p.name))
        except OSError:
            entries = []
        for entry in entries:
            child_key = f"{key}/{entry.name}"
            if entry.is_dir():
                children.append(_walk_dir(entry, child_key, "dir"))
            else:
                children.append(
                    {
                        "key": child_key,
                        "label": entry.name,
                        "type": "file",
                        "size": entry.stat().st_size,
                    }
                )
        node["children"] = children
    return node


# ============ 任务清理 ============


def remove_task_storage(task_id: int) -> None:
    """删除任务的整个存储目录。调用方需先确认无 running job。"""
    path = task_root(task_id)
    if path.exists():
        shutil.rmtree(path, ignore_errors=True)


def write_job_log(task_id: int, job_id: int, content: str, append: bool = True) -> None:
    """把批改日志写到 jobs/{jobId}/logs/batch.log。"""
    log_path = job_log_path(task_id, job_id)
    log_path.parent.mkdir(parents=True, exist_ok=True)
    mode = "a" if append else "w"
    with open(log_path, mode, encoding="utf-8") as f:
        f.write(content)
