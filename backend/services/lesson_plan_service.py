"""教案生成异步服务（threading + 内存进度缓存，参考 obe_grading.py）。

并发保护：
- 同一任务不能并发跑（重新生成前检查 status 不是 active）
- worker 内串行（一个任务一个线程，不并发调 LLM）
- 进度查询优先读内存缓存
"""

from __future__ import annotations

import threading
from datetime import datetime
from pathlib import Path
from typing import Optional

from flask import current_app

from extensions import db
from models.lesson_plan import LessonPlanTask
from services.lesson_plan.docx_builder import DocxBuilder
from services.lesson_plan.parser import DEFAULT_SYSTEM_PROMPT, SyllabusParser
from services.lesson_plan_storage import (
    LessonPlanStorageError,
    backup_output_for_retry,
    ensure_task_dirs,
    output_dir,
    output_rel_to_abs,
    remove_task_storage,
    save_syllabus,
    syllabus_abs_path,
    task_root,
    template_path,
)


ACTIVE_STATUSES = {"pending", "parsing", "generating", "building"}


class LessonPlanError(Exception):
    def __init__(self, message: str):
        super().__init__(message)
        self.message = message


_WORKERS: dict[int, threading.Thread] = {}
_PROGRESS: dict[int, dict] = {}
_LOCK = threading.Lock()


def _update_progress(task_id: int, **fields) -> None:
    with _LOCK:
        prog = _PROGRESS.setdefault(
            task_id,
            {"status": "pending", "total": 0, "done": 0, "label": None},
        )
        prog.update(fields)
        prog["lastUpdate"] = datetime.utcnow().isoformat()


def _clear_progress(task_id: int) -> None:
    with _LOCK:
        _PROGRESS.pop(task_id, None)


def _make_llm_client():
    from common.llm_client import LLMClient

    return LLMClient(
        api_key=current_app.config.get("ARK_API_KEY", ""),
        base_url=current_app.config.get("ARK_BASE_URL"),
        model=current_app.config.get("ARK_MODEL"),
    )


def create_task(
    user_id: str,
    syllabus_bytes: bytes,
    syllabus_name: str,
    total_lessons: int,
    batch_size: int,
    course_info: dict,
    teacher_info: dict,
    system_prompt: Optional[str] = None,
) -> int:
    """创建任务记录 + 保存大纲 + 启动 worker。返回 task_id。"""
    if total_lessons < 1 or total_lessons > 200:
        raise LessonPlanError("教案数必须在 1-200 之间")
    if batch_size < 1 or batch_size > 5:
        raise LessonPlanError("批量大小必须在 1-5 之间")

    task = LessonPlanTask(
        user_id=user_id,
        syllabus_name=syllabus_name,
        total_lessons=total_lessons,
        batch_size=batch_size,
        course_info=course_info,
        teacher_info=teacher_info,
        system_prompt=system_prompt,
        status="pending",
        progress_total=total_lessons,
        progress_done=0,
        progress_label="等待开始",
    )
    db.session.add(task)
    db.session.flush()

    ensure_task_dirs(task.id)
    save_syllabus(task.id, syllabus_bytes, syllabus_name)

    db.session.commit()

    _update_progress(
        task.id,
        status="pending",
        total=total_lessons,
        done=0,
        label="等待开始",
    )

    _start_worker(task.id)
    return task.id


def _start_worker(task_id: int) -> None:
    app = current_app._get_current_object()

    def worker():
        with app.app_context():
            try:
                _run_generation(task_id)
            except Exception as exc:
                app.logger.exception("lesson_plan task %s failed", task_id)
                _mark_failed(task_id, str(exc))
            finally:
                db.session.remove()
                _clear_progress(task_id)
                _WORKERS.pop(task_id, None)

    t = threading.Thread(target=worker, daemon=True, name=f"lp-{task_id}")
    _WORKERS[task_id] = t
    t.start()


def _run_generation(task_id: int) -> None:
    task = LessonPlanTask.query.get(task_id)
    if task is None:
        raise LessonPlanError(f"task {task_id} 不存在")

    syllabus_path = syllabus_abs_path(task.id, task.syllabus_name)
    if not syllabus_path.exists():
        raise LessonPlanError(f"大纲文件丢失: {syllabus_path}")

    # 状态 → parsing
    task.status = "parsing"
    task.started_at = datetime.utcnow()
    db.session.commit()
    _update_progress(task_id, status="parsing", label="读取教学大纲")

    llm = _make_llm_client()
    sys_prompt = task.system_prompt or DEFAULT_SYSTEM_PROMPT

    parser = SyllabusParser(llm, system_prompt=sys_prompt)

    def on_progress(done: int, total: int, label: str):
        task.status = "generating"
        task.progress_done = done
        task.progress_label = label
        db.session.commit()
        _update_progress(task_id, status="generating", done=done, total=total, label=label)

    lessons = parser.parse(
        syllabus_path=str(syllabus_path),
        total_lessons=task.total_lessons,
        batch_size=task.batch_size,
        progress_callback=on_progress,
    )

    if not lessons:
        raise LessonPlanError("AI 未生成任何教案")

    # 状态 → building
    task.status = "building"
    task.progress_done = task.total_lessons
    task.progress_label = "正在生成 docx"
    db.session.commit()
    _update_progress(task_id, status="building", done=task.total_lessons, label="正在生成 docx")

    # 生成 docx
    course_name = (task.course_info or {}).get("课程名称", "")
    course_name_safe = course_name or syllabus_path.stem
    out_name = f"{course_name_safe}-教案.docx"
    out_path = output_dir(task.id) / out_name

    builder = DocxBuilder(str(template_path()))
    builder.build_all(
        lessons=lessons,
        output_path=str(out_path),
        course_info=task.course_info or {},
        teacher_info=task.teacher_info or {},
    )

    # 完工
    rel = f"output/{out_name}"
    task.output_file = rel
    task.status = "completed"
    task.progress_label = "完成"
    task.finished_at = datetime.utcnow()
    db.session.commit()
    _update_progress(task_id, status="completed", done=task.total_lessons, label="完成")


def _mark_failed(task_id: int, err_msg: str) -> None:
    try:
        task = LessonPlanTask.query.get(task_id)
        if task:
            task.status = "failed"
            task.error_summary = err_msg[:2000]
            task.finished_at = datetime.utcnow()
            db.session.commit()
    except Exception:
        db.session.rollback()


def get_progress(task_id: int) -> dict:
    """查询进度：优先读内存缓存，fallback DB。"""
    prog = _PROGRESS.get(task_id)
    if prog:
        return {**prog, "taskId": task_id}

    task = LessonPlanTask.query.get(task_id)
    if not task:
        return {"taskId": task_id, "status": "none", "total": 0, "done": 0}
    return {
        "taskId": task_id,
        "status": task.status,
        "total": task.progress_total,
        "done": task.progress_done,
        "label": task.progress_label,
    }


def regenerate_task(task_id: int, user_id: str) -> int:
    """重新生成：备份旧 output → 复用原大纲和参数 → 启动 worker。"""
    task = LessonPlanTask.query.filter_by(id=task_id, user_id=user_id).first()
    if task is None:
        raise LessonPlanError("任务不存在或无权访问")

    if task.status in ACTIVE_STATUSES:
        raise LessonPlanError("当前任务正在运行，请等待完成")

    # 备份旧 output
    if task.output_file:
        backup_output_for_retry(task.id, task.output_file)

    # 重置状态
    task.status = "pending"
    task.progress_done = 0
    task.progress_label = "等待重新生成"
    task.error_summary = None
    task.started_at = None
    task.finished_at = None
    # output_file 字段保留（备份后启动 worker 时再覆写）
    db.session.commit()

    _update_progress(
        task_id,
        status="pending",
        total=task.total_lessons,
        done=0,
        label="等待重新生成",
    )
    _start_worker(task_id)
    return task_id


def delete_task(task_id: int, user_id: str) -> bool:
    task = LessonPlanTask.query.filter_by(id=task_id, user_id=user_id).first()
    if task is None:
        raise LessonPlanError("任务不存在或无权访问")
    if task.status in ACTIVE_STATUSES:
        raise LessonPlanError("当前任务正在运行，请等待完成或刷新页面")

    try:
        remove_task_storage(task_id)
        db.session.delete(task)
        db.session.commit()
    except Exception as exc:
        db.session.rollback()
        raise LessonPlanError(f"删除失败: {exc}")
    return True


def cleanup_zombie_tasks() -> None:
    """启动时把 active 状态的任务全部改为 failed。"""
    tasks = LessonPlanTask.query.filter(LessonPlanTask.status.in_(ACTIVE_STATUSES)).all()
    for t in tasks:
        t.status = "failed"
        t.error_summary = "进程重启，任务中断"
        t.finished_at = datetime.utcnow()
    if tasks:
        db.session.commit()
        print(f"[cleanup] 清理 {len(tasks)} 个 lesson_plan 僵尸任务")


def get_owned_task(task_id: int, user_id: str) -> Optional[LessonPlanTask]:
    return LessonPlanTask.query.filter_by(id=task_id, user_id=user_id).first()
