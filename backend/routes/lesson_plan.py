"""教案生成路由。

端点：
    POST   /api/lesson-plan/generate                    创建任务并启动后台生成
    GET    /api/lesson-plan/tasks?page&size              任务列表（分页）
    GET    /api/lesson-plan/tasks/<taskId>               单任务详情
    GET    /api/lesson-plan/tasks/<taskId>/progress      进度（内存缓存优先）
    POST   /api/lesson-plan/tasks/<taskId>/regenerate    重新生成
    GET    /api/lesson-plan/tasks/<taskId>/download      下载 docx
    DELETE /api/lesson-plan/tasks/<taskId>               删除任务
"""

from __future__ import annotations

import json
from pathlib import Path
from urllib.parse import quote

from flask import Blueprint, current_app, g, make_response, request, send_file

from extensions import db
from models.lesson_plan import LessonPlanTask
from services.lesson_plan_service import (
    ACTIVE_STATUSES,
    LessonPlanError,
    cleanup_zombie_tasks,
    create_task,
    delete_task,
    get_owned_task,
    get_progress,
    regenerate_task,
)
from services.lesson_plan_storage import (
    LessonPlanStorageError,
    output_rel_to_abs,
)
from utils.jwt_helper import jwt_required
from utils.response import fail, success

lesson_plan_bp = Blueprint("lesson_plan", __name__)


def _parse_json_field(raw: str, field_name: str) -> dict:
    if not raw:
        return {}
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        raise LessonPlanError(f"{field_name} 不是合法的 JSON 对象")
    if not isinstance(data, dict):
        raise LessonPlanError(f"{field_name} 必须是 JSON 对象")
    return data


@lesson_plan_bp.get("/default-prompt")
@jwt_required
def default_prompt():
    from services.lesson_plan.parser import DEFAULT_SYSTEM_PROMPT, DEFAULT_USER_PROMPT

    return success({
        "systemPrompt": DEFAULT_SYSTEM_PROMPT,
        "userPrompt": DEFAULT_USER_PROMPT,
    })


@lesson_plan_bp.post("/generate")
@jwt_required
def generate():
    if request.content_length and request.content_length > int(
        current_app.config.get("LESSON_PLAN_UPLOAD_MAX_BYTES", 20_000_000)
    ):
        return fail("上传内容过大")

    syllabus_name = (request.form.get("syllabusName") or "").strip()
    if "syllabus" not in request.files:
        return fail("syllabus 文件必填")
    file_storage = request.files["syllabus"]
    if not file_storage.filename:
        return fail("syllabus 文件为空")
    if not syllabus_name:
        syllabus_name = file_storage.filename

    try:
        total_lessons = int(request.form.get("totalLessons") or "0")
    except ValueError:
        return fail("totalLessons 必须为整数")
    try:
        batch_size = int(request.form.get("batchSize") or "2")
    except ValueError:
        return fail("batchSize 必须为整数")
    try:
        total_hours = int(request.form.get("totalHours") or "0")
    except ValueError:
        return fail("totalHours 必须为整数")

    try:
        course_info = _parse_json_field(request.form.get("courseInfo"), "courseInfo")
        teacher_info = _parse_json_field(request.form.get("teacherInfo"), "teacherInfo")
    except LessonPlanError as exc:
        return fail(exc.message)

    if not course_info:
        return fail("courseInfo 必填")
    if not teacher_info:
        return fail("teacherInfo 必填")

    system_prompt = (request.form.get("systemPrompt") or "").strip() or None
    user_prompt = (request.form.get("userPrompt") or "").strip() or None

    file_bytes = file_storage.read()
    if not file_bytes:
        return fail("syllabus 文件内容为空")

    try:
        task_id = create_task(
            user_id=g.current_user_id,
            syllabus_bytes=file_bytes,
            syllabus_name=syllabus_name,
            total_lessons=total_lessons,
            batch_size=batch_size,
            course_info=course_info,
            teacher_info=teacher_info,
            system_prompt=system_prompt,
            user_prompt=user_prompt,
            total_hours=total_hours or None,
        )
    except LessonPlanError as exc:
        return fail(exc.message)
    except LessonPlanStorageError as exc:
        return fail(exc.message)
    except Exception as exc:
        current_app.logger.exception("lesson_plan generate failed: %s", exc)
        return fail("创建任务失败")

    return success({"taskId": task_id, "status": "pending"})


@lesson_plan_bp.get("/tasks")
@jwt_required
def list_tasks():
    try:
        page = max(int(request.args.get("page", "1")), 1)
        size = min(max(int(request.args.get("size", "20")), 1), 100)
    except ValueError:
        return fail("page/size 必须为整数")

    query = LessonPlanTask.query.filter_by(user_id=g.current_user_id).order_by(
        LessonPlanTask.created_at.desc()
    )
    pagination = query.paginate(page=page, per_page=size, error_out=False)
    items = [t.to_summary() for t in pagination.items]
    return success({"total": pagination.total, "list": items, "page": page, "size": size})


@lesson_plan_bp.get("/tasks/<int:task_id>")
@jwt_required
def get_task(task_id: int):
    task = get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")
    return success(task.to_summary())


@lesson_plan_bp.get("/tasks/<int:task_id>/progress")
@jwt_required
def task_progress(task_id: int):
    task = get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")
    try:
        progress = get_progress(task_id)
    except Exception as exc:
        current_app.logger.exception("progress failed: %s", exc)
        return fail("查询进度失败")
    return success(progress)


@lesson_plan_bp.post("/tasks/<int:task_id>/regenerate")
@jwt_required
def regenerate(task_id: int):
    task = get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")

    system_prompt = (request.form.get("systemPrompt") or "").strip() or None
    user_prompt = (request.form.get("userPrompt") or "").strip() or None
    try:
        total_hours_raw = request.form.get("totalHours")
        total_hours = int(total_hours_raw) if total_hours_raw else None
    except ValueError:
        total_hours = None
    changed = False
    if system_prompt is not None:
        task.system_prompt = system_prompt
        changed = True
    if user_prompt is not None:
        task.user_prompt = user_prompt
        changed = True
    if total_hours is not None:
        task.total_hours = total_hours
        changed = True
    if changed:
        db.session.commit()

    try:
        regenerate_task(task_id, g.current_user_id)
    except LessonPlanError as exc:
        return fail(exc.message)
    except Exception as exc:
        current_app.logger.exception("regenerate failed: %s", exc)
        return fail("重新生成失败")

    return success({"taskId": task_id, "status": "pending"})


@lesson_plan_bp.get("/tasks/<int:task_id>/download")
@jwt_required
def download(task_id: int):
    task = get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")
    if not task.output_file:
        return fail("该任务尚无输出文件")
    if task.status != "completed":
        return fail(f"任务状态为 {task.status}，无法下载")

    try:
        abs_path = output_rel_to_abs(task_id, task.output_file)
    except LessonPlanStorageError as exc:
        return fail(exc.message)

    if not abs_path.exists():
        return fail("输出文件丢失")

    download_name = abs_path.name
    resp = make_response(
        send_file(
            str(abs_path),
            as_attachment=True,
            download_name=download_name,
        )
    )
    resp.headers["Content-Disposition"] = (
        f"attachment; filename*=UTF-8''{quote(download_name)}"
    )
    return resp


@lesson_plan_bp.delete("/tasks/<int:task_id>")
@jwt_required
def delete(task_id: int):
    try:
        delete_task(task_id, g.current_user_id)
    except LessonPlanError as exc:
        return fail(exc.message)
    except Exception as exc:
        current_app.logger.exception("delete failed: %s", exc)
        return fail("删除失败")
    return success({"taskId": task_id, "deleted": True})


__all__ = ["lesson_plan_bp", "cleanup_zombie_tasks", "ACTIVE_STATUSES"]
