"""OBE 路由：目录生成持久化 + 实验报告批改全流程（experiment 维度）。

端点：
    POST   /api/obe/mkdir                                              创建任务
    GET    /api/obe/tasks?page&size                                    任务列表
    GET    /api/obe/tasks/<taskId>                                      任务详情
    GET    /api/obe/tasks/<taskId>/tree                                 目录树
    POST   /api/obe/tasks/<taskId>/upload                               上传学生文件（自动按 ZIP 分 experiment）
    POST   /api/obe/tasks/<taskId>/upload/resolve                       解决 ambiguous
    POST   /api/obe/tasks/<taskId>/grade                                触发批改
    GET    /api/obe/tasks/<taskId>/progress?dirType&experimentLabel&jobId  批改进度
    POST   /api/obe/tasks/<taskId>/students/<studentPk>/retry           单学生重试
    GET    /api/obe/tasks/<taskId>/download/zip?dirType&experimentLabel  下载 ZIP
    GET    /api/obe/tasks/<taskId>/download/excel?dirType&experimentLabel&jobId  下载 Excel
    DELETE /api/obe/tasks/<taskId>                                      删除任务
"""

from __future__ import annotations

import json
import zipfile
from collections import defaultdict
from io import BytesIO
from pathlib import Path
from urllib.parse import quote

from flask import Blueprint, current_app, g, make_response, request, send_file

from extensions import db
from models.obe import ObeGradingJob, ObeStudent, ObeTask
from services.obe_grading import (
    ObeGradingError,
    build_excel_summary,
    get_progress,
    has_running_job,
    list_running_experiments,
    retry_student,
    start_grading_job,
)
from services.obe_match import (
    ObeMatchError,
    match_uploads,
    resolve_ambiguous,
)
from services.obe_mkdir_service import MAX_UPLOAD_BYTES, ObeMkdirError, create_task
from services.obe_storage import (
    ObeStorageError,
    _safe_segment,
    build_tree,
    excel_dir,
    remove_task_storage,
    root_dir,
    save_signature,
    task_root,
)
from utils.jwt_helper import jwt_required
from utils.response import fail, success

obe_bp = Blueprint("obe", __name__)


def _get_owned_task(task_id: int, user_id: str):
    try:
        tid = int(task_id)
    except (TypeError, ValueError):
        return None
    return ObeTask.query.filter_by(id=tid, user_id=user_id).first()


def _parse_dir_types(raw: str, field_name: str) -> list:
    if not raw:
        return []
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        raise ObeMkdirError(f"{field_name} 不是合法的 JSON 数组")
    if not isinstance(parsed, list):
        raise ObeMkdirError(f"{field_name} 不是合法的 JSON 数组")
    result = []
    for item in parsed:
        if not isinstance(item, str):
            raise ObeMkdirError(f"{field_name} 数组元素必须为字符串")
        s = item.strip()
        if s:
            result.append(s)
    return result


# ============ /mkdir ============


@obe_bp.post("/mkdir")
@jwt_required
def obe_mkdir():
    class_name = (request.form.get("className") or "").strip()
    course_name = (request.form.get("courseName") or "").strip()
    teacher_name = (request.form.get("teacherName") or "").strip()

    if not class_name:
        return fail("className 必填")
    if not course_name:
        return fail("courseName 必填")
    if not teacher_name:
        return fail("teacherName 必填")

    if request.content_length and request.content_length > MAX_UPLOAD_BYTES * 2:
        return fail(f"上传内容过大（超过 {MAX_UPLOAD_BYTES // (1024 * 1024)}MB 限制）")

    if "roster" not in request.files:
        return fail("roster 文件必填")
    file_storage = request.files["roster"]

    try:
        fixed_dir_types = _parse_dir_types(request.form.get("fixedDirTypes") or "[]", "fixedDirTypes")
        student_dir_types = _parse_dir_types(
            request.form.get("studentDirTypes") or "[]", "studentDirTypes"
        )
    except ObeMkdirError as exc:
        return fail(exc.message)

    if not student_dir_types and not fixed_dir_types:
        return fail("至少需要选择一个固定目录或考核目录")

    try:
        result = create_task(
            g.current_user_id,
            file_storage,
            class_name,
            course_name,
            teacher_name,
            fixed_dir_types,
            student_dir_types,
        )
    except ObeMkdirError as exc:
        return fail(exc.message)
    except Exception as exc:
        current_app.logger.exception("obe_mkdir failed: %s", exc)
        return fail("服务异常，请稍后重试")

    return success(result)


@obe_bp.get("/ping")
@jwt_required
def ping():
    return success({"pong": True, "user": g.current_user_name})


# ============ 任务列表 / 详情 ============


@obe_bp.get("/tasks")
@jwt_required
def list_tasks():
    try:
        page = max(int(request.args.get("page", "1")), 1)
        size = min(max(int(request.args.get("size", "20")), 1), 100)
    except ValueError:
        return fail("page/size 必须为整数")

    query = ObeTask.query.filter_by(user_id=g.current_user_id).order_by(
        ObeTask.created_at.desc()
    )
    pagination = query.paginate(page=page, per_page=size, error_out=False)
    items = [t.to_summary() for t in pagination.items]
    return success({"total": pagination.total, "list": items, "page": page, "size": size})


@obe_bp.get("/tasks/<int:task_id>")
@jwt_required
def get_task(task_id: int):
    task = _get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")

    students = (
        ObeStudent.query.filter_by(task_id=task_id)
        .order_by(ObeStudent.dir_type, ObeStudent.experiment_label, ObeStudent.student_id)
        .all()
    )

    # 按 (dir_type, experiment_label) 二级分组
    students_by_dir_exp: dict[str, dict[str, list]] = defaultdict(lambda: defaultdict(list))
    for s in students:
        students_by_dir_exp[s.dir_type][s.experiment_label].append(s.to_dict())

    # 转为普通 dict
    students_by_dir_exp_plain = {dt: dict(exps) for dt, exps in students_by_dir_exp.items()}

    # 各 (dir_type, experiment_label) 的最近 job
    jobs = (
        ObeGradingJob.query.filter_by(task_id=task_id)
        .order_by(ObeGradingJob.id.desc())
        .all()
    )
    latest_job_by_dir_exp: dict[str, dict[str, dict]] = defaultdict(dict)
    for j in jobs:
        if j.experiment_label not in latest_job_by_dir_exp[j.dir_type]:
            latest_job_by_dir_exp[j.dir_type][j.experiment_label] = j.to_dict()

    # 各 dir_type 下所有 experiment_label 列表（含 "default"）
    experiments_by_dir: dict[str, list[str]] = {}
    for dt, exps in students_by_dir_exp.items():
        experiments_by_dir[dt] = sorted(exps.keys())

    return success(
        {
            "task": task.to_summary(),
            "studentsByDirAndExperiment": students_by_dir_exp_plain,
            "latestJobByDirAndExperiment": {dt: dict(v) for dt, v in latest_job_by_dir_exp.items()},
            "experimentsByDir": experiments_by_dir,
        }
    )


@obe_bp.get("/tasks/<int:task_id>/tree")
@jwt_required
def get_tree(task_id: int):
    task = _get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")
    try:
        tree = build_tree(task_id)
    except Exception as exc:
        current_app.logger.exception("build_tree failed: %s", exc)
        return fail("构建目录树失败")
    return success({"tree": tree})


# ============ 上传学生文件 ============


@obe_bp.post("/tasks/<int:task_id>/upload")
@jwt_required
def upload_student_files(task_id: int):
    task = _get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")

    dir_type = (request.form.get("dirType") or "").strip()
    if not dir_type:
        return fail("dirType 必填")

    if dir_type not in (task.student_dir_types or []):
        return fail(f"dirType '{dir_type}' 不在该任务的考核目录列表中")

    files = request.files.getlist("files")
    if not files or all(not f.filename for f in files):
        return fail("未选择文件")

    try:
        result = match_uploads(task_id, dir_type, files)
    except ObeMatchError as exc:
        return fail(exc.message)
    except ObeStorageError as exc:
        return fail(exc.message)
    except Exception as exc:
        current_app.logger.exception("upload failed: %s", exc)
        return fail("上传处理失败")

    return success(result)


@obe_bp.post("/tasks/<int:task_id>/upload/resolve")
@jwt_required
def upload_resolve(task_id: int):
    task = _get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")

    dir_type = (request.form.get("dirType") or "").strip()
    experiment_label = (request.form.get("experimentLabel") or "").strip()
    if not dir_type or not experiment_label:
        return fail("dirType 和 experimentLabel 必填")

    raw = request.form.get("resolutions") or "[]"
    try:
        resolutions = json.loads(raw)
        if not isinstance(resolutions, list):
            raise ValueError("resolutions 必须是数组")
    except ValueError as exc:
        return fail(f"resolutions 解析失败：{exc}")

    try:
        result = resolve_ambiguous(task_id, dir_type, experiment_label, resolutions)
    except Exception as exc:
        current_app.logger.exception("resolve failed: %s", exc)
        return fail("解决归属失败")

    return success(result)


# ============ 触发批改 ============


@obe_bp.post("/tasks/<int:task_id>/grade")
@jwt_required
def grade_task(task_id: int):
    task = _get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")

    dir_type = (request.form.get("dirType") or "").strip()
    experiment_label = (request.form.get("experimentLabel") or "").strip()
    teacher_name = (request.form.get("teacherName") or "").strip()
    sign_date = (request.form.get("signDate") or "").strip()
    teacher_prompt = (request.form.get("teacherPrompt") or "").strip()
    system_prompt = (request.form.get("systemPrompt") or "").strip() or None
    # 默认 skip_graded=true（只批 pending/failed）；用户勾选「覆盖已批改」时传 false
    skip_graded_raw = (request.form.get("skipGraded") or "true").strip().lower()
    skip_graded = skip_graded_raw not in ("false", "0", "no", "off")

    if not dir_type:
        return fail("dirType 必填")
    if not experiment_label:
        return fail("experimentLabel 必填")
    if not teacher_name:
        return fail("teacherName 必填")
    if not sign_date:
        return fail("signDate 必填")
    if not teacher_prompt:
        return fail("teacherPrompt 必填")

    if "signPicture" not in request.files:
        return fail("签名图片必填")
    sign_picture_bytes = request.files["signPicture"].read()
    if not sign_picture_bytes:
        return fail("签名图片为空")

    try:
        sign_rel, _ = save_signature(task_id, sign_picture_bytes)
        job_id = start_grading_job(
            task_id=task_id,
            dir_type=dir_type,
            experiment_label=experiment_label,
            teacher_name=teacher_name,
            sign_date=sign_date,
            sign_picture_rel=sign_rel,
            teacher_prompt=teacher_prompt,
            system_prompt=system_prompt,
            skip_graded=skip_graded,
        )
    except ObeGradingError as exc:
        return fail(exc.message)
    except Exception as exc:
        current_app.logger.exception("grade failed: %s", exc)
        return fail("触发批改失败")

    return success({"jobId": job_id, "status": "running"})


@obe_bp.get("/tasks/<int:task_id>/progress")
@jwt_required
def grade_progress(task_id: int):
    task = _get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")

    dir_type = (request.args.get("dirType") or "").strip() or None
    experiment_label = (request.args.get("experimentLabel") or "").strip() or None
    job_id_raw = request.args.get("jobId")
    job_id = None
    if job_id_raw:
        try:
            job_id = int(job_id_raw)
        except ValueError:
            return fail("jobId 必须为整数")

    if not dir_type and not job_id:
        return fail("dirType 或 jobId 至少传一个")

    try:
        progress = get_progress(task_id, dir_type, experiment_label, job_id)
    except Exception as exc:
        current_app.logger.exception("progress failed: %s", exc)
        return fail("查询进度失败")

    return success(progress)


@obe_bp.post("/tasks/<int:task_id>/students/<int:student_pk>/retry")
@jwt_required
def retry_student_endpoint(task_id: int, student_pk: int):
    task = _get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")

    system_prompt = (request.form.get("systemPrompt") or "").strip() or None

    try:
        result = retry_student(task_id, student_pk, system_prompt)
    except ObeGradingError as exc:
        return fail(exc.message)
    except Exception as exc:
        current_app.logger.exception("retry failed: %s", exc)
        return fail("重试失败")

    return success(result)


# ============ 下载 ============


@obe_bp.get("/tasks/<int:task_id>/download/zip")
@jwt_required
def download_zip(task_id: int):
    task = _get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")

    dir_type = (request.args.get("dirType") or "").strip()
    experiment_label = (request.args.get("experimentLabel") or "").strip()
    if not dir_type:
        return fail("dirType 必填")

    try:
        zip_bytes, zip_name = _build_dir_type_zip(task_id, dir_type, experiment_label or None)
    except ObeStorageError as exc:
        return fail(exc.message)
    except Exception as exc:
        current_app.logger.exception("download zip failed: %s", exc)
        return fail("打包失败")

    resp = make_response(zip_bytes)
    resp.headers["Content-Type"] = "application/zip"
    resp.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(zip_name)}"
    resp.headers["Content-Length"] = str(len(zip_bytes))
    return resp


def _build_dir_type_zip(
    task_id: int, dir_type: str, experiment_label: str | None = None
) -> tuple[bytes, str]:
    """把 root/ 下指定 dir_type 的目录打包成 ZIP 字节流。

    experiment_label 仅作为下载文件名前缀（不影响打包内容，整个 dir_type 都打）。
    """
    root = root_dir(task_id)
    candidates = [d for d in root.iterdir() if d.is_dir() and dir_type in d.name]
    if not candidates:
        raise ObeStorageError(f"未在 root/ 下找到 dirType='{dir_type}' 的目录")

    import os

    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for top in candidates:
            base = top.parent
            for dirpath, dirnames, filenames in os.walk(top):
                for d in dirnames:
                    full = Path(dirpath) / d
                    arcname = full.relative_to(base).as_posix() + "/"
                    zf.writestr(arcname, b"")
                for fn in filenames:
                    full_path = Path(dirpath) / fn
                    arcname = full_path.relative_to(base).as_posix()
                    zf.write(full_path, arcname)

    name_prefix = experiment_label or dir_type
    return buf.getvalue(), f"{name_prefix}.zip"


@obe_bp.get("/tasks/<int:task_id>/download/excel")
@jwt_required
def download_excel(task_id: int):
    task = _get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")

    dir_type = (request.args.get("dirType") or "").strip()
    experiment_label = (request.args.get("experimentLabel") or "").strip()
    job_id_raw = request.args.get("jobId")
    if not dir_type or not experiment_label:
        return fail("dirType 和 experimentLabel 必填")

    job_id: int | None = None
    if job_id_raw:
        try:
            job_id = int(job_id_raw)
        except ValueError:
            return fail("jobId 必须为整数")

    if job_id is None:
        job = (
            ObeGradingJob.query.filter_by(
                task_id=task_id, dir_type=dir_type, experiment_label=experiment_label
            )
            .order_by(ObeGradingJob.id.desc())
            .first()
        )
        if job is None:
            return fail("该实验尚无批改记录")
        job_id = job.id

    excel_path = excel_dir(task_id) / f"{_safe_segment(experiment_label)}_{job_id}_成绩.xlsx"
    if not excel_path.exists():
        try:
            excel_path = build_excel_summary(task_id, dir_type, experiment_label, job_id)
        except Exception as exc:
            current_app.logger.exception("excel generate failed: %s", exc)
            return fail("生成 Excel 失败")

    download_name = f"{experiment_label}_成绩.xlsx"
    return send_file(
        str(excel_path),
        as_attachment=True,
        download_name=download_name,
    )


# ============ 删除任务 ============


@obe_bp.delete("/tasks/<int:task_id>")
@jwt_required
def delete_task(task_id: int):
    task = _get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")

    # 检查该任务下所有 dir_type 是否有运行中的 job
    for dir_type in task.student_dir_types or []:
        running = list_running_experiments(task_id, dir_type)
        if running:
            return fail(
                f"目录 '{dir_type}' 的实验 '{running[0]}' 有批改任务正在运行，请等待完成或刷新页面重试"
            )

    try:
        remove_task_storage(task_id)
        db.session.delete(task)
        db.session.commit()
    except Exception as exc:
        db.session.rollback()
        current_app.logger.exception("delete task failed: %s", exc)
        return fail("删除任务失败")

    return success({"taskId": task_id, "deleted": True})
