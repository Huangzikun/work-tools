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
    POST   /api/obe/tasks/<taskId>/jobs/<jobId>/cancel                  取消批改任务（协作式，跑完当前学生后停）
    POST   /api/obe/tasks/<taskId>/students/<studentPk>/retry           单学生重试
    POST   /api/obe/tasks/<taskId>/students/<studentPk>/upload          单学生上传/替换报告
    GET    /api/obe/tasks/<taskId>/students/<studentPk>/download        单学生下载报告（批改后优先，回退原始）
    GET    /api/obe/tasks/<taskId>/download/zip?dirType&experimentLabel  下载 ZIP
    GET    /api/obe/tasks/<taskId>/download/excel?dirType&experimentLabel&jobId  下载 Excel
    GET    /api/obe/tasks/<taskId>/download/all                          整体打包下载整个任务（全部目录 + 成绩 Excel）
    DELETE /api/obe/tasks/<taskId>                                      删除任务
"""

from __future__ import annotations

import json
import zipfile
from collections import defaultdict
from io import BytesIO
from pathlib import Path
from urllib.parse import quote

from flask import Blueprint, Response, current_app, g, make_response, request, send_file

from extensions import db
from models.obe import ObeGradingJob, ObeStudent, ObeTask
from services.obe_grading import (
    ObeGradingError,
    build_excel_summary,
    get_progress,
    get_rubric,
    has_running_job,
    list_running_experiments,
    request_cancel,
    retry_student,
    start_grading_job,
    upsert_rubric,
)
from services.obe_match import (
    ObeMatchError,
    assign_file_to_student,
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


@obe_bp.get("/rubric")
@jwt_required
def get_rubric_handler():
    """按课程载入当前教师最近一次的评分标准。"""
    course_name = (request.args.get("courseName") or "").strip()
    if not course_name:
        return fail("courseName 必填")
    rec = get_rubric(g.current_user_id, course_name)
    return success(rec.to_dict() if rec else None)


@obe_bp.post("/rubric")
@jwt_required
def save_rubric_handler():
    """显式保存评分标准（按 user_id+course_name upsert）。grade 时也会自动 upsert。"""
    data = request.get_json(silent=True) or {}
    course_name = (data.get("courseName") or "").strip()
    if not course_name:
        return fail("courseName 必填")
    dimensions = []
    for d in data.get("dimensions") or []:
        name = (d.get("name") or "").strip()
        try:
            ms = int(d.get("max_score", 0))
        except (TypeError, ValueError):
            ms = 0
        if not name or ms <= 0:
            continue
        dimensions.append(
            {"name": name, "max_score": ms, "criteria": (d.get("criteria") or "").strip()}
        )
    free_text = (data.get("freeText") or "").strip()
    score_levels = data.get("scoreLevels")
    if score_levels:
        try:
            score_levels = [int(x) for x in score_levels]
        except (TypeError, ValueError):
            score_levels = None
    rec = upsert_rubric(g.current_user_id, course_name, dimensions, free_text, score_levels)
    return success(rec.to_dict())


@obe_bp.post("/rubric/generate")
@jwt_required
def generate_rubric_handler():
    """AI 根据教师提供的实验内容生成评分标准（积极评分 + 严谨扣分）。"""
    data = request.get_json(silent=True) or {}
    content = (data.get("experimentContent") or "").strip()
    if len(content) < 10:
        return fail("请提供实验内容（至少 10 字）")
    try:
        from services.obe_sign_core import generate_rubric

        result = generate_rubric(content)
    except Exception as exc:
        current_app.logger.exception("generate_rubric failed: %s", exc)
        return fail("生成评分标准失败")
    return success(result)


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

    # 教师结构化评分维度与档位（可选，完全由教师控制评分标准）
    rubric_dimensions = None
    rubric_dims_raw = (request.form.get("rubricDimensions") or "").strip()
    if rubric_dims_raw:
        try:
            parsed = json.loads(rubric_dims_raw)
            if not isinstance(parsed, list):
                raise ValueError
            rubric_dimensions = parsed
        except (ValueError, TypeError):
            return fail("rubricDimensions 必须是 JSON 数组")
    score_levels = None
    score_levels_raw = (request.form.get("scoreLevels") or "").strip()
    if score_levels_raw:
        try:
            score_levels = [int(x) for x in json.loads(score_levels_raw)]
        except (ValueError, TypeError):
            return fail("scoreLevels 必须是数字数组")

    if not dir_type:
        return fail("dirType 必填")
    if not experiment_label:
        return fail("experimentLabel 必填")
    if not teacher_name:
        return fail("teacherName 必填")
    if not sign_date:
        return fail("signDate 必填")
    # 评分标准放宽：自由文本或结构化维度至少一个非空
    if not teacher_prompt and not rubric_dimensions:
        return fail("评分标准（总体要求或评分维度）至少填一项")

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
            rubric_dimensions=rubric_dimensions,
            score_levels=score_levels,
        )
        # 自动按课程保存评分标准，下次批改同课程一键载入
        if rubric_dimensions:
            try:
                upsert_rubric(
                    g.current_user_id,
                    task.course_name,
                    rubric_dimensions,
                    teacher_prompt,
                    score_levels,
                )
            except Exception as exc:
                current_app.logger.exception("upsert_rubric failed: %s", exc)
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


@obe_bp.post("/tasks/<int:task_id>/jobs/<int:job_id>/cancel")
@jwt_required
def cancel_grade_job(task_id: int, job_id: int):
    task = _get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")

    try:
        result = request_cancel(task_id, job_id)
    except ObeGradingError as exc:
        return fail(exc.message)
    except Exception as exc:
        current_app.logger.exception("cancel failed: %s", exc)
        return fail("取消失败")

    return success(result)


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


@obe_bp.post("/tasks/<int:task_id>/students/<int:student_pk>/upload")
@jwt_required
def upload_student_file(task_id: int, student_pk: int):
    task = _get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")

    student = ObeStudent.query.filter_by(id=student_pk, task_id=task_id).first()
    if student is None:
        return fail("学生不存在")

    files = request.files.getlist("file")
    if not files or all(not f.filename for f in files):
        return fail("未选择文件")
    file_storage = files[0]

    try:
        result = assign_file_to_student(task_id, student, file_storage)
    except ObeMatchError as exc:
        return fail(exc.message)
    except ObeStorageError as exc:
        return fail(exc.message)
    except Exception as exc:
        current_app.logger.exception("upload student file failed: %s", exc)
        return fail("上传失败")

    return success(result)


@obe_bp.get("/tasks/<int:task_id>/students/<int:student_pk>/download")
@jwt_required
def download_student_file(task_id: int, student_pk: int):
    task = _get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")

    student = ObeStudent.query.filter_by(id=student_pk, task_id=task_id).first()
    if student is None:
        return fail("学生不存在")

    # 优先批改后文件，回退原始上传件
    rel = None
    is_graded = False
    if student.last_graded_file:
        rel = student.last_graded_file
        is_graded = student.grade_status == "graded"
    elif student.uploaded_file:
        rel = student.uploaded_file

    if not rel:
        return fail("该学生暂无可下载的文件")

    abs_path = (task_root(task_id) / rel).resolve()
    base = task_root(task_id).resolve()
    if abs_path != base and base not in abs_path.parents:
        return fail("文件路径非法")
    if not abs_path.exists():
        return fail("文件不存在（可能已被移动或删除）")

    ext = abs_path.suffix or ".docx"
    suffix = "_批改" if is_graded else ""
    filename = f"{student.student_id}{student.student_name}{suffix}{ext}"

    resp = send_file(abs_path, as_attachment=True, download_name=filename)
    resp.headers["X-Content-Type-Options"] = "nosniff"
    return resp


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


def _task_zip_name(task: ObeTask) -> str:
    return f"{task.class_name}《{task.course_name}》{task.teacher_name}.zip"


def _task_has_content(task_id: int) -> bool:
    """任务是否有可打包内容（root/ 或 excel/ 任一非空）。"""
    root = root_dir(task_id)
    excel = excel_dir(task_id)
    has_root = root.exists() and any(root.iterdir())
    has_excel = excel.exists() and any(excel.glob("*.xlsx"))
    return has_root or has_excel


def _stream_task_zip(root: Path, excel: Path):
    """流式生成整个任务的 ZIP 字节块（生成器）。

    内容：root/ 下全部目录（mkdir 创建的 OBE 结构 + 学生批改文件）+ excel/ 下全部成绩表。
    arcname 相对 root/（不含 root 这一层），解压即得到原始 OBE 目录名；成绩 Excel 放 ZIP 顶层。

    实现要点：zipfile 是 push 模式且中央目录在最后才写入，无法用纯协程流式，
    因此用一个后台线程跑 zipfile、其 write 经 queue 转发给本生成器逐块 yield。
    内存占用 ≈ queue maxsize × 块大小（数 MB 量级），不随包体线性增长。

    重要：本生成器只接收 Path、不访问 current_app —— 生成器体在 WSGI 迭代时才执行，
    彼时 Flask 的 app/request context 已被 pop。调用方须在视图里（仍有 context 时）
    算好 root/excel 并完成 _task_has_content 校验后再传入。
    """
    import os
    import queue
    import threading

    end = object()
    q: "queue.Queue[object]" = queue.Queue(maxsize=64)

    class _Sink:
        def write(self, data):
            q.put(data)
            return len(data)  # Python 3.13+ zipfile 依赖 write 返回写入字节数（self.offset += n）

        def flush(self):  # zipfile 不调用，仅为满足 file-like 协议
            pass

        def seekable(self):  # 告知不可随机定位 → zipfile 走流式（中央目录置末尾，不回写文件头）
            return False

    def _produce():
        try:
            with zipfile.ZipFile(_Sink(), "w", zipfile.ZIP_DEFLATED) as zf:
                if root.exists():
                    for dirpath, dirnames, filenames in os.walk(root):
                        for d in dirnames:
                            full = Path(dirpath) / d
                            arcname = full.relative_to(root).as_posix() + "/"
                            zf.writestr(arcname, b"")
                        for fn in filenames:
                            full_path = Path(dirpath) / fn
                            arcname = full_path.relative_to(root).as_posix()
                            zf.write(full_path, arcname)
                if excel.exists():
                    for xlsx in sorted(excel.glob("*.xlsx")):
                        zf.write(xlsx, xlsx.name)
        except Exception as exc:  # 生产端异常透传给消费生成器
            q.put(exc)
        finally:
            q.put(end)

    threading.Thread(target=_produce, daemon=True).start()

    while True:
        item = q.get()
        if item is end:
            break
        if isinstance(item, Exception):
            raise item
        if item:
            yield item


@obe_bp.get("/tasks/<int:task_id>/download/all")
@jwt_required
def download_all(task_id: int):
    task = _get_owned_task(task_id, g.current_user_id)
    if task is None:
        return fail("任务不存在或无权访问")

    # 同步校验：空任务直接返回 JSON，避免响应头发出后无法改 status
    if not _task_has_content(task_id):
        return fail("任务目录为空，无可下载内容")

    # 在视图里（仍有 app context）解析好 Path 再传给生成器——生成器体在 WSGI
    # 迭代时执行，彼时 context 已 pop，不能在生成器内调 root_dir/excel_dir
    root = root_dir(task_id)
    excel = excel_dir(task_id)

    # 流式响应：不设 Content-Length（打包时未知总大小），走 chunked 传输；
    # 边打包边发送，后端内存不随包体线性增长
    resp = Response(_stream_task_zip(root, excel), mimetype="application/zip")
    resp.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(_task_zip_name(task))}"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    return resp


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
