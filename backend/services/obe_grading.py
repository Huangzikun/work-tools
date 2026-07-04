"""OBE 异步批改服务（threading + 内存进度缓存，experiment 维度）。

每个 (task_id, dir_type, experiment_label) 三元组对应一个独立批改任务。

并发保护：
- 启动 job 前检查同 task_id+dir_type+experiment_label 是否有 running job
- worker 内串行处理学生（LibreOffice 共享 profile 会冲突）
- 进度查询优先读内存缓存，fallback DB 聚合
"""

from __future__ import annotations

import threading
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
from flask import current_app

from extensions import db
from models.obe import ObeGradingJob, ObeGradingJobDetail, ObeStudent, ObeTask
from services.obe_sign_core import (
    DEFAULT_SYSTEM_PROMPT,
    SignContext,
    grade_student_docx,
)
from services.obe_storage import (
    _safe_segment,
    excel_dir,
    job_lo_profile_dir,
    resolve_signature_path,
    root_dir,
    task_root,
    write_job_log,
)


class ObeGradingError(Exception):
    def __init__(self, message: str):
        super().__init__(message)
        self.message = message


_JOB_PROGRESS: dict[int, dict] = {}
_PROGRESS_LOCK = threading.Lock()
_GRADED_WORKERS: dict[int, threading.Thread] = {}
# 协作式取消标志：request_cancel 置位，worker 循环顶部检查
_CANCEL_FLAGS: dict[int, threading.Event] = {}


def _update_progress(job_id: int, **fields) -> None:
    with _PROGRESS_LOCK:
        prog = _JOB_PROGRESS.setdefault(
            job_id,
            {"total": 0, "graded": 0, "failed": 0, "currentStudent": None, "status": "running"},
        )
        prog.update(fields)
        prog["lastUpdate"] = datetime.utcnow().isoformat()


def _clear_progress(job_id: int) -> None:
    with _PROGRESS_LOCK:
        _JOB_PROGRESS.pop(job_id, None)


def _is_cancelled(job_id: int) -> bool:
    flag = _CANCEL_FLAGS.get(job_id)
    return bool(flag and flag.is_set())


def _make_llm_client():
    from common.llm_client import LLMClient

    return LLMClient(
        api_key=current_app.config.get("ARK_API_KEY", ""),
        base_url=current_app.config.get("ARK_BASE_URL"),
        model=current_app.config.get("ARK_MODEL"),
    )


def has_running_job(task_id: int, dir_type: str, experiment_label: Optional[str] = None) -> Optional[int]:
    """检查是否有正在运行的 job。experiment_label=None 时检查该 dir_type 下所有 experiment。"""
    query = ObeGradingJob.query.filter_by(task_id=task_id, dir_type=dir_type, status="running")
    if experiment_label is not None:
        query = query.filter_by(experiment_label=experiment_label)
    job = query.first()
    return job.id if job else None


def list_running_experiments(task_id: int, dir_type: str) -> list[str]:
    """列出该 dir_type 下所有正在 running 的 experiment_label。"""
    jobs = (
        ObeGradingJob.query.filter_by(task_id=task_id, dir_type=dir_type, status="running")
        .with_entities(ObeGradingJob.experiment_label)
        .all()
    )
    return [j[0] for j in jobs]


def start_grading_job(
    task_id: int,
    dir_type: str,
    experiment_label: str,
    teacher_name: str,
    sign_date: str,
    sign_picture_rel: str,
    teacher_prompt: str,
    system_prompt: Optional[str] = None,
    skip_graded: bool = True,
    rubric_dimensions: Optional[list] = None,
    score_levels: Optional[list] = None,
) -> int:
    """创建 ObeGradingJob 并启动后台 worker。返回 job_id。

    skip_graded=True 时只批改 grade_status IN (pending, failed) 的学生；
    False 时全量重跑（包括已经 graded 的）。
    """
    running = has_running_job(task_id, dir_type, experiment_label)
    if running:
        raise ObeGradingError(
            f"该实验「{experiment_label}」已有批改任务在运行（jobId={running}），请等待完成"
        )

    matched_query = ObeStudent.query.filter_by(
        task_id=task_id,
        dir_type=dir_type,
        experiment_label=experiment_label,
        matched=True,
    )
    matched_total = matched_query.count()
    if matched_total == 0:
        raise ObeGradingError("该实验下没有已匹配的学生文件，请先上传学生文件")

    if skip_graded:
        target_query = matched_query.filter(
            ObeStudent.grade_status.in_(["pending", "failed"])
        )
        target_count = target_query.count()
        if target_count == 0:
            raise ObeGradingError(
                f"该实验下全部 {matched_total} 个学生已批改完成，"
                "如需重新批改请勾选「覆盖已批改学生」"
            )
        matched_count = target_count
    else:
        matched_count = matched_total

    job = ObeGradingJob(
        task_id=task_id,
        dir_type=dir_type,
        experiment_label=experiment_label,
        teacher_name=teacher_name,
        sign_date=sign_date,
        sign_picture_path=sign_picture_rel,
        teacher_prompt=teacher_prompt,
        rubric_dimensions=rubric_dimensions,
        score_levels=score_levels,
        status="running",
        total=matched_count,
        graded=0,
        failed=0,
        started_at=datetime.utcnow(),
    )
    db.session.add(job)
    db.session.flush()

    ObeTask.query.filter_by(id=task_id).update({"status": "grading"})
    db.session.commit()

    job_id = job.id
    _update_progress(
        job_id, total=matched_count, graded=0, failed=0, currentStudent=None, status="running"
    )

    app = current_app._get_current_object()
    sys_prompt = system_prompt or DEFAULT_SYSTEM_PROMPT

    def worker():
        with app.app_context():
            try:
                _run_grading(
                    task_id, job_id, dir_type, experiment_label, sys_prompt, teacher_prompt, skip_graded, rubric_dimensions, score_levels
                )
            except Exception as exc:
                app.logger.exception("grading job %s failed", job_id)
                _mark_job_failed(task_id, job_id, dir_type, experiment_label, str(exc))
            finally:
                db.session.remove()
                _clear_progress(job_id)
                _GRADED_WORKERS.pop(job_id, None)
                _CANCEL_FLAGS.pop(job_id, None)

    t = threading.Thread(target=worker, daemon=True, name=f"grade-{job_id}")
    _GRADED_WORKERS[job_id] = t
    t.start()

    return job_id


def _run_grading(
    task_id: int,
    job_id: int,
    dir_type: str,
    experiment_label: str,
    system_prompt: str,
    teacher_prompt: str,
    skip_graded: bool = True,
    rubric_dimensions: Optional[list] = None,
    score_levels: Optional[list] = None,
) -> None:
    job = ObeGradingJob.query.get(job_id)
    if job is None:
        raise ObeGradingError(f"job {job_id} 不存在")

    sign_picture_abs = resolve_signature_path(task_id, job.sign_picture_path)
    if not sign_picture_abs.exists():
        raise ObeGradingError(f"签名图不存在: {sign_picture_abs}")

    lo_profile_dir = str(job_lo_profile_dir(task_id, job_id))
    Path(lo_profile_dir).mkdir(parents=True, exist_ok=True)

    sign_ctx = SignContext(sign_picture_path=str(sign_picture_abs), sign_date_str=job.sign_date)
    llm_client = _make_llm_client()

    write_job_log(
        task_id,
        job_id,
        f"[{datetime.utcnow().isoformat()}] job {job_id} start: dir_type={dir_type}, "
        f"experiment={experiment_label}, skip_graded={skip_graded}\n",
        append=False,
    )

    query = ObeStudent.query.filter_by(
        task_id=task_id,
        dir_type=dir_type,
        experiment_label=experiment_label,
        matched=True,
    )
    if skip_graded:
        query = query.filter(ObeStudent.grade_status.in_(["pending", "failed"]))
    students = query.order_by(ObeStudent.student_id).all()

    graded = 0
    failed = 0

    for s in students:
        # 协作式取消：循环顶部检查标志位，跑完上一份后立即退出（当前那份会跑完）
        if _is_cancelled(job_id):
            break
        s.grade_status = "grading"
        # 进入批改中：清空旧结果（批改中无有效分数/评语），前端轮询能拿到一致的「批改中、无分数」状态
        s.last_score = None
        s.last_comment = None
        s.last_graded_file = None
        s.last_grade_error = None
        db.session.commit()

        _update_progress(
            job_id,
            currentStudent=f"{s.student_id} {s.student_name}",
            currentStudentId=s.student_id,
        )

        try:
            score, comment, graded_file = _grade_one_student(
                task_id=task_id,
                student=s,
                sign_ctx=sign_ctx,
                llm_client=llm_client,
                lo_profile_dir=lo_profile_dir,
                system_prompt=system_prompt,
                teacher_prompt=teacher_prompt,
                rubric_dimensions=rubric_dimensions,
                score_levels=score_levels,
            )

            if score > 0:
                s.grade_status = "graded"
                s.last_score = score
                s.last_comment = comment
                try:
                    rel = (
                        str(Path(graded_file).relative_to(task_root(task_id))).replace("\\", "/")
                        if graded_file
                        else None
                    )
                except ValueError:
                    rel = None
                s.last_graded_file = rel
                s.last_grade_error = None
                s.last_graded_at = datetime.utcnow()
                detail_status = "graded"
                err = None
                graded += 1
            else:
                s.grade_status = "failed"
                s.last_grade_error = "批改失败（score=0，可能 LLM 返回异常或文件格式不符）"
                s.last_graded_at = datetime.utcnow()
                detail_status = "failed"
                err = s.last_grade_error
                failed += 1

            db.session.add(
                ObeGradingJobDetail(
                    job_id=job_id,
                    student_id=s.id,
                    score=score if score > 0 else None,
                    comment=comment,
                    status=detail_status,
                    error_msg=err,
                    graded_file=s.last_graded_file,
                    finished_at=datetime.utcnow(),
                )
            )
            job.graded = graded
            job.failed = failed
            db.session.commit()

            _update_progress(job_id, graded=graded, failed=failed)

            write_job_log(
                task_id,
                job_id,
                f"[{datetime.utcnow().isoformat()}] {s.student_id} {s.student_name}: score={score}\n",
            )

        except Exception as e:
            db.session.rollback()
            s2 = ObeStudent.query.get(s.id)
            if s2:
                s2.grade_status = "failed"
                s2.last_grade_error = str(e)
                s2.last_graded_at = datetime.utcnow()
                db.session.add(
                    ObeGradingJobDetail(
                        job_id=job_id,
                        student_id=s2.id,
                        status="failed",
                        error_msg=str(e),
                        finished_at=datetime.utcnow(),
                    )
                )
                failed += 1
                job.graded = graded
                job.failed = failed
                db.session.commit()
                _update_progress(job_id, failed=failed)
                write_job_log(
                    task_id,
                    job_id,
                    f"[{datetime.utcnow().isoformat()}] {s.student_id} EXCEPTION: {e}\n",
                )

    if _is_cancelled(job_id):
        job.status = "cancelled"
        job.error_summary = "用户取消"
    else:
        job.status = "completed"
    job.finished_at = datetime.utcnow()
    # 保险：把可能的 grading 残留重置为 pending（取消后可继续批改）。
    # 正常 completed 时无 grading 残留，此 update 命中 0 行，无副作用。
    ObeStudent.query.filter_by(
        task_id=task_id,
        dir_type=dir_type,
        experiment_label=experiment_label,
        grade_status="grading",
    ).update(
        {"grade_status": "pending", "last_grade_error": "批改已取消"},
        synchronize_session=False,
    )
    ObeTask.query.filter_by(id=task_id).update({"status": "graded"})
    db.session.commit()

    # 把最终状态同步到内存进度缓存：取消场景下，worker 跑完当前学生后才真正停，
    # 此时 graded/failed 才是最终值。前端轮询到此（或随后 DB fallback）才停止，
    # 避免提前停止导致进度数字少 1。随后 worker finally 会 _clear_progress 清掉缓存。
    _update_progress(
        job_id, status=job.status, graded=graded, failed=failed, currentStudent=None
    )

    try:
        excel_path = build_excel_summary(task_id, dir_type, experiment_label, job_id)
        write_job_log(
            task_id,
            job_id,
            f"[{datetime.utcnow().isoformat()}] excel generated: {excel_path}\n",
        )
    except Exception as e:
        write_job_log(task_id, job_id, f"[excel] 生成失败: {e}\n")

    write_job_log(
        task_id,
        job_id,
        f"[{datetime.utcnow().isoformat()}] job {job_id} done: graded={graded}, failed={failed}\n",
    )


def _grade_one_student(
    task_id: int,
    student: ObeStudent,
    sign_ctx: SignContext,
    llm_client,
    lo_profile_dir: str,
    system_prompt: str,
    teacher_prompt: str,
    rubric_dimensions: Optional[list] = None,
    score_levels: Optional[list] = None,
) -> tuple[int, Optional[str], Optional[str]]:
    if not student.uploaded_file:
        return 0, None, None

    # uploaded_file 已经是学生目录下的路径（match 阶段已解压到学生目录）
    abs_file_path = (task_root(task_id) / student.uploaded_file).resolve()
    if not abs_file_path.exists():
        return 0, None, None

    # 原地批改（不再 shutil.copy，文件已经在学生目录下）
    score, comment, graded_file = grade_student_docx(
        docx_path=str(abs_file_path),
        system_prompt=system_prompt,
        teacher_prompt=teacher_prompt,
        llm_client=llm_client,
        sign_ctx=sign_ctx,
        lo_profile_dir=lo_profile_dir,
        rubric_dimensions=rubric_dimensions,
        score_levels=score_levels,
    )
    return score, comment, graded_file


def _mark_job_failed(
    task_id: int,
    job_id: int,
    dir_type: str,
    experiment_label: str,
    err_msg: str,
) -> None:
    try:
        job = ObeGradingJob.query.get(job_id)
        if job:
            job.status = "failed"
            job.error_summary = err_msg
            job.finished_at = datetime.utcnow()
        ObeStudent.query.filter_by(
            task_id=task_id,
            dir_type=dir_type,
            experiment_label=experiment_label,
            grade_status="grading",
        ).update(
            {"grade_status": "failed", "last_grade_error": "批改任务异常中断"},
            synchronize_session=False,
        )
        ObeTask.query.filter_by(id=task_id).update({"status": "graded"})
        db.session.commit()
    except Exception:
        db.session.rollback()


def request_cancel(task_id: int, job_id: int) -> dict:
    """请求取消正在运行的批改 job。

    协作式取消：只置标志位 + 更新内存进度状态；DB job.status 由 worker 在退出时
    统一写入，避免与 worker 竞争。worker 跑完当前学生后会在循环顶部检查到标志位并退出，
    所以取消不会立即生效（最多等一份批改完成，约 10-30s）。
    """
    job = ObeGradingJob.query.get(job_id)
    if job is None or job.task_id != task_id:
        raise ObeGradingError("批改任务不存在")
    if job.status != "running":
        raise ObeGradingError(f"任务当前状态为 {job.status}，无法取消")

    _CANCEL_FLAGS.setdefault(job_id, threading.Event()).set()
    # 注意：此处不把 progress.status 改成 cancelled。worker 是协作式取消，要跑完当前
    # 学生才真正停。若提前把 status 置为 cancelled，前端轮询会立即停止，丢失「当前学生」
    # 跑完后的 graded+1，导致显示的「已批改 x/y」比实际少 1。status 仍保持 running，
    # 由 worker 结束时统一写最终状态（cancelled/completed）。
    write_job_log(
        task_id,
        job_id,
        f"[{datetime.utcnow().isoformat()}] cancel requested by user\n",
    )
    return {"jobId": job_id, "status": "cancelling"}


def get_progress(
    task_id: int,
    dir_type: str,
    experiment_label: Optional[str] = None,
    job_id: Optional[int] = None,
) -> dict:
    """查询批改进度。优先读内存缓存，fallback DB 聚合。"""
    if job_id and job_id in _JOB_PROGRESS:
        prog = dict(_JOB_PROGRESS[job_id])
        prog["jobId"] = job_id
        return prog

    if not job_id:
        query = ObeGradingJob.query.filter_by(task_id=task_id, dir_type=dir_type)
        if experiment_label:
            query = query.filter_by(experiment_label=experiment_label)
        job = query.order_by(ObeGradingJob.id.desc()).first()
        if not job:
            return {"total": 0, "graded": 0, "failed": 0, "status": "none"}
        job_id = job.id

    prog = _JOB_PROGRESS.get(job_id)
    if prog:
        return {**prog, "jobId": job_id}

    job = ObeGradingJob.query.get(job_id)
    if not job:
        return {"total": 0, "graded": 0, "failed": 0, "status": "none"}
    return {
        "jobId": job.id,
        "status": job.status,
        "total": job.total,
        "graded": job.graded,
        "failed": job.failed,
        "currentStudent": None,
        "lastUpdate": job.finished_at.isoformat() if job.finished_at else None,
    }


def cleanup_zombie_grading() -> None:
    running_jobs = ObeGradingJob.query.filter_by(status="running").all()
    for job in running_jobs:
        job.status = "failed"
        job.error_summary = "进程重启，任务中断"
        job.finished_at = datetime.utcnow()

    ObeStudent.query.filter_by(grade_status="grading").update(
        {"grade_status": "failed", "last_grade_error": "进程重启，任务中断"},
        synchronize_session=False,
    )

    ObeTask.query.filter_by(status="grading").update({"status": "graded"}, synchronize_session=False)

    if running_jobs:
        db.session.commit()
        print(f"[cleanup] 清理 {len(running_jobs)} 个僵尸 job")


def ensure_obe_rubric_schema() -> None:
    """幂等：建 obe_grading_rubric 表 + 给 obe_grading_job 加列。

    项目无 alembic（靠 db.create_all()），但 create_all 不会改已存在表，
    故新表用 __table__.create(checkfirst=True)，加列用 inspect 检测后 ALTER。
    开发/生产启动自动跑，失败不阻塞 app。
    """
    from sqlalchemy import inspect, text

    from models.obe import ObeGradingRubric

    ObeGradingRubric.__table__.create(db.engine, checkfirst=True)

    insp = inspect(db.engine)
    cols = {c["name"] for c in insp.get_columns("obe_grading_job")}
    with db.engine.begin() as conn:
        if "rubric_dimensions" not in cols:
            conn.execute(text("ALTER TABLE obe_grading_job ADD COLUMN rubric_dimensions JSON NULL"))
        if "score_levels" not in cols:
            conn.execute(text("ALTER TABLE obe_grading_job ADD COLUMN score_levels JSON NULL"))


def upsert_rubric(
    user_id: str,
    course_name: str,
    dimensions: list,
    free_text: str,
    score_levels: Optional[list] = None,
):
    """按 (user_id, course_name) upsert 评分标准，返回 ObeGradingRubric 记录。"""
    from models.obe import ObeGradingRubric

    rec = ObeGradingRubric.query.filter_by(user_id=user_id, course_name=course_name).first()
    if rec is None:
        rec = ObeGradingRubric(user_id=user_id, course_name=course_name)
        db.session.add(rec)
    rec.dimensions = dimensions
    rec.free_text = free_text
    rec.score_levels = score_levels
    db.session.commit()
    return rec


def get_rubric(user_id: str, course_name: str):
    """读取该教师该课程的最近评分标准，无则返回 None。"""
    from models.obe import ObeGradingRubric

    return ObeGradingRubric.query.filter_by(user_id=user_id, course_name=course_name).first()


def build_excel_summary(
    task_id: int,
    dir_type: str,
    experiment_label: str,
    job_id: int,
) -> Path:
    rows = []
    students = (
        ObeStudent.query.filter_by(
            task_id=task_id, dir_type=dir_type, experiment_label=experiment_label
        )
        .order_by(ObeStudent.student_id)
        .all()
    )
    for s in students:
        rows.append(
            {
                "学号": s.student_id,
                "姓名": s.student_name,
                "班级": s.student_class,
                "上传文件": Path(s.uploaded_file).name if s.uploaded_file else "",
                "是否上传": "是" if s.matched else "否",
                "批改状态": s.grade_status,
                "分数": s.last_score if s.last_score is not None else "",
                "评语": s.last_comment or "",
                "错误": s.last_grade_error or "",
            }
        )

    df = pd.DataFrame(rows)
    excel_root = excel_dir(task_id)
    excel_root.mkdir(parents=True, exist_ok=True)
    safe_exp = _safe_segment(experiment_label)
    out_path = excel_root / f"{safe_exp}_{job_id}_成绩.xlsx"
    df.to_excel(out_path, index=False)
    return out_path


def retry_student(
    task_id: int,
    student_pk: int,
    system_prompt: Optional[str] = None,
) -> dict:
    """同步重试单个学生的批改（复用最近一次同 experiment 的 job 签名参数）。"""
    student = ObeStudent.query.get(student_pk)
    if student is None or student.task_id != task_id:
        raise ObeGradingError("学生不存在或不属于该任务")

    if not student.matched or not student.uploaded_file:
        raise ObeGradingError("该学生未上传文件，无法重试")

    job = (
        ObeGradingJob.query.filter_by(
            task_id=task_id,
            dir_type=student.dir_type,
            experiment_label=student.experiment_label,
        )
        .order_by(ObeGradingJob.id.desc())
        .first()
    )
    if job is None:
        raise ObeGradingError("未找到该实验的批改任务历史，无法重试（请先触发一次完整批改）")

    sign_picture_abs = resolve_signature_path(task_id, job.sign_picture_path)
    if not sign_picture_abs.exists():
        raise ObeGradingError(f"签名图丢失: {sign_picture_abs}")

    lo_profile_dir = str(job_lo_profile_dir(task_id, job.id))
    Path(lo_profile_dir).mkdir(parents=True, exist_ok=True)

    sign_ctx = SignContext(
        sign_picture_path=str(sign_picture_abs), sign_date_str=job.sign_date
    )
    llm_client = _make_llm_client()
    sys_prompt = system_prompt or DEFAULT_SYSTEM_PROMPT

    student.grade_status = "grading"
    # 进入批改中：清空旧结果（批改中无有效分数/评语）
    student.last_score = None
    student.last_comment = None
    student.last_graded_file = None
    student.last_grade_error = None
    db.session.commit()

    score, comment, graded_file = _grade_one_student(
        task_id=task_id,
        student=student,
        sign_ctx=sign_ctx,
        llm_client=llm_client,
        lo_profile_dir=lo_profile_dir,
        system_prompt=sys_prompt,
        teacher_prompt=job.teacher_prompt,
        rubric_dimensions=job.rubric_dimensions,
        score_levels=job.score_levels,
    )

    if score > 0:
        student.grade_status = "graded"
        student.last_score = score
        student.last_comment = comment
        try:
            rel = (
                str(Path(graded_file).relative_to(task_root(task_id))).replace("\\", "/")
                if graded_file
                else None
            )
        except ValueError:
            rel = None
        student.last_graded_file = rel
        student.last_grade_error = None
        student.last_graded_at = datetime.utcnow()
        detail_status = "graded"
        err = None
    else:
        student.grade_status = "failed"
        student.last_grade_error = "重试失败（score=0）"
        student.last_graded_at = datetime.utcnow()
        detail_status = "failed"
        err = student.last_grade_error

    db.session.add(
        ObeGradingJobDetail(
            job_id=job.id,
            student_id=student.id,
            score=score if score > 0 else None,
            comment=comment,
            status=detail_status,
            error_msg=err,
            graded_file=student.last_graded_file,
            finished_at=datetime.utcnow(),
        )
    )
    db.session.commit()

    return student.to_dict()
