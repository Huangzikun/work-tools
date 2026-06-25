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


def _update_progress(job_id: int, **fields) -> None:
    with _PROGRESS_LOCK:
        prog = _JOB_PROGRESS.setdefault(
            job_id,
            {"total": 0, "graded": 0, "failed": 0, "currentStudent": None},
        )
        prog.update(fields)
        prog["lastUpdate"] = datetime.utcnow().isoformat()


def _clear_progress(job_id: int) -> None:
    with _PROGRESS_LOCK:
        _JOB_PROGRESS.pop(job_id, None)


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
    _update_progress(job_id, total=matched_count, graded=0, failed=0, currentStudent=None)

    app = current_app._get_current_object()
    sys_prompt = system_prompt or DEFAULT_SYSTEM_PROMPT

    def worker():
        with app.app_context():
            try:
                _run_grading(
                    task_id, job_id, dir_type, experiment_label, sys_prompt, teacher_prompt, skip_graded
                )
            except Exception as exc:
                app.logger.exception("grading job %s failed", job_id)
                _mark_job_failed(task_id, job_id, dir_type, experiment_label, str(exc))
            finally:
                db.session.remove()
                _clear_progress(job_id)
                _GRADED_WORKERS.pop(job_id, None)

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
        s.grade_status = "grading"
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

    job.status = "completed"
    job.finished_at = datetime.utcnow()
    ObeTask.query.filter_by(id=task_id).update({"status": "graded"})
    db.session.commit()

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
    db.session.commit()

    score, comment, graded_file = _grade_one_student(
        task_id=task_id,
        student=student,
        sign_ctx=sign_ctx,
        llm_client=llm_client,
        lo_profile_dir=lo_profile_dir,
        system_prompt=sys_prompt,
        teacher_prompt=job.teacher_prompt,
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
