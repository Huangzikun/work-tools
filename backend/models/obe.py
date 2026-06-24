from datetime import datetime

from extensions import db


class ObeTask(db.Model):
    """OBE 目录任务：用户上传名单后创建，持久化目录结构与生成历史。"""

    __tablename__ = "obe_task"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    user_id = db.Column(db.String(64), nullable=False, index=True, comment="创建者 sys_user.user_id")
    class_name = db.Column(db.String(128), nullable=False, comment="班级名")
    course_name = db.Column(db.String(128), nullable=False, comment="课程名")
    teacher_name = db.Column(db.String(64), nullable=False, comment="教师姓名")
    fixed_dir_types = db.Column(db.JSON, default=list, comment="固定目录类型列表")
    student_dir_types = db.Column(db.JSON, default=list, comment="考核目录类型列表")
    roster_file_name = db.Column(db.String(255), nullable=False, comment="原始名单文件名")
    student_count = db.Column(db.Integer, nullable=False, default=0, comment="名单学生数")
    storage_path = db.Column(db.String(512), nullable=False, comment="相对 OBE_STORAGE_ROOT 的任务根路径")
    status = db.Column(
        db.String(16),
        nullable=False,
        default="created",
        comment="created/uploading/ready/grading/graded",
    )
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(
        db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False
    )

    students = db.relationship(
        "ObeStudent",
        backref="task",
        cascade="all, delete-orphan",
        passive_deletes=True,
    )
    jobs = db.relationship(
        "ObeGradingJob",
        backref="task",
        cascade="all, delete-orphan",
        passive_deletes=True,
    )

    __table_args__ = (
        db.Index("idx_obe_task_user_created", "user_id", "created_at"),
    )

    def to_summary(self) -> dict:
        return {
            "id": self.id,
            "className": self.class_name,
            "courseName": self.course_name,
            "teacherName": self.teacher_name,
            "fixedDirTypes": self.fixed_dir_types or [],
            "studentDirTypes": self.student_dir_types or [],
            "studentCount": self.student_count,
            "status": self.status,
            "createdAt": self.created_at.isoformat() if self.created_at else None,
        }


class ObeStudent(db.Model):
    """某个任务下、某个目录类型 + 实验维度中的学生记录。

    一个学生在 (dir_type, experiment_label) 下有一条记录：
    - dir_type = "实验实训报告" / "课程考核" 等（mkdir 时配置）
    - experiment_label = "图书管理系统-管理员端" / "图书管理系统-用户端" 等（上传 ZIP 时自动按文件名生成）
    一个学生在不同 experiment 下有独立的上传文件和批改结果。
    """

    __tablename__ = "obe_student"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    task_id = db.Column(
        db.Integer,
        db.ForeignKey("obe_task.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    dir_type = db.Column(db.String(128), nullable=False, comment="所属目录类型")
    experiment_label = db.Column(
        db.String(128),
        nullable=False,
        default="default",
        comment="实验名（每个上传的 ZIP 自动成为一个 experiment）",
    )
    student_id = db.Column(db.String(32), nullable=False, comment="学号")
    student_name = db.Column(db.String(64), nullable=False, comment="姓名")
    student_class = db.Column(db.String(128), comment="行政班级")
    student_dir_name = db.Column(db.String(256), nullable=False, comment="学生子目录名")
    uploaded_file = db.Column(db.String(512), comment="相对任务根的上传原文件路径")
    matched = db.Column(db.Boolean, nullable=False, default=False, comment="是否已匹配到上传文件")
    grade_status = db.Column(
        db.String(16),
        nullable=False,
        default="pending",
        comment="pending/grading/graded/failed",
    )
    last_score = db.Column(db.Integer, comment="最新批改分数")
    last_comment = db.Column(db.Text, comment="最新评语")
    last_graded_file = db.Column(db.String(512), comment="最新批改后 docx 路径")
    last_grade_error = db.Column(db.Text, comment="最新批改错误信息")
    last_graded_at = db.Column(db.DateTime, comment="最新批改完成时间")
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    job_details = db.relationship(
        "ObeGradingJobDetail",
        backref="student",
        cascade="all, delete-orphan",
        passive_deletes=True,
    )

    __table_args__ = (
        db.UniqueConstraint(
            "task_id",
            "dir_type",
            "experiment_label",
            "student_id",
            name="uq_obe_student_task_dir_exp_student",
        ),
        db.Index(
            "idx_obe_student_task_dir_exp_status",
            "task_id",
            "dir_type",
            "experiment_label",
            "grade_status",
        ),
    )

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "taskId": self.task_id,
            "dirType": self.dir_type,
            "experimentLabel": self.experiment_label,
            "studentId": self.student_id,
            "studentName": self.student_name,
            "studentClass": self.student_class,
            "studentDirName": self.student_dir_name,
            "uploadedFile": self.uploaded_file,
            "matched": self.matched,
            "gradeStatus": self.grade_status,
            "lastScore": self.last_score,
            "lastComment": self.last_comment,
            "lastGradedFile": self.last_graded_file,
            "lastGradeError": self.last_grade_error,
            "lastGradedAt": self.last_graded_at.isoformat() if self.last_graded_at else None,
        }


class ObeGradingJob(db.Model):
    """一次批改任务（用户点「开始批改」产生的一条记录）。"""

    __tablename__ = "obe_grading_job"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    task_id = db.Column(
        db.Integer,
        db.ForeignKey("obe_task.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    dir_type = db.Column(db.String(128), nullable=False, comment="批改的目录类型")
    experiment_label = db.Column(
        db.String(128),
        nullable=False,
        default="default",
        comment="实验名",
    )
    teacher_name = db.Column(db.String(64), nullable=False, comment="教师姓名")
    sign_date = db.Column(db.String(32), nullable=False, comment="签名日期字符串")
    sign_picture_path = db.Column(db.String(512), nullable=False, comment="签名图相对路径")
    teacher_prompt = db.Column(db.Text, nullable=False, comment="教师要求文本")
    status = db.Column(
        db.String(16),
        nullable=False,
        default="running",
        comment="running/completed/failed/cancelled",
    )
    total = db.Column(db.Integer, nullable=False, default=0)
    graded = db.Column(db.Integer, nullable=False, default=0)
    failed = db.Column(db.Integer, nullable=False, default=0)
    started_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    finished_at = db.Column(db.DateTime)
    error_summary = db.Column(db.Text)

    details = db.relationship(
        "ObeGradingJobDetail",
        backref="job",
        cascade="all, delete-orphan",
        passive_deletes=True,
    )

    __table_args__ = (
        db.Index("idx_obe_job_task_dir_exp", "task_id", "dir_type", "experiment_label"),
    )

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "taskId": self.task_id,
            "dirType": self.dir_type,
            "experimentLabel": self.experiment_label,
            "teacherName": self.teacher_name,
            "signDate": self.sign_date,
            "signPicturePath": self.sign_picture_path,
            "teacherPrompt": self.teacher_prompt,
            "status": self.status,
            "total": self.total,
            "graded": self.graded,
            "failed": self.failed,
            "startedAt": self.started_at.isoformat() if self.started_at else None,
            "finishedAt": self.finished_at.isoformat() if self.finished_at else None,
            "errorSummary": self.error_summary,
        }


class ObeGradingJobDetail(db.Model):
    """单次批改任务下每个学生的结果（历史保留，用于审计与回溯）。"""

    __tablename__ = "obe_grading_job_detail"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    job_id = db.Column(
        db.Integer,
        db.ForeignKey("obe_grading_job.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    student_id = db.Column(
        db.Integer,
        db.ForeignKey("obe_student.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    score = db.Column(db.Integer)
    comment = db.Column(db.Text)
    status = db.Column(db.String(16), nullable=False, comment="graded/failed")
    error_msg = db.Column(db.Text)
    graded_file = db.Column(db.String(512))
    finished_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
