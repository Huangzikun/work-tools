from datetime import datetime

from extensions import db


class LessonPlanTask(db.Model):
    """教案生成任务：用户上传大纲 + 填表单 → 后台异步生成 docx 教案。"""

    __tablename__ = "lesson_plan_task"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    user_id = db.Column(db.String(64), nullable=False, index=True, comment="创建者 sys_user.user_id")
    syllabus_name = db.Column(db.String(255), nullable=False, comment="原始大纲文件名")
    total_lessons = db.Column(db.Integer, nullable=False, comment="教案数（要生成的教案份数，每份对应一次课）")
    batch_size = db.Column(db.Integer, nullable=False, default=2, comment="每批生成的教案数")
    course_info = db.Column(db.JSON, nullable=False, comment="课程基本信息（首页表格 0 字段）")
    teacher_info = db.Column(db.JSON, nullable=False, comment="教师信息（首页表格 1 字段）")
    system_prompt = db.Column(db.Text, comment="自定义 system prompt（高级）")
    user_prompt = db.Column(db.Text, comment="自定义用户提示词补充要求（追加到 user prompt 末尾）")
    total_hours = db.Column(db.Integer, comment="课程总课时数（用于按课时分配每教案时间）")
    status = db.Column(
        db.String(32),
        nullable=False,
        default="pending",
        comment="pending/parsing/generating/building/completed/failed",
    )
    progress_total = db.Column(db.Integer, nullable=False, default=0)
    progress_done = db.Column(db.Integer, nullable=False, default=0)
    progress_label = db.Column(db.String(255), comment="当前批次描述")
    output_file = db.Column(db.String(512), comment="相对任务根的输出 docx 路径")
    error_summary = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    started_at = db.Column(db.DateTime)
    finished_at = db.Column(db.DateTime)

    __table_args__ = (
        db.Index("idx_lp_task_user_created", "user_id", "created_at"),
    )

    def to_summary(self) -> dict:
        return {
            "id": self.id,
            "syllabusName": self.syllabus_name,
            "totalLessons": self.total_lessons,
            "batchSize": self.batch_size,
            "courseInfo": self.course_info or {},
            "teacherInfo": self.teacher_info or {},
            "status": self.status,
            "progressTotal": self.progress_total,
            "progressDone": self.progress_done,
            "progressLabel": self.progress_label,
            "outputFile": self.output_file,
            "errorSummary": self.error_summary,
            "createdAt": self.created_at.isoformat() if self.created_at else None,
            "startedAt": self.started_at.isoformat() if self.started_at else None,
            "finishedAt": self.finished_at.isoformat() if self.finished_at else None,
        }
