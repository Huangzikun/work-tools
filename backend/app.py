import sys
from pathlib import Path

# 把项目根加到 sys.path，让 backend 能 import common.llm_client（v6 共享客户端）
_PROJECT_ROOT = str(Path(__file__).resolve().parents[1])
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from flask import Flask

from config import Config
from extensions import db, migrate, cors
from routes import register_routes


def create_app(config_class=Config):
    app = Flask(__name__)
    app.config.from_object(config_class)

    db.init_app(app)
    migrate.init_app(app, db)
    cors.init_app(
        app,
        resources={r"/api/*": {"origins": app.config["CORS_ORIGINS"]}},
        supports_credentials=True,
    )

    register_routes(app)

    # 启动时清理上次进程遗留的 running job / grading student（daemon thread 被强杀后状态未收敛）
    from services.obe_grading import cleanup_zombie_grading

    try:
        with app.app_context():
            cleanup_zombie_grading()
    except Exception as exc:  # 启动期清理失败不应阻塞 app
        app.logger.exception("cleanup_zombie_grading failed: %s", exc)

    # 幂等建 obe_grading_rubric 表 + 给 obe_grading_job 加列（项目无 alembic，靠启动钩子）
    from services.obe_grading import ensure_obe_rubric_schema

    try:
        with app.app_context():
            ensure_obe_rubric_schema()
    except Exception as exc:
        app.logger.exception("ensure_obe_rubric_schema failed: %s", exc)

    # 清理教案生成僵尸任务
    from services.lesson_plan_service import cleanup_zombie_tasks

    try:
        with app.app_context():
            cleanup_zombie_tasks()
    except Exception as exc:
        app.logger.exception("cleanup_zombie_tasks failed: %s", exc)

    return app


if __name__ == "__main__":
    app = create_app()
    app.run(host="0.0.0.0", port=5001, debug=True)
