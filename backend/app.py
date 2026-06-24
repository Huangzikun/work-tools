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

    return app


if __name__ == "__main__":
    app = create_app()
    app.run(host="0.0.0.0", port=5001, debug=True)
