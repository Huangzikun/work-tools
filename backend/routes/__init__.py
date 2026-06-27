from flask import Flask

from .auth import auth_bp
from .route import route_bp
from .obe import obe_bp
from .lesson_plan import lesson_plan_bp


def register_routes(app: Flask) -> None:
    app.register_blueprint(auth_bp, url_prefix="/api/auth")
    app.register_blueprint(route_bp, url_prefix="/api/route")
    app.register_blueprint(obe_bp, url_prefix="/api/obe")
    app.register_blueprint(lesson_plan_bp, url_prefix="/api/lesson-plan")
