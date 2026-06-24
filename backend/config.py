import os
from dotenv import load_dotenv

load_dotenv()


class Config:
    SECRET_KEY = os.getenv("FLASK_SECRET_KEY", "dev-secret-change-me")

    SQLALCHEMY_DATABASE_URI = os.getenv(
        "DATABASE_URI",
        "mysql+pymysql://root:root123@localhost:3306/teacher_recruitment?charset=utf8mb4",
    )
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SQLALCHEMY_ENGINE_OPTIONS = {"pool_pre_ping": True, "pool_recycle": 3600}

    JWT_SECRET_KEY = os.getenv("JWT_SECRET_KEY", "jwt-dev-secret-change-me")
    JWT_ACCESS_TTL = int(os.getenv("JWT_ACCESS_TTL", "86400"))
    JWT_REFRESH_TTL = int(os.getenv("JWT_REFRESH_TTL", "604800"))

    _cors_default = "http://localhost:9530,http://127.0.0.1:9530"
    CORS_ORIGINS = [
        origin.strip()
        for origin in os.getenv("CORS_ORIGINS", _cors_default).split(",")
        if origin.strip()
    ]

    DEFAULT_ADMIN = {
        "user_id": "1",
        "user_name": "admin",
        "password": "123456",
        "roles": ["R_ADMIN"],
        "buttons": [
            "bAuth:btn:add",
            "bAuth:btn:edit",
            "bAuth:btn:delete",
        ],
    }

    DB_ROOT = {
        "host": os.getenv("DB_HOST", "localhost"),
        "port": int(os.getenv("DB_PORT", "3306")),
        "user": os.getenv("DB_USER", "root"),
        "password": os.getenv("DB_PASSWORD", "root123"),
        "dbname": os.getenv("DB_NAME", "teacher_recruitment"),
    }
