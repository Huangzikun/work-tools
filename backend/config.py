import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent


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

    # OBE 任务存储根目录（持久化的目录树、上传文件、签名图、Excel 都在这下面）
    OBE_STORAGE_ROOT = os.getenv(
        "OBE_STORAGE_ROOT", str(BASE_DIR / "storage" / "obe")
    )

    # 单个 OBE 任务上传总大小上限（学生文件批量上传），默认 200MB
    OBE_UPLOAD_MAX_BYTES = int(os.getenv("OBE_UPLOAD_MAX_BYTES", str(200 * 1024 * 1024)))

    # 火山引擎 Ark（AI 批阅）
    ARK_API_KEY = os.getenv("ARK_API_KEY", "")
    ARK_BASE_URL = os.getenv(
        "ARK_BASE_URL", "https://ark.cn-beijing.volces.com/api/v3"
    )
    ARK_MODEL = os.getenv("ARK_MODEL", "doubao-seed-2-0-mini-260428")
