"""一键初始化数据库：建库 → 建表 → 写入默认 admin。

使用：
    conda activate teacherrecruitment
    python init_db.py
"""

import sys
from pathlib import Path

import pymysql

sys.path.insert(0, str(Path(__file__).parent))

from config import Config


def ensure_database() -> None:
    root = Config.DB_ROOT
    conn = pymysql.connect(
        host=root["host"],
        port=root["port"],
        user=root["user"],
        password=root["password"],
        charset="utf8mb4",
    )
    try:
        with conn.cursor() as cur:
            cur.execute(
                f"CREATE DATABASE IF NOT EXISTS `{root['dbname']}` "
                "DEFAULT CHARACTER SET utf8mb4 DEFAULT COLLATE utf8mb4_unicode_ci"
            )
        conn.commit()
    finally:
        conn.close()
    print(f"[init] database `{root['dbname']}` ready")


def ensure_tables_and_seed() -> None:
    from app import create_app
    from extensions import db
    # 显式 import 所有 Model，确保 db.create_all() 能感知到全部表
    from models import (  # noqa: F401
        LessonPlanTask,
        ObeGradingJob,
        ObeGradingJobDetail,
        ObeGradingRubric,
        ObeStudent,
        ObeTask,
        User,
    )
    from utils.password import hash_password

    app = create_app()
    with app.app_context():
        db.create_all()
        admin_cfg = Config.DEFAULT_ADMIN
        existing = User.query.filter_by(user_name=admin_cfg["user_name"]).first()
        if existing is not None:
            print(f"[init] admin `{admin_cfg['user_name']}` already exists, skip seed")
            return
        admin = User(
            user_id=admin_cfg["user_id"],
            user_name=admin_cfg["user_name"],
            password=hash_password(admin_cfg["password"]),
            roles=admin_cfg["roles"],
            buttons=admin_cfg["buttons"],
        )
        db.session.add(admin)
        db.session.commit()
        print(
            f"[init] admin `{admin_cfg['user_name']}` / `{admin_cfg['password']}` seeded"
        )


if __name__ == "__main__":
    ensure_database()
    ensure_tables_and_seed()
    print("[init] done")
