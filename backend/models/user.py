from datetime import datetime

from extensions import db


class User(db.Model):
    __tablename__ = "sys_user"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    user_id = db.Column(db.String(64), unique=True, nullable=False, comment="业务用户ID")
    user_name = db.Column(db.String(64), unique=True, nullable=False, comment="登录用户名")
    password = db.Column(db.String(255), nullable=False, comment="密码哈希")
    roles = db.Column(db.JSON, default=list, comment="角色编码列表")
    buttons = db.Column(db.JSON, default=list, comment="按钮权限码列表")
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(
        db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False
    )

    def to_user_info(self) -> dict:
        return {
            "userId": self.user_id,
            "userName": self.user_name,
            "roles": self.roles or [],
            "buttons": self.buttons or [],
        }
