from extensions import db
from models.user import User
from utils.password import verify_password
from utils.jwt_helper import encode_token
from utils.response import CODE_BUSINESS, CODE_LOGOUT


class AuthError(Exception):
    def __init__(self, message: str, code: str = CODE_BUSINESS):
        super().__init__(message)
        self.message = message
        self.code = code


def login(user_name: str, password: str) -> dict:
    user = User.query.filter_by(user_name=user_name).first()
    if user is None or not verify_password(password, user.password):
        # 不暴露用户存在性，统一错误信息
        raise AuthError("用户名或密码错误", code=CODE_BUSINESS)
    token = encode_token(user.user_id, user.user_name, "access")
    refresh = encode_token(user.user_id, user.user_name, "refresh")
    return {"token": token, "refreshToken": refresh}


def refresh_token(refresh_jwt: str) -> dict:
    from utils.jwt_helper import decode_token
    try:
        payload = decode_token(refresh_jwt)
    except Exception as exc:
        # refreshToken 失败必须用登出码，避免前端死循环 refresh
        raise AuthError("refreshToken 无效或已过期", code=CODE_LOGOUT) from exc
    if payload.get("type") != "refresh":
        raise AuthError("token 类型错误", code=CODE_LOGOUT)
    user = User.query.filter_by(user_id=payload.get("userId")).first()
    if user is None:
        raise AuthError("用户不存在", code=CODE_LOGOUT)
    token = encode_token(user.user_id, user.user_name, "access")
    new_refresh = encode_token(user.user_id, user.user_name, "refresh")
    return {"token": token, "refreshToken": new_refresh}


def get_user_info(user_id: str) -> dict:
    user = User.query.filter_by(user_id=user_id).first()
    if user is None:
        raise AuthError("用户不存在", code=CODE_BUSINESS)
    return user.to_user_info()
