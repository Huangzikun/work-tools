import time
import jwt
from flask import request, g, current_app

from utils.response import fail, CODE_TOKEN_INVALID, CODE_TOKEN_EXPIRED


def _secret():
    return current_app.config["JWT_SECRET_KEY"]


def encode_token(user_id: str, user_name: str, token_type: str = "access") -> str:
    now = int(time.time())
    ttl = (
        current_app.config["JWT_ACCESS_TTL"]
        if token_type == "access"
        else current_app.config["JWT_REFRESH_TTL"]
    )
    payload = {
        "userId": user_id,
        "userName": user_name,
        "type": token_type,
        "iat": now,
        "exp": now + ttl,
    }
    return jwt.encode(payload, _secret(), algorithm="HS256")


def decode_token(token: str) -> dict:
    return jwt.decode(token, _secret(), algorithms=["HS256"])


def jwt_required(fn):
    def wrapper(*args, **kwargs):
        auth = request.headers.get("Authorization", "")
        if not auth.startswith("Bearer "):
            return fail("未授权：缺少 Authorization 头", code=CODE_TOKEN_INVALID)
        token = auth[7:].strip()
        try:
            payload = decode_token(token)
        except jwt.ExpiredSignatureError:
            return fail("token 已过期", code=CODE_TOKEN_EXPIRED)
        except jwt.PyJWTError:
            return fail("token 无效", code=CODE_TOKEN_INVALID)
        if payload.get("type") != "access":
            return fail("token 类型错误", code=CODE_TOKEN_INVALID)
        g.current_user_id = payload.get("userId")
        g.current_user_name = payload.get("userName")
        return fn(*args, **kwargs)

    wrapper.__name__ = fn.__name__
    return wrapper
