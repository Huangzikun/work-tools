from flask import Blueprint, request, g

from services.auth_service import login, refresh_token, get_user_info, AuthError
from utils.response import success, fail
from utils.jwt_helper import jwt_required

auth_bp = Blueprint("auth", __name__)


@auth_bp.post("/login")
def do_login():
    data = request.get_json(silent=True) or {}
    user_name = (data.get("userName") or "").strip()
    password = data.get("password") or ""
    if not user_name or not password:
        return fail("userName 与 password 必填")
    try:
        result = login(user_name, password)
    except AuthError as exc:
        return fail(exc.message, code=exc.code)
    return success(result)


@auth_bp.post("/refreshToken")
def do_refresh():
    data = request.get_json(silent=True) or {}
    refresh = data.get("refreshToken")
    if not refresh:
        return fail("refreshToken 必填")
    try:
        result = refresh_token(refresh)
    except AuthError as exc:
        return fail(exc.message, code=exc.code)
    return success(result)


@auth_bp.get("/getUserInfo")
@jwt_required
def do_user_info():
    try:
        info = get_user_info(g.current_user_id)
    except AuthError as exc:
        return fail(exc.message, code=exc.code)
    return success(info)
