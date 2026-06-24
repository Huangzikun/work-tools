import json
from urllib.parse import quote

from flask import Blueprint, current_app, g, make_response, request

from services.obe_mkdir_service import (
    MAX_UPLOAD_BYTES,
    ObeMkdirError,
    generate_zip,
)
from utils.jwt_helper import jwt_required
from utils.response import fail, success

obe_bp = Blueprint("obe", __name__)


@obe_bp.get("/ping")
@jwt_required
def ping():
    return success({"pong": True, "user": g.current_user_name})


def _parse_dir_types(raw: str, field_name: str) -> list:
    """把表单里的 JSON 数组字符串解析为 list[str]，并去空白。"""
    if not raw:
        return []
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        raise ObeMkdirError(f"{field_name} 不是合法的 JSON 数组")
    if not isinstance(parsed, list):
        raise ObeMkdirError(f"{field_name} 不是合法的 JSON 数组")
    result = []
    for item in parsed:
        if not isinstance(item, str):
            raise ObeMkdirError(f"{field_name} 数组元素必须为字符串")
        s = item.strip()
        if s:
            result.append(s)
    return result


@obe_bp.post("/mkdir")
@jwt_required
def obe_mkdir():
    # 1. 表单字段（multipart/form-data，字段名 camelCase）
    class_name = (request.form.get("className") or "").strip()
    course_name = (request.form.get("courseName") or "").strip()
    teacher_name = (request.form.get("teacherName") or "").strip()

    if not class_name:
        return fail("className 必填")
    if not course_name:
        return fail("courseName 必填")
    if not teacher_name:
        return fail("teacherName 必填")

    # 2. 大小预检（避免读到一半才发现超限）
    if request.content_length and request.content_length > MAX_UPLOAD_BYTES * 2:
        return fail(f"上传内容过大（超过 {MAX_UPLOAD_BYTES // (1024 * 1024)}MB 限制）")

    # 3. 文件
    if "roster" not in request.files:
        return fail("roster 文件必填")
    file_storage = request.files["roster"]

    # 4. 解析目录类型数组（JSON 字符串）
    try:
        fixed_dir_types = _parse_dir_types(request.form.get("fixedDirTypes") or "[]", "fixedDirTypes")
        student_dir_types = _parse_dir_types(
            request.form.get("studentDirTypes") or "[]", "studentDirTypes"
        )
    except ObeMkdirError as exc:
        return fail(exc.message)

    if not student_dir_types and not fixed_dir_types:
        return fail("至少需要选择一个固定目录或考核目录")

    # 5. 生成 ZIP
    try:
        zip_data, zip_name = generate_zip(
            file_storage,
            class_name,
            course_name,
            teacher_name,
            fixed_dir_types,
            student_dir_types,
        )
    except ObeMkdirError as exc:
        return fail(exc.message)
    except Exception as exc:
        current_app.logger.exception("obe_mkdir failed: %s", exc)
        return fail("服务异常，请稍后重试")

    # 6. 流式返回（中文文件名用 RFC 5987 编码）
    resp = make_response(zip_data)
    resp.headers["Content-Type"] = "application/zip"
    resp.headers["Content-Disposition"] = (
        f"attachment; filename*=UTF-8''{quote(zip_name)}"
    )
    resp.headers["Content-Length"] = str(len(zip_data))
    return resp
