from flask import jsonify

# 业务码（对齐 soybean-admin 前端拦截器约定，见 web/.env）
CODE_SUCCESS = "0000"          # 成功
CODE_BUSINESS = "1000"         # 普通业务错误（前端弹出 msg）
CODE_TOKEN_INVALID = "9998"    # token 无效
CODE_TOKEN_EXPIRED = "9999"    # token 过期（前端自动 refreshToken）
CODE_LOGOUT = "8888"           # 强制登出


def success(data=None, msg="ok", code=CODE_SUCCESS):
    return jsonify({"code": code, "msg": msg, "data": data})


def fail(msg="error", code=CODE_BUSINESS, data=None):
    return jsonify({"code": code, "msg": msg, "data": data})
