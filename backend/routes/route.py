from flask import Blueprint, request

from services.route_service import list_user_routes, list_constant_routes, is_route_exist
from utils.response import success, fail
from utils.jwt_helper import jwt_required

route_bp = Blueprint("route", __name__)


@route_bp.get("/getUserRoutes")
@jwt_required
def user_routes():
    return success(list_user_routes())


@route_bp.get("/getConstantRoutes")
def constant_routes():
    return success(list_constant_routes())


@route_bp.get("/isRouteExist")
@jwt_required
def route_exist():
    name = request.args.get("routeName", "").strip()
    if not name:
        return fail("routeName 必填")
    return success(is_route_exist(name))
