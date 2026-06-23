from seed.menus import USER_ROUTES, CONSTANT_ROUTES, HOME_NAME


def list_user_routes() -> dict:
    return {"routes": USER_ROUTES, "home": HOME_NAME}


def list_constant_routes() -> list:
    return CONSTANT_ROUTES


def is_route_exist(route_name: str) -> bool:
    names = {item["name"] for item in USER_ROUTES}
    return route_name in names
