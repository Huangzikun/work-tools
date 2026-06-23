HOME_NAME = "home"

USER_ROUTES = [
    {
        "name": "home",
        "path": "/home",
        "component": "layout.base$view.home",
        "meta": {
            "title": "首页",
            "i18nKey": "route.home",
            "icon": "mdi:monitor-dashboard",
            "order": 1,
        },
    },
    {
        "name": "user-center",
        "path": "/user-center",
        "component": "layout.base$view.user-center",
        "meta": {
            "title": "个人中心",
            "i18nKey": "route.user-center",
            "icon": "mdi:user",
            "order": 2,
        },
    },
]

CONSTANT_ROUTES = []
