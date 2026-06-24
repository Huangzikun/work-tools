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
    {
        "name": "obe",
        "path": "/obe",
        "component": "layout.base",
        "meta": {
            "title": "OBE工具",
            "i18nKey": "route.obe",
            "icon": "mdi:toolbox-outline",
            "order": 10,
        },
        "children": [
            {
                "name": "obe_mkdir",
                "path": "/obe/mkdir",
                "component": "layout.base$view.obe_mkdir",
                "meta": {
                    "title": "目录生成",
                    "i18nKey": "route.obe_mkdir",
                    "icon": "mdi:folder-multiple-plus-outline",
                    "order": 1,
                },
            },
            {
                "name": "obe_tasks",
                "path": "/obe/tasks",
                "component": "layout.base$view.obe_tasks",
                "meta": {
                    "title": "任务列表",
                    "i18nKey": "route.obe_tasks",
                    "icon": "mdi:clipboard-list-outline",
                    "order": 2,
                },
            },
        ],
    },
]

CONSTANT_ROUTES = []
