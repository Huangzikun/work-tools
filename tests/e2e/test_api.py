"""
教案生成功能的 API 测试
直接测试后端 API 验证功能是否正常工作
"""
import requests
import json

BASE_URL = "http://localhost:5001/api"

def test_login():
    """测试登录"""
    print("🔐 测试登录...")
    response = requests.post(f"{BASE_URL}/auth/login", json={
        "userName": "admin",
        "password": "123456"
    })

    if response.status_code == 200:
        data = response.json()
        if data.get("code") in ["0000", 200]:
            token = data.get("data", {}).get("token")
            print(f"✅ 登录成功，获得 Token: {token[:20]}...")
            return token
        else:
            print(f"❌ 登录失败: {data.get('msg')}")
    else:
        print(f"❌ 登录请求失败: {response.status_code}")
    return None

def test_get_user_info(token):
    """测试获取用户信息"""
    print("\n👤 测试获取用户信息...")
    response = requests.get(f"{BASE_URL}/auth/getUserInfo", headers={
        "Authorization": f"Bearer {token}"
    })

    if response.status_code == 200:
        data = response.json()
        if data.get("code") in ["0000", 200]:
            user_name = data.get("data", {}).get("userName")
            print(f"✅ 用户信息获取成功: {user_name}")
            return True
        else:
            print(f"❌ 获取用户信息失败: {data.get('msg')}")
    else:
        print(f"❌ 请求失败: {response.status_code}")
    return False

def test_get_tasks(token):
    """测试获取任务列表"""
    print("\n📋 测试获取任务列表...")
    response = requests.get(f"{BASE_URL}/lesson-plan/tasks", headers={
        "Authorization": f"Bearer {token}"
    }, params={
        "page": 1,
        "size": 20
    })

    if response.status_code == 200:
        data = response.json()
        if data.get("code") in ["0000", 200]:
            tasks = data.get("data", {}).get("list", [])
            total = data.get("data", {}).get("total", 0)
            print(f"✅ 任务列表获取成功，共 {total} 个任务")
            for task in tasks:
                print(f"  - ID: {task.get('id')}, 课程: {task.get('courseInfo', {}).get('课程名称')}, 状态: {task.get('status')}")
            return tasks
        else:
            print(f"❌ 获取任务列表失败: {data.get('msg')}")
    else:
        print(f"❌ 请求失败: {response.status_code}")
    return []

def test_get_progress(token, task_id):
    """测试获取任务进度"""
    print(f"\n📊 测试获取任务 {task_id} 的进度...")
    response = requests.get(f"{BASE_URL}/lesson_plan/tasks/{task_id}/progress", headers={
        "Authorization": f"Bearer {token}"
    })

    if response.status_code == 200:
        data = response.json()
        if data.get("code") in ["0000", 200]:
            progress = data.get("data", {})
            status = progress.get("status")
            done = progress.get("done")
            total_count = progress.get("total")
            label = progress.get("label")
            print(f"✅ 进度获取成功: {status}, {done}/{total_count} - {label}")
            return progress
        else:
            print(f"❌ 获取进度失败: {data.get('msg')}")
    else:
        print(f"❌ 请求失败: {response.status_code}")
    return None

def test_get_default_system_prompt(token):
    """测试获取默认系统提示词"""
    print("\n🤖 测试获取默认系统提示词...")
    response = requests.get(f"{BASE_URL}/lesson-plan/default-prompt", headers={
        "Authorization": f"Bearer {token}"
    })

    if response.status_code == 200:
        data = response.json()
        if data.get("code") in ["0000", 200]:
            prompt = data.get("data")
            print(f"✅ 默认系统提示词获取成功，长度: {len(prompt)} 字符")
            print(f"前 100 字符: {prompt[:100]}...")
            return prompt
        else:
            print(f"❌ 获取默认系统提示词失败: {data.get('msg')}")
    else:
        print(f"❌ 请求失败: {response.status_code}")
    return None

def test_get_routes(token):
    """测试获取用户路由"""
    print("\n🛣️ 测试获取用户路由...")
    response = requests.get(f"{BASE_URL}/route/getUserRoutes", headers={
        "Authorization": f"Bearer {token}"
    })

    if response.status_code == 200:
        data = response.json()
        if data.get("code") in ["0000", 200]:
            routes = data.get("data", [])
            print(f"✅ 用户路由获取成功，共 {len(routes)} 个路由")
            # 查找教案相关路由
            lesson_routes = []
            def find_lesson_routes(routes, path=""):
                for route in routes:
                    if not isinstance(route, dict):
                        continue
                    current_path = f"{path}/{route.get('path')}" if path else route.get('path')
                    if "lesson" in route.get('path', '').lower():
                        lesson_routes.append({
                            "path": route.get('path'),
                            "name": route.get('name'),
                            "component": route.get('component'),
                            "children": route.get('children')
                        })
                    if route.get('children'):
                        find_lesson_routes(route.get('children'), current_path)

            find_lesson_routes(routes)

            if lesson_routes:
                print("📚 找到教案相关路由:")
                for route in lesson_routes:
                    print(f"  - {route}")
            return routes
        else:
            print(f"❌ 获取用户路由失败: {data.get('msg')}")
    else:
        print(f"❌ 请求失败: {response.status_code}")
    return []

def main():
    """主测试函数"""
    print("🚀 开始教案生成功能 API 测试...")
    print("="*60)

    # 1. 登录获取 token
    token = test_login()
    if not token:
        print("\n❌ 测试终止：无法获取登录 Token")
        return

    # 2. 获取用户信息
    test_get_user_info(token)

    # 3. 获取用户路由（验证菜单配置）
    test_get_routes(token)

    # 4. 获取任务列表
    tasks = test_get_tasks(token)

    # 5. 如果有任务，测试获取进度
    if tasks:
        # 找一个已完成的任务
        completed_task = next((t for t in tasks if t.get('status') == 'completed'), None)
        if completed_task:
            test_get_progress(token, completed_task.get('id'))

    # 6. 测试获取默认系统提示词
    test_get_default_system_prompt(token)

    print("\n" + "="*60)
    print("✅ API 测试完成")
    print("\n💡 提示：请在浏览器中访问 http://localhost:9530 进行手动 UI 验证")
    print("   - 登录用户名: admin")
    print("   - 登录密码: 123456")
    print("   - 导航路径: 教案工具 > 任务列表 / 生成教案")

if __name__ == "__main__":
    main()
