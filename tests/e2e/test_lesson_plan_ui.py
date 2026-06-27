"""
教案生成 Web 化功能的端到端 UI 测试
验证前端页面的视觉和工作流（不实际提交 LLM 任务）
"""
import pytest
import time
from pathlib import Path
from playwright.sync_api import Page, expect


class TestLessonPlanUI:
    """教案生成功能 UI 测试"""

    def test_login_and_menu_verification(self, page: Page, screenshots_dir: Path):
        """测试登录和菜单验证"""
        # 1. 访问首页
        page.goto("http://localhost:9530")
        page.wait_for_load_state("networkidle")

        # 截图：登录页面
        page.screenshot(path=str(screenshots_dir / "lp_ui_00_login.png"))

        # 2. 登录
        page.fill("input[placeholder='用户名: admin']", "admin")
        page.fill("input[type='password']", "123456")
        page.click("button:has-text('登录')")

        # 等待跳转到首页
        page.wait_for_url("**/home")
        page.wait_for_load_state("networkidle")

        # 3. 验证左侧菜单
        # 查找「教案工具」菜单
        lesson_menu = page.locator(".menu-item:has-text('教案工具')")
        expect(lesson_menu).to_be_visible()

        # 截图：首页和菜单
        page.screenshot(path=str(screenshots_dir / "lp_ui_01_home_menu.png"))

        # 展开教案工具菜单
        lesson_menu.click()
        page.wait_for_timeout(500)  # 等待展开动画

        # 验证子菜单
        expect(page.locator(".menu-item:has-text('生成教案')")).to_be_visible()
        expect(page.locator(".menu-item:has-text('任务列表')")).to_be_visible()

        # 截图：展开的菜单
        page.screenshot(path=str(screenshots_dir / "lp_ui_02_menu_expanded.png"))

        print("✅ 登录和菜单验证通过")

    def test_task_list_page(self, page: Page, screenshots_dir: Path):
        """测试任务列表页面"""
        # 登录（如果未登录）
        if "login" in page.url:
            page.fill("input[placeholder='用户名: admin']", "admin")
            page.fill("input[type='password']", "123456")
            page.click("button:has-text('登录')")
            page.wait_for_url("**/home")

        # 导航到任务列表
        page.click(".menu-item:has-text('教案工具')")
        page.wait_for_timeout(500)
        page.click(".menu-item:has-text('任务列表')")
        page.wait_for_url("**/lesson_plan/tasks")
        page.wait_for_load_state("networkidle")

        # 截图：任务列表页
        page.screenshot(path=str(screenshots_dir / "lp_ui_03_task_list.png"))

        # 验证表格存在
        table = page.locator(".n-data-table")
        expect(table).to_be_visible()

        # 验证表头
        expect(page.locator("th:has-text('ID')")).to_be_visible()
        expect(page.locator("th:has-text('大纲文件')")).to_be_visible()
        expect(page.locator("th:has-text('课程名')")).to_be_visible()
        expect(page.locator("th:has-text('总课时')")).to_be_visible()
        expect(page.locator("th:has-text('教师')")).to_be_visible()
        expect(page.locator("th:has-text('状态')")).to_be_visible()
        expect(page.locator("th:has-text('进度')")).to_be_visible()
        expect(page.locator("th:has-text('创建时间')")).to_be_visible()
        expect(page.locator("th:has-text('操作')")).to_be_visible()

        # 验证已存在的任务记录（ID=2）
        expect(page.locator("td:has-text('2')")).to_be_visible()
        expect(page.locator("td:has-text('测试课程')")).to_be_visible()

        # 验证状态标签和操作按钮
        expect(page.locator(".n-tag--success")).to_be_visible()  # 已完成状态
        expect(page.locator("button:has-text('下载')")).to_be_visible()
        expect(page.locator("button:has-text('重新生成')")).to_be_visible()
        expect(page.locator("button:has-text('删除')")).to_be_visible()

        # 验证「新建生成」按钮
        create_btn = page.locator("button:has-text('新建生成')")
        expect(create_btn).to_be_visible()

        print("✅ 任务列表页面验证通过")

    def test_generate_form_page(self, page: Page, screenshots_dir: Path):
        """测试生成表单页面"""
        # 登录并导航到任务列表（如果需要）
        if "login" in page.url:
            page.fill("input[placeholder='用户名: admin']", "admin")
            page.fill("input[type='password']", "123456")
            page.click("button:has-text('登录')")
            page.wait_for_url("**/home")

        if "lesson_plan/tasks" not in page.url:
            page.click(".menu-item:has-text('教案工具')")
            page.wait_for_timeout(500)
            page.click(".menu-item:has-text('任务列表')")
            page.wait_for_url("**/lesson_plan/tasks")
            page.wait_for_load_state("networkidle")

        # 点击「新建生成」按钮
        page.click("button:has-text('新建生成')")
        page.wait_for_url("**/lesson_plan/generate")
        page.wait_for_load_state("networkidle")

        # 截图：生成表单页初始状态
        page.screenshot(path=str(screenshots_dir / "lp_ui_04_generate_form.png"), full_page=True)

        # 验证 5 个卡片
        expect(page.locator("text=上传教学大纲")).to_be_visible()
        expect(page.locator("text=生成参数")).to_be_visible()
        expect(page.locator("text=课程基本信息（首页）")).to_be_visible()
        expect(page.locator("text=教师信息（首页）")).to_be_visible()
        expect(page.locator("text=高级设置")).to_be_visible()

        print("✅ 生成表单页面卡片验证通过")

    def test_form_fill_and_validation(self, page: Page, screenshots_dir: Path):
        """测试表单填写和验证"""
        # 确保在生成表单页面
        if "lesson_plan/generate" not in page.url:
            # 登录
            if "login" in page.url:
                page.fill("input[placeholder='用户名: admin']", "admin")
                page.fill("input[type='password']", "123456")
                page.click("button:has-text('登录')")
                page.wait_for_url("**/home")

            # 导航到生成页面
            page.click(".menu-item:has-text('教案工具')")
            page.wait_for_timeout(500)
            page.click(".menu-item:has-text('任务列表')")
            page.wait_for_url("**/lesson_plan/tasks")
            page.click("button:has-text('新建生成')")
            page.wait_for_url("**/lesson_plan/generate")
            page.wait_for_load_state("networkidle")

        # 填写表单
        page.fill("input[placeholder='请输入课程名称']", "测试UI")
        page.fill("input[placeholder='请输入学分']", "4")
        page.fill("input[placeholder='请输入理论学时']", "64")
        page.fill("input[placeholder='请输入适用专业']", "计算机科学")
        page.fill("input[placeholder='请输入授课教师']", "admin")
        page.fill("input[placeholder='请输入总课时数']", "2")
        page.fill("input[placeholder='请输入批量大小']", "1")

        # 等待输入完成
        page.wait_for_timeout(500)

        # 截图：表单填写后
        page.screenshot(path=str(screenshots_dir / "lp_ui_05_form_filled.png"), full_page=True)

        # 展开「高级设置」折叠面板
        advanced_collapse = page.locator(".n-collapse-item:has-text('高级设置')")
        if not advanced_collapse.get_attribute("aria-expanded") == "true":
            advanced_collapse.click()
            page.wait_for_timeout(300)

        # 验证 System Prompt textarea 可见
        expect(page.locator("textarea[placeholder*='System Prompt']")).to_be_visible()
        expect(page.locator("textarea[placeholder*='System Prompt']")).not_to_have_value("")

        # 截图：高级设置展开
        page.screenshot(path=str(screenshots_dir / "lp_ui_06_advanced_settings.png"), full_page=True)

        print("✅ 表单填写和验证通过")

    def test_table_visual_verification(self, page: Page, screenshots_dir: Path):
        """测试表格视觉验证（状态标签、进度条）"""
        # 导航到任务列表
        if "lesson_plan/tasks" not in page.url:
            # 登录
            if "login" in page.url:
                page.fill("input[placeholder='用户名: admin']", "admin")
                page.fill("input[type='password']", "123456")
                page.click("button:has-text('登录')")
                page.wait_for_url("**/home")

            page.click(".menu-item:has-text('教案工具')")
            page.wait_for_timeout(500)
            page.click(".menu-item:has-text('任务列表')")
            page.wait_for_url("**/lesson_plan/tasks")
            page.wait_for_load_state("networkidle")

        # 截图：表格详细视图
        page.screenshot(path=str(screenshots_dir / "lp_ui_07_table_visual.png"))

        # 验证状态标签颜色
        success_tag = page.locator(".n-tag--success:has-text('已完成')")
        expect(success_tag).to_be_visible()

        # 验证进度条
        progress_bar = page.locator(".n-progress")
        expect(progress_bar).to_be_visible()

        # 截图：状态标签和进度条特写
        success_tag.screenshot(path=str(screenshots_dir / "lp_ui_08_success_tag.png"))
        progress_bar.screenshot(path=str(screenshots_dir / "lp_ui_09_progress_bar.png"))

        print("✅ 表格视觉验证通过")

    def test_console_and_network(self, page: Page, screenshots_dir: Path):
        """测试控制台错误和网络请求"""
        # 收集控制台日志
        console_errors = []
        page.on("console", lambda msg: console_errors.append(msg.text) if msg.type == "error" else None)

        # 收集网络失败
        failed_requests = []
        def handle_request_failed(request):
            failed_requests.append(request.url)

        page.on("requestfailed", handle_request_failed)

        # 执行完整流程
        # 登录
        page.goto("http://localhost:9530")
        page.wait_for_load_state("networkidle")
        page.fill("input[placeholder='用户名: admin']", "admin")
        page.fill("input[type='password']", "123456")
        page.click("button:has-text('登录')")
        page.wait_for_url("**/home")

        # 导航到任务列表
        page.click(".menu-item:has-text('教案工具')")
        page.wait_for_timeout(500)
        page.click(".menu-item:has-text('任务列表')")
        page.wait_for_url("**/lesson_plan/tasks")
        page.wait_for_load_state("networkidle")

        # 导航到生成页面
        page.click("button:has-text('新建生成')")
        page.wait_for_url("**/lesson_plan/generate")
        page.wait_for_load_state("networkidle")

        # 填写表单
        page.fill("input[placeholder='请输入课程名称']", "测试UI")
        page.fill("input[placeholder='请输入学分']", "4")
        page.fill("input[placeholder='请输入理论学时']", "64")
        page.fill("input[placeholder='请输入适用专业']", "计算机科学")
        page.fill("input[placeholder='请输入授课教师']", "admin")
        page.fill("input[placeholder='请输入总课时数']", "2")
        page.fill("input[placeholder='请输入批量大小']", "1")
        page.wait_for_timeout(500)

        # 展开高级设置
        page.locator(".n-collapse-item:has-text('高级设置')").click()
        page.wait_for_timeout(300)

        # 等待所有请求完成
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)

        # 截图：最终状态
        page.screenshot(path=str(screenshots_dir / "lp_ui_10_final_state.png"), full_page=True)

        # 输出结果
        print(f"\n📊 控制台错误数量: {len(console_errors)}")
        if console_errors:
            print("❌ 控制台错误:")
            for error in console_errors:
                print(f"  - {error}")
        else:
            print("✅ 无控制台错误")

        print(f"\n📊 网络请求失败数量: {len(failed_requests)}")
        if failed_requests:
            print("❌ 网络请求失败:")
            for url in failed_requests:
                print(f"  - {url}")
        else:
            print("✅ 无网络请求失败")


@pytest.fixture
def screenshots_dir(tmp_path):
    """创建截图目录"""
    screenshots = tmp_path / "screenshots"
    screenshots.mkdir(exist_ok=True)
    return screenshots


if __name__ == "__main__":
    # 可以直接运行此文件进行测试
    import sys
    from playwright.sync_api import sync_playwright

    screenshots_path = Path("/tmp/lesson_plan_ui_screenshots")
    screenshots_path.mkdir(exist_ok=True)

    print("🚀 开始教案生成功能 UI 测试...")
    print(f"📸 截图将保存到: {screenshots_path}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # 设置为 False 可以看到浏览器
        context = browser.new_context()
        page = context.new_page()

        try:
            # 执行所有测试
            test = TestLessonPlanUI()

            print("\n" + "="*60)
            print("🔐 测试 1: 登录和菜单验证")
            print("="*60)
            test.test_login_and_menu_verification(page, screenshots_path)

            print("\n" + "="*60)
            print("📋 测试 2: 任务列表页面")
            print("="*60)
            test.test_task_list_page(page, screenshots_path)

            print("\n" + "="*60)
            print("📝 测试 3: 生成表单页面")
            print("="*60)
            test.test_generate_form_page(page, screenshots_path)

            print("\n" + "="*60)
            print("✍️ 测试 4: 表单填写和验证")
            print("="*60)
            test.test_form_fill_and_validation(page, screenshots_path)

            print("\n" + "="*60)
            print("🎨 测试 5: 表格视觉验证")
            print("="*60)
            test.test_table_visual_verification(page, screenshots_path)

            print("\n" + "="*60)
            print("🔍 测试 6: 控制台和网络验证")
            print("="*60)
            test.test_console_and_network(page, screenshots_path)

            print("\n" + "="*60)
            print("✅ 所有测试完成！")
            print("="*60)
            print(f"\n📸 截图已保存到: {screenshots_path}")
            print("文件列表:")
            for screenshot in sorted(screenshots_path.glob("*.png")):
                print(f"  - {screenshot.name}")

        except Exception as e:
            print(f"\n❌ 测试失败: {e}")
            import traceback
            traceback.print_exc()

            # 失败时截图
            error_screenshot = screenshots_path / f"error_{int(time.time())}.png"
            page.screenshot(path=str(error_screenshot), full_page=True)
            print(f"📸 错误截图: {error_screenshot}")

        finally:
            page.close()
            context.close()
            browser.close()
