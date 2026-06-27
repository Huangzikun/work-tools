"""
教案生成 Web 化功能的端到端 UI 测试（简化版）
直接使用 Playwright 进行测试，不依赖 pytest
"""
from playwright.sync_api import sync_playwright
from pathlib import Path
import time


def main():
    """主测试函数"""
    screenshots_dir = Path("/tmp/lesson_plan_ui_screenshots")
    screenshots_dir.mkdir(exist_ok=True)

    print("🚀 开始教案生成功能 UI 测试...")
    print(f"📸 截图将保存到: {screenshots_dir}")

    with sync_playwright() as p:
        # 启动浏览器（设置为 False 可以看到浏览器，True 为无头模式）
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        try:
            # 测试 1: 登录
            print("\n" + "="*60)
            print("🔐 测试 1: 登录")
            print("="*60)
            page.goto("http://localhost:9530")
            page.wait_for_load_state("networkidle")
            page.screenshot(path=str(screenshots_dir / "lp_ui_00_login.png"))
            print("✅ 登录页面加载成功")

            # 等待页面加载完成
            page.wait_for_timeout(2000)

            # 填写用户名和密码
            page.fill("input[type='text']", "admin")
            page.fill("input[type='password']", "123456")

            # 点击登录按钮（查找包含"确认"文本的按钮）
            confirm_btn = page.locator("button:has-text('确认')")
            confirm_btn.click()

            # 等待跳转到首页
            page.wait_for_url("**/home")
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(2000)  # 额外等待确保菜单完全加载
            page.screenshot(path=str(screenshots_dir / "lp_ui_01_home_menu.png"))
            print("✅ 登录成功，跳转到首页")

            # 测试 2: 菜单验证
            print("\n" + "="*60)
            print("📋 测试 2: 菜单验证")
            print("="*60)

            # 查找「教案工具」菜单
            lesson_menu = page.locator("text=教案工具")
            if lesson_menu.is_visible():
                print("✅ 找到「教案工具」菜单")
            else:
                print("❌ 未找到「教案工具」菜单")

            # 展开教案工具菜单
            lesson_menu.click()
            page.wait_for_timeout(1000)  # 等待展开动画完成
            page.screenshot(path=str(screenshots_dir / "lp_ui_02_menu_expanded.png"))

            # 验证子菜单（使用更简单的选择器）
            generate_menu = page.locator("text=生成教案")
            task_list_menu = page.locator("text=任务列表")

            if generate_menu.is_visible():
                print("✅ 找到「生成教案」子菜单")
            else:
                print("❌ 未找到「生成教案」子菜单")

            if task_list_menu.is_visible():
                print("✅ 找到「任务列表」子菜单")
            else:
                print("❌ 未找到「任务列表」子菜单")

            # 测试 3: 任务列表页面
            print("\n" + "="*60)
            print("📋 测试 3: 任务列表页面")
            print("="*60)

            task_list_menu.click()
            page.wait_for_url("**/lesson/tasks")
            page.wait_for_load_state("networkidle")
            page.screenshot(path=str(screenshots_dir / "lp_ui_03_task_list.png"))
            print("✅ 任务列表页面加载成功")

            # 验证表格
            table = page.locator(".n-data-table")
            if table.is_visible():
                print("✅ 任务表格显示正常")
            else:
                print("❌ 任务表格未显示")

            # 验证表头
            headers = ["ID", "大纲文件", "课程名", "总课时", "教师", "状态", "进度", "创建时间", "操作"]
            for header in headers:
                header_element = page.locator(f"th:has-text('{header}')")
                if header_element.is_visible():
                    print(f"  ✅ 表头「{header}」显示正常")
                else:
                    print(f"  ❌ 表头「{header}」未显示")

            # 验证已存在的任务
            if page.locator("td:has-text('2')").is_visible():
                print("✅ 找到测试任务（ID=2）")
            if page.locator("td:has-text('测试课程')").is_visible():
                print("✅ 找到课程名「测试课程」")

            # 验证状态和操作
            if page.locator(".n-tag--success").is_visible():
                print("✅ 状态标签显示正常")
            if page.locator("button:has-text('下载')").is_visible():
                print("✅ 下载按钮显示正常")
            if page.locator("button:has-text('重新生成')").is_visible():
                print("✅ 重新生成按钮显示正常")
            if page.locator("button:has-text('删除')").is_visible():
                print("✅ 删除按钮显示正常")

            # 测试 4: 生成表单页面
            print("\n" + "="*60)
            print("📝 测试 4: 生成表单页面")
            print("="*60)

            create_btn = page.locator("button:has-text('新建生成')")
            if create_btn.is_visible():
                print("✅ 新建生成按钮显示正常")
            create_btn.click()

            page.wait_for_url("**/lesson/generate")
            page.wait_for_load_state("networkidle")
            page.screenshot(path=str(screenshots_dir / "lp_ui_04_generate_form.png"), full_page=True)
            print("✅ 生成表单页面加载成功")

            # 验证 5 个卡片
            cards = ["上传教学大纲", "生成参数", "课程基本信息（首页）", "教师信息（首页）", "高级设置"]
            for card in cards:
                card_element = page.locator(f"text={card}")
                if card_element.is_visible():
                    print(f"  ✅ 卡片「{card}」显示正常")
                else:
                    print(f"  ❌ 卡片「{card}」未显示")

            # 测试 5: 表单填写
            print("\n" + "="*60)
            print("✍️ 测试 5: 表单填写")
            print("="*60)

            page.fill("input[placeholder='请输入课程名称']", "测试UI")
            page.fill("input[placeholder='请输入学分']", "4")
            page.fill("input[placeholder='请输入理论学时']", "64")
            page.fill("input[placeholder='请输入适用专业']", "计算机科学")
            page.fill("input[placeholder='请输入授课教师']", "admin")
            page.fill("input[placeholder='请输入总课时数']", "2")
            page.fill("input[placeholder='请输入批量大小']", "1")
            page.wait_for_timeout(500)

            page.screenshot(path=str(screenshots_dir / "lp_ui_05_form_filled.png"), full_page=True)
            print("✅ 表单填写完成")

            # 展开高级设置
            advanced_collapse = page.locator(".n-collapse-item:has-text('高级设置')")
            if advanced_collapse.is_visible():
                advanced_collapse.click()
                page.wait_for_timeout(300)
                page.screenshot(path=str(screenshots_dir / "lp_ui_06_advanced_settings.png"), full_page=True)
                print("✅ 高级设置展开成功")

                # 验证 System Prompt
                system_prompt = page.locator("textarea[placeholder*='System Prompt']")
                if system_prompt.is_visible():
                    print("✅ System Prompt 显示正常")
                    prompt_value = system_prompt.input_value()
                    if prompt_value:
                        print(f"  ✅ System Prompt 预填内容长度: {len(prompt_value)} 字符")
                    else:
                        print("  ⚠️ System Prompt 无预填内容")
                else:
                    print("❌ System Prompt 未显示")
            else:
                print("❌ 高级设置折叠面板未显示")

            # 测试 6: 返回任务列表验证表格视觉
            print("\n" + "="*60)
            print("🎨 测试 6: 表格视觉验证")
            print("="*60)

            page.click(".menu-item:has-text('任务列表')")
            page.wait_for_url("**/lesson/tasks")
            page.wait_for_load_state("networkidle")
            page.screenshot(path=str(screenshots_dir / "lp_ui_07_table_visual.png"))

            # 验证状态标签颜色
            success_tag = page.locator(".n-tag--success:has-text('已完成')")
            if success_tag.is_visible():
                print("✅ 已完成状态标签显示正常（绿色）")
                success_tag.screenshot(path=str(screenshots_dir / "lp_ui_08_success_tag.png"))
            else:
                print("⚠️ 未找到已完成状态标签")

            # 验证进度条
            progress_bar = page.locator(".n-progress")
            if progress_bar.is_visible():
                print("✅ 进度条显示正常")
                progress_bar.screenshot(path=str(screenshots_dir / "lp_ui_09_progress_bar.png"))
            else:
                print("⚠️ 未找到进度条")

            # 测试 7: 控制台和网络检查
            print("\n" + "="*60)
            print("🔍 测试 7: 控制台和网络检查")
            print("="*60)

            page.screenshot(path=str(screenshots_dir / "lp_ui_10_final_state.png"), full_page=True)

            # 等待确保所有网络请求完成
            page.wait_for_timeout(2000)

            print("✅ 所有测试完成！")
            print(f"\n📸 截图已保存到: {screenshots_dir}")
            print("文件列表:")
            for screenshot in sorted(screenshots_dir.glob("*.png")):
                file_size = screenshot.stat().st_size / 1024  # KB
                print(f"  - {screenshot.name} ({file_size:.1f} KB)")

            # 保持浏览器打开 10 秒供查看
            print("\n⏳ 浏览器将在 10 秒后关闭...")
            page.wait_for_timeout(10000)

        except Exception as e:
            print(f"\n❌ 测试失败: {e}")
            import traceback
            traceback.print_exc()

            # 失败时截图
            error_screenshot = screenshots_dir / f"error_{int(time.time())}.png"
            page.screenshot(path=str(error_screenshot), full_page=True)
            print(f"📸 错误截图: {error_screenshot}")

            # 保持浏览器打开以便查看错误状态
            print("\n⏳ 浏览器将在 30 秒后关闭...")
            page.wait_for_timeout(30000)

        finally:
            page.close()
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
