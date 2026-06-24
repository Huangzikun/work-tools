// OBE 目录生成 E2E 测试（Playwright）
// 依赖 playwright.config.mjs 的 globalSetup：API 登录后写入 .auth/state.json，
// 通过 storageState 注入到每个 test 的 context（localStorage key=SOY_token，值是 JSON 字符串）

import { test, expect } from '@playwright/test';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { execSync } from 'node:child_process';

const BASE = 'http://127.0.0.1:9530';
const ROSTER = '/Users/huangzikun/PycharmProjects/work-tools/backend/tests/fixtures/test_roster.xls';

// vite dev 首次访问路由需要 transform 大量 .vue 文件，渲染较慢（约 10-15s）。
// 单次 goto + 长 timeout（30s）等 DOM 元素出现。
async function gotoStable(page, url, domWaitFn, retries = 3) {
  let lastErr;
  for (let i = 0; i < retries; i++) {
    await page.goto(url);
    await page.waitForLoadState('domcontentloaded');
    try {
      await domWaitFn({ timeout: 30000 });
      return;
    } catch (e) {
      lastErr = e;
      console.log(`[gotoStable] ${url} attempt ${i + 1}/${retries} DOM not ready, retrying...`);
      await page.waitForTimeout(2000);
    }
  }
  throw lastErr;
}

test.describe('OBE 目录生成 E2E', () => {
  test('步骤 1：登录（UI 验证）', async ({ page }) => {
    await page.goto(BASE + '/login?module=pwd-login');
    await page.waitForLoadState('domcontentloaded');
    await page.evaluate(() => {
      for (const k of Object.keys(window.localStorage)) {
        if (k.endsWith('token') || k.endsWith('refreshToken')) {
          window.localStorage.removeItem(k);
        }
      }
    });
    await page.reload();
    await page.waitForLoadState('domcontentloaded');

    const pwdInput = page.locator('input[type="password"]').first();
    await pwdInput.waitFor({ state: 'visible', timeout: 30000 });
    await page.locator('input').first().fill('admin');
    await pwdInput.fill('123456');

    const loginRespPromise = page.waitForResponse(
      r => r.url().includes('/proxy-default/auth/login') && r.request().method() === 'POST',
      { timeout: 30000 }
    );
    await page.getByRole('button', { name: /^确认$/ }).first().click();
    const loginResp = await loginRespPromise;
    console.log('[Step1] login HTTP', loginResp.status());
    expect(loginResp.ok()).toBeTruthy();
    const body = await loginResp.json();
    expect(body.code).toBe('0000');
    console.log('[Step1] 登录 API 返回 token 成功（UI 登录链路正常）');
  });

  // 步骤 2 + 3 合并：避免重复 goto /obe/mkdir 触发 vite dev 重新 transform 模块
  test('步骤 2+3：验证菜单与页面渲染', async ({ page }) => {
    const obeMenu = page.locator('text=OBE工具').first();
    // 注意：访问 /home 在 vite dev 下偶发 504 Outdated Optimize Dep（home/index.vue 模块懒加载问题），
    // 这里改从 /obe/mkdir 验证菜单（菜单属于 base-layout，任何 authed 路由都有）
    await gotoStable(page, BASE + '/obe/mkdir', async ({ timeout }) => {
      await expect(obeMenu).toBeVisible({ timeout });
    });

    // === 步骤 2：菜单交互 ===
    await obeMenu.click();
    await page.waitForTimeout(800);
    const subMenu = page.locator('text=目录生成').first();
    await expect(subMenu).toBeVisible({ timeout: 5000 });
    await subMenu.click();
    await page.waitForURL(/\/obe\/mkdir/, { timeout: 10000 });
    expect(page.url()).toMatch(/\/obe\/mkdir/);
    console.log('[Step2] 菜单点击导航到 /obe/mkdir 成功');

    // === 步骤 3：验证页面渲染 ===
    const alert = page.locator('div[role=alert]').first();
    await expect(alert).toBeVisible({ timeout: 45000 });
    const alertText = await alert.innerText();
    console.log('[Step3] Alert 内容:', alertText.slice(0, 100));
    await expect(page.getByRole('heading', { name: 'OBE 目录生成' }).first()).toBeVisible();
    for (const label of ['班级名称', '课程名称', '教师姓名', '固定目录', '考核目录', '名单文件']) {
      await expect(page.locator('.n-form-item-label', { hasText: label }).first()).toBeVisible();
    }
    const checkedCount = await page.locator('.n-checkbox.n-checkbox--checked').count();
    console.log('[Step3] 默认勾选 checkbox 数量:', checkedCount);
    const tags = await page.locator('.n-base-selection-tag-wrapper .n-tag').count();
    console.log('[Step3] select 默认已选 tag 数量:', tags);
    expect(checkedCount).toBeGreaterThanOrEqual(2);
    expect(tags).toBeGreaterThanOrEqual(2);
    await page.screenshot({ path: '/Users/huangzikun/PycharmProjects/work-tools/backend/tests/e2e/step3_default.png', fullPage: true });
  });

  test('步骤 4：表单校验测试', async ({ page }) => {
    const submitBtn = page.getByRole('button', { name: /生成/ }).first();
    await gotoStable(page, BASE + '/obe/mkdir', async ({ timeout }) => {
      await expect(submitBtn).toBeVisible({ timeout });
    });
    await submitBtn.click();
    await page.waitForTimeout(1500);
    const errCount = await page.locator('.n-form-item-feedback__line, .n-form-item--error').count();
    console.log('[Step4] 表单错误条数:', errCount);
    await page.screenshot({ path: '/Users/huangzikun/PycharmProjects/work-tools/backend/tests/e2e/step4_validation.png', fullPage: true });
    expect(errCount).toBeGreaterThan(0);
  });

  test('步骤 5：完整提交流程', async ({ page }) => {
    const submitBtn = page.getByRole('button', { name: /生成/ }).first();
    await gotoStable(page, BASE + '/obe/mkdir', async ({ timeout }) => {
      await expect(submitBtn).toBeVisible({ timeout });
    });

    const inputs = page.locator('.n-form-item .n-input__input-el');
    await inputs.nth(0).fill('2022级数据科学与大数据技术2班');
    await inputs.nth(1).fill('面向对象程序设计');
    await inputs.nth(2).fill('黄子坤');

    const fileInput = page.locator('input[type="file"]').first();
    await fileInput.setInputFiles(ROSTER);
    await page.waitForTimeout(1500);
    const hasFile = await page.locator('.n-upload-file').count();
    console.log('[Step5] 上传后文件列表条数:', hasFile);

    const downloadPromise = page.waitForEvent('download', { timeout: 120000 });
    await submitBtn.click();

    const download = await downloadPromise;
    const suggested = download.suggestedFilename();
    console.log('[Step5] 下载文件名:', suggested);

    const tmpDir = await mkdtemp(join(tmpdir(), 'obe-'));
    const zipPath = join(tmpDir, suggested);
    await download.saveAs(zipPath);
    console.log('[Step5] ZIP 保存到:', zipPath);

    const py = `import zipfile,sys,json;
z=zipfile.ZipFile(sys.argv[1])
out=[i.filename for i in z.infolist()]
print(json.dumps(out, ensure_ascii=False))`;
    const out = execSync(`python3 -c "${py}" "${zipPath}"`, { encoding: 'utf-8' });
    const entries = JSON.parse(out);
    console.log('[Step5] zip entries:\n' + entries.join('\n'));

    expect(suggested).toBe('2022级数据科学与大数据技术2班《面向对象程序设计》OBE目录.zip');
    const joined = entries.join('\n');
    expect(joined).toContain('教学课件');
    expect(joined).toContain('教学教案');
    expect(joined).toContain('课程考核');
    expect(joined).toContain('实验实训报告');
    const dirEntries = entries.filter(e => e.endsWith('/'));
    const topLevel = dirEntries.filter(e => e.split('/').filter(Boolean).length === 1);
    const studentSub = dirEntries.filter(e => e.split('/').filter(Boolean).length === 2);
    console.log('[Step5] 顶层目录数:', topLevel.length, '学生子目录数:', studentSub.length);
    expect(topLevel.length).toBe(4);
    expect(studentSub.length).toBe(10);
  });

  test('步骤 6a：仅固定目录', async ({ page }) => {
    const submitBtn = page.getByRole('button', { name: /生成/ }).first();
    await gotoStable(page, BASE + '/obe/mkdir', async ({ timeout }) => {
      await expect(submitBtn).toBeVisible({ timeout });
    });

    const inputs = page.locator('.n-form-item .n-input__input-el');
    await inputs.nth(0).fill('测试班级A');
    await inputs.nth(1).fill('测试课程A');
    await inputs.nth(2).fill('测试教师A');

    for (let i = 0; i < 10; i++) {
      const b = page.locator('.n-base-selection-tag-wrapper .n-base-close').first();
      if (!(await b.count())) break;
      await b.click();
      await page.waitForTimeout(200);
    }

    await page.locator('input[type="file"]').first().setInputFiles(ROSTER);
    await page.waitForTimeout(1000);

    const downloadPromise = page.waitForEvent('download', { timeout: 120000 });
    await submitBtn.click();
    const download = await downloadPromise;
    const tmpDir = await mkdtemp(join(tmpdir(), 'obe-a-'));
    const zipPath = join(tmpDir, download.suggestedFilename());
    await download.saveAs(zipPath);

    const py = `import zipfile,sys,json;
z=zipfile.ZipFile(sys.argv[1])
out=[i.filename for i in z.infolist()]
print(json.dumps(out, ensure_ascii=False))`;
    const out = execSync(`python3 -c "${py}" "${zipPath}"`, { encoding: 'utf-8' });
    const entries = JSON.parse(out);
    const dirEntries = entries.filter(e => e.endsWith('/'));
    const topLevel = dirEntries.filter(e => e.split('/').filter(Boolean).length === 1);
    console.log('[Step6a] 仅固定目录 - 顶层目录数:', topLevel.length);
    console.log('[Step6a] 内容:\n' + entries.join('\n'));
    expect(topLevel.length).toBe(2);
    const joined = entries.join('\n');
    expect(joined).toContain('教学课件');
    expect(joined).toContain('教学教案');
    expect(joined).not.toContain('课程考核');
  });

  test('步骤 6b：全空（无固定也无考核）', async ({ page }) => {
    const submitBtn = page.getByRole('button', { name: /生成/ }).first();
    await gotoStable(page, BASE + '/obe/mkdir', async ({ timeout }) => {
      await expect(submitBtn).toBeVisible({ timeout });
    });

    const inputs = page.locator('.n-form-item .n-input__input-el');
    await inputs.nth(0).fill('测试班级B');
    await inputs.nth(1).fill('测试课程B');
    await inputs.nth(2).fill('测试教师B');

    for (let i = 0; i < 10; i++) {
      const cb = page.locator('.n-checkbox.n-checkbox--checked').first();
      if (!(await cb.count())) break;
      await cb.click();
      await page.waitForTimeout(200);
    }
    for (let i = 0; i < 10; i++) {
      const b = page.locator('.n-base-selection-tag-wrapper .n-base-close').first();
      if (!(await b.count())) break;
      await b.click();
      await page.waitForTimeout(200);
    }

    await page.locator('input[type="file"]').first().setInputFiles(ROSTER);
    await page.waitForTimeout(1000);

    let downloadFired = false;
    page.on('download', () => { downloadFired = true; });
    await submitBtn.click();
    await page.waitForTimeout(3500);
    console.log('[Step6b] 是否触发下载:', downloadFired);
    await page.screenshot({ path: '/Users/huangzikun/PycharmProjects/work-tools/backend/tests/e2e/step6b.png', fullPage: true });
    expect(downloadFired).toBe(false);
  });
});
