import { defineConfig } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const __dirname = fileURLToPath(new URL('.', import.meta.url));

export default defineConfig({
  testDir: './',
  timeout: 180000,
  expect: { timeout: 20000 },
  fullyParallel: false,
  workers: 1,
  reporter: [['list'], ['html', { outputFolder: './playwright-report', open: 'never' }]],
  globalSetup: __dirname + '/global-setup.mjs',
  use: {
    headless: true,
    baseURL: 'http://127.0.0.1:9530',
    actionTimeout: 20000,
    viewport: { width: 1440, height: 900 },
    screenshot: 'only-on-failure',
    video: 'retain-on-failure',
    trace: 'retain-on-failure',
    acceptDownloads: true,
    channel: 'chrome',
    storageState: __dirname + '/.auth/state.json',
    navigationTimeout: 60000,
  },
  projects: [{ name: 'chrome' }],
});
