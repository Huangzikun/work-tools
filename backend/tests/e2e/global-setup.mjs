import { request as pwRequest } from '@playwright/test';
import { writeFileSync, mkdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname } from 'node:path';

const API_BASE = process.env.E2E_API_BASE || 'http://127.0.0.1:9530/proxy-default';
const __dirname = dirname(fileURLToPath(import.meta.url));
const STATE_PATH = __dirname + '/.auth/state.json';

export default async function globalSetup() {
  const ctx = await pwRequest.newContext();
  const res = await ctx.post(`${API_BASE}/auth/login`, {
    data: { userName: 'admin', password: '123456' }
  });
  if (!res.ok()) throw new Error(`login failed: ${res.status()}`);
  const body = await res.json();
  if (body.code !== '0000') throw new Error(`login business error: ${JSON.stringify(body)}`);
  await ctx.dispose();

  // soybean-admin 的 localStg.set 会写 `${VITE_STORAGE_PREFIX}${key}` = JSON.stringify(value)
  // VITE_STORAGE_PREFIX=SOY_（见 web/.env）
  const PREFIX = 'SOY_';
  const state = {
    cookies: [],
    origins: [
      {
        origin: 'http://127.0.0.1:9530',
        localStorage: [
          { name: `${PREFIX}token`, value: JSON.stringify(body.data.token) },
          { name: `${PREFIX}refreshToken`, value: JSON.stringify(body.data.refreshToken) }
        ]
      }
    ]
  };
  mkdirSync(__dirname + '/.auth', { recursive: true });
  writeFileSync(STATE_PATH, JSON.stringify(state, null, 2));
  console.log('[globalSetup] auth state written to', STATE_PATH, '(prefix=' + PREFIX + ')');
}
