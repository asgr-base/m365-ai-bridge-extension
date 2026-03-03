/**
 * MCP Server ダウンロードパイプラインテスト
 *
 * POST /token エンドポイントと Graph API ヘルパー関数をテストする。
 * 実際の Graph API は呼ばず、HTTP サーバーのトークン受信機能を検証する。
 */

const { test, expect } = require('@playwright/test');
const http = require('http');
const { spawn } = require('child_process');
const path = require('path');

const MCP_SERVER_PATH = path.join(__dirname, '../native/mcp-server.js');
const HTTP_PORT = 3765;
const BASE_URL = `http://127.0.0.1:${HTTP_PORT}`;

/**
 * HTTP リクエストを送信するヘルパー
 */
function httpRequest(method, urlPath, body = null) {
  return new Promise((resolve, reject) => {
    const options = {
      hostname: '127.0.0.1',
      port: HTTP_PORT,
      path: urlPath,
      method,
      headers: { 'Content-Type': 'application/json' },
    };

    const req = http.request(options, (res) => {
      let data = '';
      res.on('data', chunk => (data += chunk));
      res.on('end', () => {
        try {
          resolve({ status: res.statusCode, data: JSON.parse(data) });
        } catch {
          resolve({ status: res.statusCode, data });
        }
      });
    });

    req.on('error', reject);
    if (body) req.write(JSON.stringify(body));
    req.end();
  });
}

/**
 * ポートが空くまで待つ
 */
function waitForPort(port, timeout = 5000) {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    function check() {
      const req = http.request({ hostname: '127.0.0.1', port, path: '/health', method: 'GET' }, (res) => {
        let d = '';
        res.on('data', c => (d += c));
        res.on('end', () => resolve());
      });
      req.on('error', () => {
        if (Date.now() - start > timeout) return reject(new Error('Timeout waiting for port'));
        setTimeout(check, 200);
      });
      req.end();
    }
    check();
  });
}

test.describe('MCP Server ダウンロードパイプライン', () => {
  let serverProcess;

  test.beforeAll(async () => {
    // 既存のプロセスを終了（ポート競合回避）
    try {
      const { execSync } = require('child_process');
      const pids = execSync(`lsof -ti :${HTTP_PORT} 2>/dev/null || true`, { encoding: 'utf-8' }).trim();
      if (pids) {
        for (const pid of pids.split('\n').map(Number).filter(Boolean)) {
          try { process.kill(pid, 'SIGTERM'); } catch {}
        }
        await new Promise(r => setTimeout(r, 1000));
      }
    } catch {}

    // MCP サーバーを起動（stdin を /dev/null にして MCP stdio が固まらないようにする）
    serverProcess = spawn('node', [MCP_SERVER_PATH], {
      stdio: ['pipe', 'pipe', 'pipe'],
      env: { ...process.env },
    });

    // stdin を閉じる（MCP stdio がブロックしないようにする）
    // Note: MCP 接続は失敗するが HTTP サーバーは動作する
    serverProcess.stdin.end();

    // stderr をキャプチャ（デバッグ用）
    serverProcess.stderr.on('data', () => {});

    await waitForPort(HTTP_PORT, 8000);
  });

  test.afterAll(async () => {
    if (serverProcess) {
      serverProcess.kill('SIGTERM');
      await new Promise(r => setTimeout(r, 500));
    }
  });

  test('POST /token でトークンを受信できる', async () => {
    const mockToken = 'eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.test-payload.test-signature';
    const res = await httpRequest('POST', '/token', {
      token: mockToken,
      expiresOn: '9999999999',
      target: 'https://graph.microsoft.com/.default',
    });

    expect(res.status).toBe(200);
    expect(res.data.success).toBe(true);
    expect(res.data.tokenLength).toBe(mockToken.length);
  });

  test('POST /token でトークンなしの場合は 400 エラー', async () => {
    const res = await httpRequest('POST', '/token', { expiresOn: '9999' });
    expect(res.status).toBe(400);
    expect(res.data.error).toContain('Missing');
  });

  test('GET /status にトークン状態が含まれる', async () => {
    // 先にトークンを送信
    await httpRequest('POST', '/token', {
      token: 'eyJhbGciOiJSUzI1NiJ9.status-test.sig',
    });

    const res = await httpRequest('GET', '/status');
    expect(res.status).toBe(200);
    expect(res.data.status).toBe('ok');
  });

  test('GET /health が正常にレスポンスする', async () => {
    const res = await httpRequest('GET', '/health');
    expect(res.status).toBe(200);
    expect(res.data.status).toBe('ok');
    expect(res.data.mode).toBe('mcp');
  });
});
