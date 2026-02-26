/**
 * ブリッジサーバー HTTP API テスト
 * - ヘルスチェック
 * - メッセージ投稿・取得
 * - 返信テキストのキュー
 */

const { test, expect } = require('@playwright/test');
const { spawn } = require('child_process');
const path = require('path');

const BRIDGE_URL = 'http://localhost:3765';
const SERVER_PATH = path.join(__dirname, '../native/bridge-server.js');

let serverProcess;

test.beforeAll(async () => {
  // ブリッジサーバーをサブプロセスで起動
  serverProcess = spawn('node', [SERVER_PATH], {
    stdio: ['ignore', 'pipe', 'pipe'],
  });

  // 起動待ち（最大3秒）
  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('Server start timeout')), 3000);
    serverProcess.stdout.on('data', (data) => {
      if (data.toString().includes('起動')) {
        clearTimeout(timer);
        resolve();
      }
    });
    serverProcess.on('error', reject);
  });
});

test.afterAll(async () => {
  if (serverProcess) serverProcess.kill();
});

test('GET /health はOKを返す', async ({ request }) => {
  const res = await request.get(`${BRIDGE_URL}/health`);
  expect(res.status()).toBe(200);
  const body = await res.json();
  expect(body.status).toBe('ok');
  expect(body.version).toBe('0.1.0');
});

test('GET /messages はメッセージなし時に404を返す', async ({ request }) => {
  const res = await request.get(`${BRIDGE_URL}/messages`);
  // 初期状態はバッファなし
  expect(res.status()).toBe(404);
  const body = await res.json();
  expect(body.error).toBeTruthy();
});

test('POST /messages → GET /messages でメッセージが取得できる', async ({ request }) => {
  const payload = {
    context: { channelName: 'General', url: 'https://teams.microsoft.com/...' },
    messages: [
      { index: 0, sender: '田中 太郎', body: 'おはようございます', timestamp: '2026-02-26T09:00:00Z' },
      { index: 1, sender: '山田 花子', body: '準備できています', timestamp: '2026-02-26T09:05:00Z' },
    ],
    extractedAt: new Date().toISOString(),
    method: 'primary',
  };

  // POST でメッセージを登録
  const postRes = await request.post(`${BRIDGE_URL}/messages`, { data: payload });
  expect(postRes.status()).toBe(200);
  const postBody = await postRes.json();
  expect(postBody.messageCount).toBe(2);

  // GET で取得
  const getRes = await request.get(`${BRIDGE_URL}/messages`);
  expect(getRes.status()).toBe(200);
  const getBody = await getRes.json();
  expect(getBody.messages).toHaveLength(2);
  expect(getBody.messages[0].sender).toBe('田中 太郎');
  expect(getBody.context.channelName).toBe('General');
  expect(getBody.receivedAt).toBeTruthy();
});

test('POST /reply → GET /pending-reply で返信テキストが取得できる', async ({ request }) => {
  const replyText = 'ご確認いただきありがとうございます。13時から参加します。';

  // 返信テキストをキュー
  const postRes = await request.post(`${BRIDGE_URL}/reply`, {
    data: { text: replyText },
  });
  expect(postRes.status()).toBe(200);

  // ポーリングで取得
  const getRes = await request.get(`${BRIDGE_URL}/pending-reply`);
  expect(getRes.status()).toBe(200);
  const body = await getRes.json();
  expect(body.pending).toBe(true);
  expect(body.text).toBe(replyText);

  // 2回目のポーリングはクリア済み
  const getRes2 = await request.get(`${BRIDGE_URL}/pending-reply`);
  const body2 = await getRes2.json();
  expect(body2.pending).toBe(false);
});

test('POST /reply に text なしは400を返す', async ({ request }) => {
  const res = await request.post(`${BRIDGE_URL}/reply`, { data: {} });
  expect(res.status()).toBe(400);
});

test('GET /status でサーバー状態が確認できる', async ({ request }) => {
  const res = await request.get(`${BRIDGE_URL}/status`);
  expect(res.status()).toBe(200);
  const body = await res.json();
  expect(body.status).toBe('ok');
  // メッセージバッファが存在する（前のテストで POST 済み）
  expect(body.messageBuffer).toBeTruthy();
  expect(body.messageBuffer.messageCount).toBe(2);
});
