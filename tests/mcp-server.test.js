/**
 * MCP サーバーテスト
 *
 * MCP TypeScript SDK のクライアントを使い、
 * stdio 経由でサーバーと通信してツール呼び出しを検証する。
 */

const { test, expect } = require('@playwright/test');
const { Client } = require('@modelcontextprotocol/sdk/client/index.js');
const { StdioClientTransport } = require('@modelcontextprotocol/sdk/client/stdio.js');
const path = require('path');

const SERVER_PATH = path.join(__dirname, '../native/mcp-server.js');
const BRIDGE_URL = 'http://localhost:3765';

// ── MCP クライアントのセットアップ ──────────────────────────────

let client;

test.beforeAll(async () => {
  const transport = new StdioClientTransport({
    command: 'node',
    args: [SERVER_PATH],
  });

  client = new Client({ name: 'test-client', version: '1.0.0' });
  await client.connect(transport);

  // HTTP サーバーが起動するまで待機
  await new Promise(resolve => setTimeout(resolve, 800));
});

test.afterAll(async () => {
  await client.close();
});

// ── ツール一覧の確認 ─────────────────────────────────────────

test('ツール一覧に3つのツールが含まれる', async () => {
  const { tools } = await client.listTools();
  const names = tools.map(t => t.name);

  expect(names).toContain('teams_read_messages');
  expect(names).toContain('teams_queue_reply');
  expect(names).toContain('teams_get_status');
});

test('各ツールに description が設定されている', async () => {
  const { tools } = await client.listTools();
  for (const tool of tools) {
    expect(tool.description).toBeTruthy();
  }
});

// ── teams_get_status ─────────────────────────────────────────

test('teams_get_status: 初期状態でメッセージバッファなしを返す', async () => {
  const result = await client.callTool({
    name: 'teams_get_status',
    arguments: {},
  });

  expect(result.isError).toBeFalsy();
  const text = result.content[0].text;
  expect(text).toContain('M365 AI Bridge ステータス');
  expect(text).toContain('localhost:3765');
  expect(text).toContain('なし');
});

// ── teams_read_messages ───────────────────────────────────────

test('teams_read_messages: データなし時はエラーメッセージを返す', async () => {
  const result = await client.callTool({
    name: 'teams_read_messages',
    arguments: {},
  });

  expect(result.isError).toBe(true);
  const text = result.content[0].text;
  expect(text).toContain('エラー');
  expect(text).toContain('メッセージを読み取る');
});

test('teams_read_messages: Extension からデータ受信後にメッセージを返す', async () => {
  // Extension の代わりに HTTP でデータを投稿
  const payload = {
    context: { channelName: 'General', chatTitle: null, pageTitle: 'Teams', url: 'https://teams.microsoft.com/...' },
    messages: [
      {
        index: 0, sender: '田中 太郎', body: 'MCPテストのメッセージです', timestamp: '2026-02-26T10:00:00Z',
        mentions: { to: ['山田 花子'], cc: ['中村 太一'] },
        replyCount: 1, replySenders: ['山田 花子'],
      },
      {
        index: 1, sender: '山田 花子', body: 'ご確認ありがとうございます', timestamp: '2026-02-26T10:01:00Z',
        replyCount: 0,
      },
    ],
    extractedAt: new Date().toISOString(),
    method: 'primary',
  };

  // HTTP POST で Extension からのデータを模擬
  const res = await fetch(`${BRIDGE_URL}/messages`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  expect(res.ok).toBe(true);

  // MCP ツール経由でメッセージ取得
  const result = await client.callTool({
    name: 'teams_read_messages',
    arguments: {},
  });

  expect(result.isError).toBeFalsy();
  const text = result.content[0].text;
  expect(text).toContain('Teams メッセージ');
  expect(text).toContain('General');
  expect(text).toContain('田中 太郎');
  expect(text).toContain('MCPテストのメッセージです');
  expect(text).toContain('山田 花子');

  // メンション行の検証
  expect(text).toContain('TO: @山田 花子');
  expect(text).toContain('CC: @中村 太一');

  // 返信ステータス行の検証
  expect(text).toContain('返信: 1件 (山田 花子)');
  expect(text).toContain('未返信');
});

test('teams_read_messages: limit パラメータで件数を絞れる', async () => {
  const result = await client.callTool({
    name: 'teams_read_messages',
    arguments: { limit: 1 },
  });

  expect(result.isError).toBeFalsy();
  const text = result.content[0].text;
  // 田中のメッセージは含む（メンション行に「山田 花子」が出るのは正常）
  expect(text).toContain('田中 太郎');
  expect(text).toContain('MCPテストのメッセージです');
  // 山田の2番目のメッセージ本文は含まない
  expect(text).not.toContain('ご確認ありがとうございます');
});

// ── teams_queue_reply ─────────────────────────────────────────

test('teams_queue_reply: 返信テキストをキューに登録できる', async () => {
  const replyText = 'はい、13時に参加します。よろしくお願いします。';

  const result = await client.callTool({
    name: 'teams_queue_reply',
    arguments: { text: replyText },
  });

  expect(result.isError).toBeFalsy();
  const text = result.content[0].text;
  expect(text).toContain('キューに登録');
  expect(text).toContain(replyText);
});

test('teams_queue_reply 後に Extension が /pending-reply でテキストを取得できる', async () => {
  const replyText = 'Extension ポーリングテスト用メッセージ';

  await client.callTool({
    name: 'teams_queue_reply',
    arguments: { text: replyText },
  });

  // Extension の代わりにポーリング
  const res = await fetch(`${BRIDGE_URL}/pending-reply`);
  const body = await res.json();
  expect(body.pending).toBe(true);
  expect(body.text).toBe(replyText);

  // 2回目はクリア済み
  const res2 = await fetch(`${BRIDGE_URL}/pending-reply`);
  const body2 = await res2.json();
  expect(body2.pending).toBe(false);
});

// ── teams_get_status（メッセージ受信後） ───────────────────────

test('teams_get_status: メッセージ受信後はバッファ情報を含む', async () => {
  // テスト間でバッファがクリアされる場合に備え、自分でデータを投入
  await fetch(`${BRIDGE_URL}/messages`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      context: { channelName: 'StatusTest', chatTitle: null, pageTitle: 'Teams' },
      messages: [
        { index: 0, sender: 'A', body: 'msg1', timestamp: '' },
        { index: 1, sender: 'B', body: 'msg2', timestamp: '' },
      ],
      extractedAt: new Date().toISOString(),
      method: 'primary',
    }),
  });

  const result = await client.callTool({
    name: 'teams_get_status',
    arguments: {},
  });

  expect(result.isError).toBeFalsy();
  const text = result.content[0].text;
  expect(text).toContain('2件');
});
