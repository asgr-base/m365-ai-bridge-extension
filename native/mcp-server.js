#!/usr/bin/env node
/**
 * M365 AI Bridge - MCP Server
 *
 * Claude Code と Chrome Extension を接続する MCP (Model Context Protocol) サーバー。
 * 1 プロセスで 2 つのインターフェースを提供する:
 *
 *   1. MCP stdio サーバー  → Claude Code がツールとして呼び出す (stdin/stdout)
 *   2. HTTP サーバー       → Chrome Extension がデータを投稿・ポーリングする (localhost:3765)
 *
 * 使い方:
 *   .mcp.json に登録して Claude Code から自動起動させる:
 *   {
 *     "mcpServers": {
 *       "teams": {
 *         "type": "stdio",
 *         "command": "node",
 *         "args": ["/path/to/native/mcp-server.js"]
 *       }
 *     }
 *   }
 */

'use strict';

const http = require('http');
const { McpServer } = require('@modelcontextprotocol/sdk/server/mcp.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const { z } = require('zod');

const fs = require('fs');
const path = require('path');
const https = require('https');

const HTTP_PORT = 3765;
const DEBUG_ENABLED = process.env.M365_DEBUG === '1';

// ========== インメモリストア ==========
let messageBuffer = null;   // Extension から受信した最新メッセージ
let messageBufferAt = null;
let pendingReply = null;    // Extension に送信待ちの返信テキスト
let graphToken = null;      // Extension から受信した Graph API トークン
let graphTokenAt = null;

// ========== CORS ホワイトリスト ==========

const ALLOWED_ORIGINS = [
  'https://teams.microsoft.com',
  'https://teams.cloud.microsoft',
];

function getCorsOrigin(req) {
  const origin = req.headers.origin;
  if (!origin) return ALLOWED_ORIGINS[0]; // Origin ヘッダーなし（同一オリジン/curl等）
  // 完全一致またはサブドメイン一致
  if (ALLOWED_ORIGINS.includes(origin)) return origin;
  if (/^https:\/\/[a-z0-9-]+\.teams\.microsoft\.com$/.test(origin)) return origin;
  if (/^https:\/\/[a-z0-9-]+\.teams\.cloud\.microsoft$/.test(origin)) return origin;
  // Chrome Extension origin (chrome-extension://xxx)
  if (/^chrome-extension:\/\/[a-z]{32}$/.test(origin)) return origin;
  return null; // 許可しない
}

// ========== HTTP サーバー（Chrome Extension 用） ==========

function startHttpServer() {
  const server = http.createServer(async (req, res) => {
    const { method, url } = req;
    const corsOrigin = getCorsOrigin(req);

    const jsonRes = (status, data) => {
      const body = JSON.stringify(data, null, 2);
      const headers = {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
      };
      if (corsOrigin) headers['Access-Control-Allow-Origin'] = corsOrigin;
      res.writeHead(status, headers);
      res.end(body);
    };

    const readBody = () => new Promise((resolve, reject) => {
      let body = '';
      req.on('data', c => (body += c));
      req.on('end', () => {
        try { resolve(body ? JSON.parse(body) : {}); }
        catch (e) { reject(new Error('Invalid JSON')); }
      });
      req.on('error', reject);
    });

    if (method === 'OPTIONS') {
      const headers = {
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
      };
      if (corsOrigin) headers['Access-Control-Allow-Origin'] = corsOrigin;
      res.writeHead(204, headers);
      return res.end();
    }

    try {
      if (method === 'GET' && url === '/health')
        return jsonRes(200, { status: 'ok', version: '0.2.0', mode: 'mcp' });

      if (method === 'GET' && url === '/status')
        return jsonRes(200, {
          status: 'ok',
          messageBuffer: messageBuffer ? {
            messageCount: messageBuffer.messages?.length || 0,
            context: messageBuffer.context,
            receivedAt: messageBufferAt,
          } : null,
          pendingReply: pendingReply ? pendingReply.slice(0, 50) + '...' : null,
          graphToken: graphToken ? {
            hasToken: true,
            length: graphToken.length,
            receivedAt: graphTokenAt,
          } : null,
        });

      // デバッグエンドポイント（M365_DEBUG=1 で有効化）
      if (DEBUG_ENABLED) {
        // デバッグ: メッセージバッファの全内容を返す
        if (method === 'GET' && url === '/debug/messages')
          return jsonRes(200, messageBuffer || { messages: [] });

        // デバッグ: Graph API テスト（/me を呼び出す）
        if (method === 'GET' && url === '/debug/graph-test') {
          if (!graphToken) return jsonRes(400, { error: 'No graph token' });
          try {
            const result = await graphApiGet('/v1.0/me', graphToken);
            return jsonRes(200, { success: true, user: { displayName: result.displayName } });
          } catch (err) {
            return jsonRes(500, { error: err.message });
          }
        }

        // デバッグ: ファイルダウンロード
        if (method === 'GET' && url.startsWith('/debug/download?')) {
          if (!graphToken) return jsonRes(400, { error: 'No graph token' });
          const params = new URL(`http://localhost${url}`).searchParams;
          const filename = params.get('filename');
          if (!filename) return jsonRes(400, { error: 'Missing ?filename= parameter' });
          try {
            const result = await downloadFileFromGraph(filename, {
              groupId: params.get('groupId') || undefined,
              chatType: params.get('chatType') || 'channel',
              downloadDir: params.get('dir') || undefined,
            });
            return jsonRes(200, {
              success: true,
              filePath: result.filePath,
              bytes: result.bytes,
              name: result.driveItem.name,
              id: result.driveItem.id,
              webUrl: result.driveItem.webUrl,
            });
          } catch (err) {
            return jsonRes(500, { error: err.message });
          }
        }

        // デバッグ: Graph Search API ファイル検索（全ドライブ横断）
        if (method === 'GET' && url.startsWith('/debug/search?')) {
          if (!graphToken) return jsonRes(400, { error: 'No graph token' });
          const params = new URL(`http://localhost${url}`).searchParams;
          const q = params.get('q');
          if (!q) return jsonRes(400, { error: 'Missing ?q= parameter' });
          try {
            const items = await searchFilesAcrossDrives(q, graphToken);
            return jsonRes(200, { count: items.length, files: items });
          } catch (err) {
            return jsonRes(500, { error: err.message });
          }
        }
      } // end DEBUG_ENABLED

      if (method === 'POST' && url === '/messages') {
        const data = await readBody();
        messageBuffer = data;
        messageBufferAt = new Date().toISOString();
        return jsonRes(200, { success: true, messageCount: data.messages?.length || 0 });
      }

      if (method === 'POST' && url === '/token') {
        const data = await readBody();
        if (data.token && typeof data.token === 'string') {
          graphToken = data.token;
          graphTokenAt = new Date().toISOString();
          process.stderr.write(`[M365 AI Bridge] Graph token received: length=${data.token.length}\n`);
          return jsonRes(200, { success: true, tokenLength: data.token.length });
        }
        return jsonRes(400, { error: 'Missing or invalid token field' });
      }

      if (method === 'GET' && url === '/pending-reply') {
        if (!pendingReply) return jsonRes(200, { pending: false });
        const reply = pendingReply;
        pendingReply = null;
        return jsonRes(200, { pending: true, text: reply });
      }

      jsonRes(404, { error: `Not found: ${method} ${url}` });
    } catch (err) {
      jsonRes(500, { error: err.message });
    }
  });

  let retryCount = 0;
  const MAX_RETRIES = 10;
  const RETRY_INTERVAL_MS = 3000;

  function tryListen() {
    server.listen(HTTP_PORT, '127.0.0.1', () => {
      process.stderr.write(`[M365 AI Bridge] HTTP server listening on localhost:${HTTP_PORT}\n`);
    });
  }

  server.on('error', (err) => {
    if (err.code === 'EADDRINUSE' && retryCount < MAX_RETRIES) {
      retryCount++;
      process.stderr.write(
        `[M365 AI Bridge] Port ${HTTP_PORT} in use, retry ${retryCount}/${MAX_RETRIES} in ${RETRY_INTERVAL_MS / 1000}s...\n`
      );
      setTimeout(tryListen, RETRY_INTERVAL_MS);
    } else if (err.code === 'EADDRINUSE') {
      process.stderr.write(`[M365 AI Bridge] Port ${HTTP_PORT} still in use after ${MAX_RETRIES} retries. HTTP server disabled.\n`);
    } else {
      process.stderr.write(`[M365 AI Bridge] HTTP error: ${err.message}\n`);
    }
  });

  tryListen();
  return server;
}

// ========== Graph API ヘルパー ==========

/**
 * Graph API に GET リクエストを送信する
 * @param {string} endpoint - Graph API エンドポイント（例: '/me/drive/root/search(q=...'）
 * @param {string} token - Bearer トークン
 * @returns {Promise<Object>} JSON レスポンス
 */
function graphApiGet(endpoint, token) {
  return new Promise((resolve, reject) => {
    const url = new URL(endpoint, 'https://graph.microsoft.com');
    const options = {
      hostname: url.hostname,
      path: url.pathname + url.search,
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json',
      },
    };

    const req = https.request(options, (res) => {
      let body = '';
      res.on('data', chunk => (body += chunk));
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try { resolve(JSON.parse(body)); }
          catch (e) { reject(new Error(`Invalid JSON response: ${body.slice(0, 200)}`)); }
        } else {
          reject(new Error(`Graph API ${res.statusCode}: ${body.slice(0, 300)}`));
        }
      });
    });

    req.on('error', reject);
    req.end();
  });
}

/**
 * Graph API に POST リクエストを送信する
 * @param {string} endpoint - Graph API エンドポイント
 * @param {string} token - Bearer トークン
 * @param {Object} body - リクエストボディ
 * @returns {Promise<Object>} JSON レスポンス
 */
function graphApiPost(endpoint, token, body) {
  return new Promise((resolve, reject) => {
    const url = new URL(endpoint, 'https://graph.microsoft.com');
    const bodyStr = JSON.stringify(body);
    const options = {
      hostname: url.hostname,
      path: url.pathname + url.search,
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Content-Length': Buffer.byteLength(bodyStr),
      },
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => (data += chunk));
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try { resolve(JSON.parse(data)); }
          catch (e) { reject(new Error(`Invalid JSON response: ${data.slice(0, 200)}`)); }
        } else {
          reject(new Error(`Graph API ${res.statusCode}: ${data.slice(0, 300)}`));
        }
      });
    });

    req.on('error', reject);
    req.write(bodyStr);
    req.end();
  });
}

/**
 * Graph Search API でファイルを検索する（全ドライブ横断）
 * @param {string} query - 検索クエリ
 * @param {string} token - Bearer トークン
 * @returns {Promise<Array>} 検索結果のドライブアイテム配列
 */
async function searchFilesAcrossDrives(query, token) {
  const searchBody = {
    requests: [{
      entityTypes: ['driveItem'],
      query: { queryString: query },
      from: 0,
      size: 10,
    }],
  };

  const result = await graphApiPost('/v1.0/search/query', token, searchBody);

  const items = [];
  for (const response of result.value || []) {
    for (const hit of response.hitsContainers?.[0]?.hits || []) {
      const resource = hit.resource;
      if (resource) {
        items.push({
          name: resource.name,
          id: resource.id,
          size: resource.size,
          webUrl: resource.webUrl,
          // Search API のレスポンスには @microsoft.graph.downloadUrl がない
          // driveId と itemId を使って別途取得が必要
          driveId: resource.parentReference?.driveId,
          siteId: resource.parentReference?.siteId,
        });
      }
    }
  }
  return items;
}

/**
 * URL からファイルをダウンロードしてローカルに保存する
 * @param {string} downloadUrl - ダウンロード URL（事前認証済み）
 * @param {string} destPath - 保存先パス
 * @returns {Promise<number>} ダウンロードしたバイト数
 */
function downloadFromUrl(downloadUrl, destPath) {
  return new Promise((resolve, reject) => {
    const parsedUrl = new URL(downloadUrl);
    const protocol = parsedUrl.protocol === 'https:' ? https : require('http');

    const req = protocol.get(downloadUrl, (res) => {
      // リダイレクト対応
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        downloadFromUrl(res.headers.location, destPath).then(resolve).catch(reject);
        return;
      }

      if (res.statusCode !== 200) {
        reject(new Error(`Download failed: HTTP ${res.statusCode}`));
        return;
      }

      const dir = path.dirname(destPath);
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }

      const fileStream = fs.createWriteStream(destPath);
      let bytes = 0;
      res.on('data', chunk => { bytes += chunk.length; });
      res.pipe(fileStream);
      fileStream.on('finish', () => {
        fileStream.close();
        resolve(bytes);
      });
      fileStream.on('error', reject);
    });

    req.on('error', reject);
  });
}

/**
 * Graph API でファイルを検索してダウンロードする
 * @param {string} filename - 検索するファイル名
 * @param {Object} options - オプション
 * @returns {Promise<{ filePath: string, bytes: number, driveItem: Object }>}
 */
async function downloadFileFromGraph(filename, options = {}) {
  const { groupId, chatType = 'channel', downloadDir } = options;

  if (!graphToken) {
    throw new Error('Graph トークンが未設定です。Teams を開いて Extension からトークンを送信してください。');
  }

  let item = null;

  // 方法1: groupId がある場合はグループドライブを検索
  if (groupId) {
    try {
      const searchResult = await graphApiGet(
        `/v1.0/groups/${groupId}/drive/root/search(q='${encodeURIComponent(filename)}')`,
        graphToken
      );
      if (searchResult.value?.length > 0) {
        item = searchResult.value.find(v => v.name === filename) || searchResult.value[0];
      }
    } catch (err) {
      process.stderr.write(`[M365 AI Bridge] Group drive search failed: ${err.message}\n`);
    }
  }

  // 方法2: Search API でクロスサイト検索（フォールバック）
  if (!item) {
    try {
      const searchItems = await searchFilesAcrossDrives(filename, graphToken);
      if (searchItems.length > 0) {
        const match = searchItems.find(s => s.name === filename) || searchItems[0];
        // Search API の結果には downloadUrl がないので、driveId + itemId で取得
        if (match.driveId && match.id) {
          item = await graphApiGet(`/v1.0/drives/${match.driveId}/items/${match.id}`, graphToken);
        }
      }
    } catch (err) {
      process.stderr.write(`[M365 AI Bridge] Search API failed: ${err.message}\n`);
    }
  }

  // 方法3: 個人 OneDrive を検索（最後のフォールバック）
  if (!item) {
    try {
      const searchResult = await graphApiGet(
        `/v1.0/me/drive/root/search(q='${encodeURIComponent(filename)}')`,
        graphToken
      );
      if (searchResult.value?.length > 0) {
        item = searchResult.value.find(v => v.name === filename) || searchResult.value[0];
      }
    } catch (err) {
      process.stderr.write(`[M365 AI Bridge] OneDrive search failed: ${err.message}\n`);
    }
  }

  if (!item) {
    throw new Error(`ファイルが見つかりません: "${filename}"（グループドライブ、Search API、OneDrive すべて検索済み）`);
  }

  const downloadUrl = item['@microsoft.graph.downloadUrl'];
  if (!downloadUrl) {
    throw new Error(`ダウンロード URL が取得できません: ${item.name} (id=${item.id})`);
  }

  // 保存先ディレクトリ
  const destDir = downloadDir || path.join(__dirname, '..', 'downloads');
  const destPath = path.join(destDir, item.name);

  const bytes = await downloadFromUrl(downloadUrl, destPath);

  return { filePath: destPath, bytes, driveItem: item };
}

// ========== MCP サーバー（Claude Code 用） ==========

async function startMcpServer() {
  const mcp = new McpServer({
    name: 'm365-ai-bridge',
    version: '0.2.0',
  });

  // ── ツール 1: メッセージ読み取り ──────────────────────────────
  mcp.tool(
    'teams_read_messages',
    '現在 Chrome Extension で開いている Teams チャンネル/チャットの最新メッセージを取得する。' +
    'Extension ポップアップで「メッセージを読み取る」を押した後に呼び出すこと。',
    {
      limit: z.number().min(1).max(100).optional()
        .describe('取得するメッセージの最大件数（デフォルト: 全件）'),
    },
    async ({ limit }) => {
      if (!messageBuffer) {
        return {
          content: [{
            type: 'text',
            text: 'エラー: メッセージがありません。\n' +
              '1. Chrome で teams.microsoft.com を開く\n' +
              '2. Extension ポップアップで「メッセージを読み取る」をクリック\n' +
              '3. 再度このツールを呼び出す',
          }],
          isError: true,
        };
      }

      const messages = limit
        ? messageBuffer.messages.slice(0, limit)
        : messageBuffer.messages;

      const ctx = messageBuffer.context;
      const lines = [
        `## Teams メッセージ`,
        `- チャンネル/チャット: ${ctx.channelName || ctx.chatTitle || ctx.pageTitle}`,
        `- 取得日時: ${messageBufferAt}`,
        `- 件数: ${messages.length}件`,
        ctx.threadId ? `- threadId: ${ctx.threadId}` : '- threadId: (未取得)',
        '',
        ...messages.map(m => {
          // メンション行（TO/CC）
          let mentionLine = '';
          if (m.mentions) {
            const parts = [];
            if (m.mentions.to?.length > 0) {
              parts.push(`TO: ${m.mentions.to.map(n => '@' + n).join(', ')}`);
            }
            if (m.mentions.cc?.length > 0) {
              parts.push(`CC: ${m.mentions.cc.map(n => '@' + n).join(', ')}`);
            }
            if (parts.length > 0) mentionLine = `\n${parts.join(' | ')}`;
          }

          // 返信ステータス行（チャンネルメッセージのみ）
          let replyLine = '';
          if (typeof m.replyCount === 'number') {
            if (m.replyCount > 0) {
              const senders = m.replySenders?.join(', ') || '';
              replyLine = `\n返信: ${m.replyCount}件${senders ? ' (' + senders + ')' : ''}`;
            } else {
              replyLine = '\n未返信';
            }
          }

          const urlLine = m.url ? `\n[メッセージへのリンク](${m.url})` : '';
          return `**[${m.index + 1}] ${m.sender}** (${m.timestamp || '不明'})${mentionLine}${replyLine}${urlLine}\n${m.body}`;
        }),
      ];

      return {
        content: [{
          type: 'text',
          text: lines.join('\n'),
        }],
      };
    }
  );

  // ── ツール 2: 返信テキストのキュー ────────────────────────────
  mcp.tool(
    'teams_queue_reply',
    'Extension 経由で Teams の返信フォームにテキストを挿入する。' +
    '送信はユーザーが手動で行う（AIが自動送信しない設計）。',
    {
      text: z.string().min(1).describe('挿入する返信テキスト'),
    },
    async ({ text }) => {
      pendingReply = text;

      return {
        content: [{
          type: 'text',
          text: `返信テキストをキューに登録しました。\n` +
            `Extension が次のポーリング時（数秒以内）に返信フォームへ自動挿入します。\n\n` +
            `--- 挿入テキスト ---\n${text}`,
        }],
      };
    }
  );

  // ── ツール 3: ステータス確認 ─────────────────────────────────
  mcp.tool(
    'teams_get_status',
    'M365 AI Bridge の現在の状態を確認する。Extension の接続状況、メッセージバッファ、返信待ち、Graph トークンを返す。',
    {},
    async () => {
      const lines = [
        '## M365 AI Bridge ステータス',
        '',
        `- HTTPサーバー: 起動中 (localhost:${HTTP_PORT})`,
        `- メッセージバッファ: ${messageBuffer
          ? `${messageBuffer.messages?.length || 0}件 (${messageBufferAt})`
          : 'なし（Extension からデータ未受信）'}`,
        `- 返信待ちテキスト: ${pendingReply ? `${pendingReply.slice(0, 50)}...` : 'なし'}`,
        `- Graph トークン: ${graphToken
          ? `保持中 (length=${graphToken.length}, 受信=${graphTokenAt})`
          : '未受信（Teams を開いて Extension が自動送信するのを待つ）'}`,
      ];

      return {
        content: [{ type: 'text', text: lines.join('\n') }],
      };
    }
  );

  // ── ツール 4: ファイルダウンロード ──────────────────────────────
  mcp.tool(
    'teams_download_file',
    'Teams メッセージに添付されているファイルを Graph API 経由でダウンロードする。' +
    'teams_read_messages で取得したメッセージの添付ファイル名を指定すること。' +
    'ダウンロードしたファイルは downloads/ に保存される。',
    {
      filename: z.string().min(1)
        .describe('ダウンロードするファイル名（例: report.xlsx）'),
      groupId: z.string().optional()
        .describe('Teams グループ（チーム）の ID。チャンネルファイルの場合に指定。省略時は個人 OneDrive を検索。'),
      chatType: z.enum(['channel', 'dm']).optional()
        .describe('チャットの種類。channel=チャンネル、dm=DM/グループチャット（デフォルト: channel）'),
      downloadDir: z.string().optional()
        .describe('保存先ディレクトリの絶対パス。省略時は Extension の downloads/ フォルダ。'),
    },
    async ({ filename, groupId, chatType, downloadDir }) => {
      try {
        const result = await downloadFileFromGraph(filename, { groupId, chatType, downloadDir });

        const lines = [
          '## ファイルダウンロード完了',
          '',
          `- ファイル名: ${result.driveItem.name}`,
          `- サイズ: ${result.bytes.toLocaleString()} bytes`,
          `- 保存先: ${result.filePath}`,
          `- ドライブアイテムID: ${result.driveItem.id}`,
          result.driveItem.webUrl ? `- SharePoint URL: ${result.driveItem.webUrl}` : '',
          '',
          `Read tool で \`${result.filePath}\` を読み取ってください。`,
        ].filter(Boolean);

        return {
          content: [{ type: 'text', text: lines.join('\n') }],
        };
      } catch (err) {
        return {
          content: [{
            type: 'text',
            text: `ファイルダウンロードエラー: ${err.message}\n\n` +
              'トラブルシューティング:\n' +
              '1. Teams を Chrome で開いているか確認\n' +
              '2. Extension ポップアップで「トークン調査」を実行し、Graph Token が検出されるか確認\n' +
              '3. groupId が正しいか確認（チャンネルファイルの場合）',
          }],
          isError: true,
        };
      }
    }
  );

  // ── stdio トランスポートで接続 ────────────────────────────────
  const transport = new StdioServerTransport();
  await mcp.connect(transport);

  process.stderr.write('[M365 AI Bridge] MCP server connected via stdio\n');
}

// ========== 起動時クリーンアップ ==========

/**
 * port 3765 を保持しているプロセスを終了する。
 * セッション再起動で古いプロセスが port を占有する問題を防止。
 */
function killPortHolder() {
  try {
    const { execSync } = require('child_process');
    const myPid = process.pid;
    // lsof で port を保持している PID を特定
    const output = execSync(
      `lsof -ti :${HTTP_PORT} 2>/dev/null || true`,
      { encoding: 'utf-8' }
    ).trim();

    if (!output) return;

    const pids = output.split('\n').map(Number).filter(pid => pid && pid !== myPid);
    for (const pid of pids) {
      try {
        process.kill(pid, 'SIGTERM');
        process.stderr.write(`[M365 AI Bridge] Killed port holder PID ${pid} on port ${HTTP_PORT}\n`);
      } catch {
        // プロセスが既に終了している場合は無視
      }
    }
  } catch {
    // lsof が使えない環境では無視
  }
}

// ========== 起動 ==========

killPortHolder();
startHttpServer();
startMcpServer().catch((err) => {
  process.stderr.write(`[M365 AI Bridge] MCP startup error: ${err.message}\n`);
  process.exit(1);
});
