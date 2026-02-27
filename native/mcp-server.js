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

const HTTP_PORT = 3765;

// ========== インメモリストア ==========
let messageBuffer = null;   // Extension から受信した最新メッセージ
let messageBufferAt = null;
let pendingReply = null;    // Extension に送信待ちの返信テキスト

// ========== HTTP サーバー（Chrome Extension 用） ==========

function startHttpServer() {
  const server = http.createServer(async (req, res) => {
    const { method, url } = req;

    const jsonRes = (status, data) => {
      const body = JSON.stringify(data, null, 2);
      res.writeHead(status, {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
      });
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
      res.writeHead(204, { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Content-Type' });
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
        });

      if (method === 'POST' && url === '/messages') {
        const data = await readBody();
        messageBuffer = data;
        messageBufferAt = new Date().toISOString();
        return jsonRes(200, { success: true, messageCount: data.messages?.length || 0 });
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
        '',
        ...messages.map(m =>
          `**[${m.index + 1}] ${m.sender}** (${m.timestamp || '不明'})\n${m.body}`
        ),
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
    'M365 AI Bridge の現在の状態を確認する。Extension の接続状況、メッセージバッファ、返信待ちを返す。',
    {},
    async () => {
      const status = {
        httpServer: `起動中 (localhost:${HTTP_PORT})`,
        messageBuffer: messageBuffer ? {
          messageCount: messageBuffer.messages?.length || 0,
          context: messageBuffer.context,
          receivedAt: messageBufferAt,
        } : null,
        pendingReply: pendingReply ? `${pendingReply.slice(0, 50)}...` : null,
      };

      const lines = [
        '## M365 AI Bridge ステータス',
        '',
        `- HTTPサーバー: ${status.httpServer}`,
        `- メッセージバッファ: ${status.messageBuffer
          ? `${status.messageBuffer.messageCount}件 (${status.messageBuffer.receivedAt})`
          : 'なし（Extension からデータ未受信）'}`,
        `- 返信待ちテキスト: ${status.pendingReply || 'なし'}`,
      ];

      return {
        content: [{ type: 'text', text: lines.join('\n') }],
      };
    }
  );

  // ── stdio トランスポートで接続 ────────────────────────────────
  const transport = new StdioServerTransport();
  await mcp.connect(transport);

  process.stderr.write('[M365 AI Bridge] MCP server connected via stdio\n');
}

// ========== 起動 ==========

startHttpServer();
startMcpServer().catch((err) => {
  process.stderr.write(`[M365 AI Bridge] MCP startup error: ${err.message}\n`);
  process.exit(1);
});
