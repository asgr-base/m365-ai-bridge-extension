#!/usr/bin/env node
/**
 * M365 AI Bridge - Local Bridge Server
 *
 * Claude Code（またはローカルスクリプト）が Teams 情報にアクセスするための
 * ローカル HTTP サーバー。Chrome Extension との通信ハブとして機能する。
 *
 * 起動方法:
 *   node native/bridge-server.js
 *
 * エンドポイント:
 *   GET  /health           ヘルスチェック
 *   GET  /messages         最新の Teams メッセージを取得
 *   POST /messages         Extension からメッセージデータを受信・保存
 *   POST /reply            Extension に返信テキストを送信
 *   GET  /status           サーバー状態とバッファ情報を返す
 */

'use strict';

const http = require('http');
const PORT = parseInt(process.env.BRIDGE_PORT, 10) || 3765;

// ========== インメモリストア ==========
let messageBuffer = null;      // Extension から受信した最新メッセージ
let messageBufferAt = null;    // 受信時刻
let pendingReply = null;       // Extension に送信待ちの返信テキスト

// ========== ユーティリティ ==========

function log(...args) {
  console.log(`[${new Date().toISOString()}]`, ...args);
}

const ALLOWED_ORIGINS = [
  'https://teams.microsoft.com',
  'https://teams.cloud.microsoft',
];

function getCorsOrigin(req) {
  const origin = req.headers.origin;
  if (!origin) return ALLOWED_ORIGINS[0];
  if (ALLOWED_ORIGINS.includes(origin)) return origin;
  if (/^https:\/\/[a-z0-9-]+\.teams\.microsoft\.com$/.test(origin)) return origin;
  if (/^https:\/\/[a-z0-9-]+\.teams\.cloud\.microsoft$/.test(origin)) return origin;
  if (/^chrome-extension:\/\/[a-z]{32}$/.test(origin)) return origin;
  return null;
}

function jsonResponse(res, statusCode, data, corsOrigin) {
  const body = JSON.stringify(data, null, 2);
  const headers = {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };
  if (corsOrigin) headers['Access-Control-Allow-Origin'] = corsOrigin;
  res.writeHead(statusCode, headers);
  res.end(body);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', (chunk) => (body += chunk));
    req.on('end', () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch (err) {
        reject(new Error('Invalid JSON'));
      }
    });
    req.on('error', reject);
  });
}

// ========== ルーター ==========

const server = http.createServer(async (req, res) => {
  const { method, url } = req;
  const corsOrigin = getCorsOrigin(req);

  // CORS プリフライト
  if (method === 'OPTIONS') {
    const headers = {
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    };
    if (corsOrigin) headers['Access-Control-Allow-Origin'] = corsOrigin;
    res.writeHead(204, headers);
    return res.end();
  }

  log(`${method} ${url}`);

  try {
    // GET /health
    if (method === 'GET' && url === '/health') {
      return jsonResponse(res, 200, { status: 'ok', version: '0.1.0' }, corsOrigin);
    }

    // GET /status
    if (method === 'GET' && url === '/status') {
      return jsonResponse(res, 200, {
        status: 'ok',
        messageBuffer: messageBuffer ? {
          messageCount: messageBuffer.messages?.length || 0,
          context: messageBuffer.context,
          receivedAt: messageBufferAt,
        } : null,
        pendingReply: pendingReply ? pendingReply.slice(0, 50) + '...' : null,
      }, corsOrigin);
    }

    // GET /messages — Claude Code がメッセージを取得する
    if (method === 'GET' && url === '/messages') {
      if (!messageBuffer) {
        return jsonResponse(res, 404, {
          error: 'No messages available. Open Teams and click "メッセージを読み取る" in the extension popup.',
        }, corsOrigin);
      }
      return jsonResponse(res, 200, {
        ...messageBuffer,
        receivedAt: messageBufferAt,
      }, corsOrigin);
    }

    // POST /messages — Extension からメッセージデータを受信
    if (method === 'POST' && url === '/messages') {
      const data = await readBody(req);
      messageBuffer = data;
      messageBufferAt = new Date().toISOString();
      log(`メッセージ受信: ${data.messages?.length || 0}件 (${data.context?.channelName || data.context?.chatTitle || 'unknown'})`);
      return jsonResponse(res, 200, { success: true, messageCount: data.messages?.length || 0 }, corsOrigin);
    }

    // POST /reply — Claude Code が返信テキストを Extension に送信する
    if (method === 'POST' && url === '/reply') {
      const { text } = await readBody(req);
      if (!text) {
        return jsonResponse(res, 400, { error: 'text is required' }, corsOrigin);
      }
      pendingReply = text;
      log(`返信テキストを受信: ${text.slice(0, 80)}...`);

      // TODO: Phase 2 で Extension へのプッシュ実装（WebSocket等）
      // 現在は Extension がポーリングで取得する方式
      return jsonResponse(res, 200, { success: true, message: 'Reply queued. Extension will pick it up on next poll.' }, corsOrigin);
    }

    // GET /pending-reply — Extension が返信テキストをポーリングで取得
    if (method === 'GET' && url === '/pending-reply') {
      if (!pendingReply) {
        return jsonResponse(res, 200, { pending: false }, corsOrigin);
      }
      const reply = pendingReply;
      pendingReply = null; // 取得後にクリア
      return jsonResponse(res, 200, { pending: true, text: reply }, corsOrigin);
    }

    // 404
    return jsonResponse(res, 404, { error: `Not found: ${method} ${url}` }, corsOrigin);

  } catch (err) {
    log('エラー:', err.message);
    return jsonResponse(res, 500, { error: err.message }, corsOrigin);
  }
});

server.listen(PORT, '127.0.0.1', () => {
  log(`M365 AI Bridge サーバー起動: http://localhost:${PORT}`);
  log('');
  log('エンドポイント:');
  log(`  GET  http://localhost:${PORT}/health         — ヘルスチェック`);
  log(`  GET  http://localhost:${PORT}/messages       — Teams メッセージ取得`);
  log(`  POST http://localhost:${PORT}/reply          — 返信テキスト送信`);
  log(`  GET  http://localhost:${PORT}/status         — サーバー状態確認`);
  log('');
  log('Teams を開き、Extension ポップアップから「メッセージを読み取る」をクリックしてください。');
});

server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.error(`ポート ${PORT} は既に使用中です。既存のプロセスを終了してください。`);
    console.error(`  lsof -ti:${PORT} | xargs kill`);
  } else {
    console.error('サーバーエラー:', err.message);
  }
  process.exit(1);
});
