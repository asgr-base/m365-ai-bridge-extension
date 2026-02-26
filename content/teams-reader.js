/**
 * M365 AI Bridge - Teams Content Script
 *
 * Teams の DOM からメッセージ・チャンネル情報を取得し、
 * バックグラウンドサービスワーカー経由でローカル HTTP サーバーに送信する。
 *
 * Phase 1 PoC: DOM読み取りとコンソール出力
 */

'use strict';

// ========== 設定 ==========
const CONFIG = {
  // ローカルブリッジサーバーのURL（Claude Code側で起動）
  bridgeUrl: 'http://localhost:3765',
  // メッセージ取得の最大件数
  maxMessages: 50,
  // DOM監視のデバウンス時間（ms）
  debounceMs: 1000,
};

// ========== Teams DOM セレクタ ==========
// ※ Teams の UI 更新で壊れる可能性あり。定期的に検証が必要。
const SELECTORS = {
  // チャンネルメッセージ一覧
  channelMessages: '[data-tid="message-body"]',
  // チャットメッセージ一覧（1:1・グループチャット）
  chatMessages: '[data-tid="chat-message"]',
  // メッセージの本文
  messageBody: '[data-tid="message-body-content"]',
  // 送信者名
  senderName: '[data-tid="message-author-name"]',
  // タイムスタンプ
  timestamp: 'time[data-tid]',
  // 現在のチャンネル名
  channelName: '[data-tid="channel-name"]',
  // 現在のチャット相手
  chatTitle: '[data-tid="chat-title"]',
  // 返信フォーム（メッセージ入力欄）
  replyBox: '[data-tid="ckeditor"]',
  // 送信ボタン
  sendButton: '[data-tid="sendMessageCommands-send"]',
};

// ========== ユーティリティ ==========

function log(level, ...args) {
  const prefix = '[M365 AI Bridge]';
  if (level === 'error') console.error(prefix, ...args);
  else if (level === 'warn') console.warn(prefix, ...args);
  else console.log(prefix, ...args);
}

function debounce(fn, ms) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), ms);
  };
}

// ========== メッセージ取得 ==========

/**
 * 現在表示されている Teams メッセージを DOM から取得する
 * @returns {Object} { context, messages }
 */
function extractMessages() {
  const messages = [];

  // チャンネルメッセージを探す
  const msgElements = document.querySelectorAll(SELECTORS.channelMessages);

  if (msgElements.length === 0) {
    // フォールバック: より広いセレクタで試みる
    return extractMessagesFallback();
  }

  msgElements.forEach((el, index) => {
    if (index >= CONFIG.maxMessages) return;

    const bodyEl = el.querySelector(SELECTORS.messageBody) || el;
    const senderEl = el.closest('[data-track-action-scenario]')
      ?.querySelector(SELECTORS.senderName);
    const timeEl = el.closest('[data-track-action-scenario]')
      ?.querySelector(SELECTORS.timestamp);

    messages.push({
      index,
      sender: senderEl?.textContent?.trim() || 'Unknown',
      body: bodyEl?.innerText?.trim() || '',
      timestamp: timeEl?.getAttribute('datetime') || timeEl?.textContent?.trim() || '',
      elementId: el.id || null,
    });
  });

  return {
    context: getCurrentContext(),
    messages,
    extractedAt: new Date().toISOString(),
    method: 'primary',
  };
}

/**
 * フォールバック: 汎用的なセレクタでメッセージを取得する
 */
function extractMessagesFallback() {
  const messages = [];

  // Teams は aria-label や data 属性でメッセージを識別することが多い
  const candidates = [
    ...document.querySelectorAll('[class*="message"][class*="body"]'),
    ...document.querySelectorAll('[data-message-id]'),
    ...document.querySelectorAll('[id*="message"]'),
  ];

  // 重複排除
  const seen = new Set();
  candidates.forEach((el) => {
    const key = el.textContent?.trim().slice(0, 50);
    if (!key || seen.has(key) || key.length < 5) return;
    seen.add(key);

    if (messages.length >= CONFIG.maxMessages) return;

    messages.push({
      index: messages.length,
      sender: 'Unknown',
      body: el.innerText?.trim() || '',
      timestamp: '',
      elementId: el.id || null,
    });
  });

  return {
    context: getCurrentContext(),
    messages,
    extractedAt: new Date().toISOString(),
    method: 'fallback',
  };
}

/**
 * 現在開いているチャンネル・チャットのコンテキスト情報を取得
 */
function getCurrentContext() {
  return {
    url: window.location.href,
    channelName: document.querySelector(SELECTORS.channelName)?.textContent?.trim() || null,
    chatTitle: document.querySelector(SELECTORS.chatTitle)?.textContent?.trim() || null,
    pageTitle: document.title,
  };
}

// ========== 返信フォーム操作 ==========

/**
 * 返信フォームにテキストを入力する（Claude の生成した下書きを挿入）
 * @param {string} text - 挿入するテキスト
 * @returns {boolean} 成功したかどうか
 */
function insertReply(text) {
  const replyBox = document.querySelector(SELECTORS.replyBox);
  if (!replyBox) {
    log('warn', '返信フォームが見つかりません');
    return false;
  }

  // contenteditable への入力
  replyBox.focus();

  // execCommand は非推奨だが Teams の contenteditable では依然有効なことが多い
  const success = document.execCommand('insertText', false, text);
  if (!success) {
    // フォールバック: clipboard API 経由
    replyBox.textContent = text;
    replyBox.dispatchEvent(new InputEvent('input', { bubbles: true }));
  }

  log('log', '返信テキストを挿入しました');
  return true;
}

// ========== ブリッジサーバー通信 ==========

/**
 * ローカルブリッジサーバーにデータを送信する
 * @param {string} endpoint - APIエンドポイント
 * @param {Object} data - 送信するデータ
 */
async function sendToBridge(endpoint, data) {
  try {
    const response = await fetch(`${CONFIG.bridgeUrl}${endpoint}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data),
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    return await response.json();
  } catch (err) {
    // サーバーが未起動の場合は警告のみ（エラーは無視）
    if (err.message.includes('Failed to fetch')) {
      log('warn', 'ブリッジサーバーに接続できません（未起動の可能性）:', CONFIG.bridgeUrl);
    } else {
      log('error', 'ブリッジサーバーエラー:', err.message);
    }
    return null;
  }
}

// ========== コマンドハンドラ ==========

/**
 * バックグラウンドからのメッセージを処理する
 */
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  log('log', 'コマンド受信:', request.command);

  switch (request.command) {
    case 'READ_MESSAGES': {
      const result = extractMessages();
      log('log', `メッセージ取得: ${result.messages.length}件`, result.context);
      sendResponse({ success: true, data: result });
      break;
    }

    case 'INSERT_REPLY': {
      const success = insertReply(request.text);
      sendResponse({ success });
      break;
    }

    case 'PING': {
      sendResponse({ success: true, status: 'active', url: window.location.href });
      break;
    }

    default:
      log('warn', '不明なコマンド:', request.command);
      sendResponse({ success: false, error: 'Unknown command' });
  }

  // 非同期レスポンスを許可
  return true;
});

// ========== 初期化 ==========

log('log', 'Teams コンテンツスクリプト起動 (Phase 1 PoC)');
log('log', 'URL:', window.location.href);

// ページ読み込み完了後に初回取得を試みる
window.addEventListener('load', () => {
  setTimeout(() => {
    const result = extractMessages();
    log('log', '初回メッセージ取得:', result);

    // ブリッジサーバーへの接続試行（未起動でも問題なし）
    sendToBridge('/messages', result).then((res) => {
      if (res) log('log', 'ブリッジサーバーへ送信成功:', res);
    });
  }, 2000);
});
