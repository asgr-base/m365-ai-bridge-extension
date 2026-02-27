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
// 最終検証: 2026-02-27 (teams.cloud.microsoft 新UI)
const SELECTORS = {
  // 個別メッセージのコンテナ
  messageContainer: '[data-tid="channel-pane-message"]',
  // メッセージの本文（コンテナ内）
  messageBody: '[data-tid="message-body"]',
  // 送信者名（id="author-{messageId}" を持つ span）
  senderName: 'span[id^="author-"]',
  // 送信者ヘッダー領域
  senderHeader: '[data-tid="post-message-subheader"], [data-tid="reply-message-header"]',
  // タイムスタンプ
  timestamp: '[data-tid="timestamp"]',
  // 現在のチャンネル名
  channelName: '[data-tid="channelTitle-text"]',
  // 現在のチャット相手
  chatTitle: '[data-tid="chat-title"]',
  // 返信フォーム（メッセージ入力欄）
  replyBox: '[data-tid="ckeditor"], [role="textbox"][contenteditable="true"]',
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

  // 新UI: channel-pane-message コンテナから取得
  const containers = document.querySelectorAll(SELECTORS.messageContainer);

  if (containers.length === 0) {
    // フォールバック: より広いセレクタで試みる
    return extractMessagesFallback();
  }

  const context = getCurrentContext();

  containers.forEach((container, index) => {
    if (index >= CONFIG.maxMessages) return;

    // メッセージ本文
    const bodyEl = container.querySelector(SELECTORS.messageBody);
    // 送信者名: span[id^="author-"] 内のテキスト
    const senderEl = container.querySelector(SELECTORS.senderName);
    // タイムスタンプ
    const timeEl = container.querySelector(SELECTORS.timestamp);

    // メッセージ ID: bodyEl の id 属性から "message-body-{id}" パターンで取得
    const rawId = bodyEl?.id || container.id || '';
    const messageId = rawId.replace(/^message-body-/, '') || null;
    // Deep link 用の数値 ID: "content-1770359698044" → "1770359698044"
    const numericId = messageId?.replace(/^content-/, '') || messageId;

    // 深リンク URL を構築
    const deepLink = buildMessageDeepLink(
      numericId,
      { threadId: context.threadId, groupId: context.groupId, tenantId: context.tenantId },
      context.channelName
    );

    messages.push({
      index,
      sender: senderEl?.textContent?.trim() || 'Unknown',
      body: bodyEl?.innerText?.trim() || '',
      timestamp: timeEl?.getAttribute('datetime') || timeEl?.textContent?.trim() || '',
      messageId,
      url: deepLink,
    });
  });

  return {
    context,
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
 * Teams のスレッドコンテキスト（threadId, groupId, tenantId）を DOM/URL から抽出する。
 * 複数のソースを試み、最初に見つかった値を返す。
 */
function extractTeamsThreadContext() {
  const ctx = { threadId: null, groupId: null, tenantId: null };

  // 1. URL ハッシュ・パラメータから取得
  //    例: https://teams.cloud.microsoft/v2/#/l/channel/19:xxx@thread.tacv2/General?groupId=yyy&tenantId=zzz
  try {
    const url = new URL(window.location.href);

    // クエリパラメータ
    ctx.groupId = ctx.groupId || url.searchParams.get('groupId');
    ctx.tenantId = ctx.tenantId || url.searchParams.get('tenantId');

    // ハッシュ内のパス部分をパースして threadId を探す
    const hashPath = url.hash.replace(/^#\/?/, '');
    const threadMatch = hashPath.match(/19:[a-zA-Z0-9._%-]+@thread\.[a-zA-Z0-9]+/);
    if (threadMatch) ctx.threadId = decodeURIComponent(threadMatch[0]);

    // ハッシュ内のクエリパラメータも試みる
    const hashQueryIdx = url.hash.indexOf('?');
    if (hashQueryIdx !== -1) {
      const hashQuery = new URLSearchParams(url.hash.slice(hashQueryIdx + 1));
      ctx.groupId = ctx.groupId || hashQuery.get('groupId');
      ctx.tenantId = ctx.tenantId || hashQuery.get('tenantId');
    }
  } catch {
    // URL パース失敗は無視
  }

  // 2. DOM の <a> href から `19:xxx@thread.xxx` を検索
  if (!ctx.threadId) {
    const links = document.querySelectorAll('a[href*="thread"]');
    for (const link of links) {
      const m = link.href.match(/19:[a-zA-Z0-9._%-]+@thread\.[a-zA-Z0-9]+/);
      if (m) { ctx.threadId = decodeURIComponent(m[0]); break; }
    }
  }

  // 3. DOM の data 属性から取得を試みる
  if (!ctx.threadId) {
    const el = document.querySelector('[data-threadid], [data-thread-id], [data-channel-id]');
    ctx.threadId = ctx.threadId
      || el?.dataset?.threadid
      || el?.dataset?.threadId
      || el?.dataset?.channelId
      || null;
  }

  // 4. ページ内スクリプトタグの JSON から groupId / tenantId を探す（Teams の埋め込み設定）
  if (!ctx.groupId || !ctx.tenantId) {
    const scripts = document.querySelectorAll('script:not([src])');
    for (const s of scripts) {
      const text = s.textContent;
      if (!text || text.length > 50000) continue;
      if (!ctx.groupId) {
        const m = text.match(/"groupId"\s*:\s*"([0-9a-f-]{36})"/i);
        if (m) ctx.groupId = m[1];
      }
      if (!ctx.tenantId) {
        const m = text.match(/"tenantId"\s*:\s*"([0-9a-f-]{36})"/i);
        if (m) ctx.tenantId = m[1];
      }
      if (ctx.groupId && ctx.tenantId) break;
    }
  }

  return ctx;
}

/**
 * Teams メッセージの深リンク URL を構築する。
 * threadId と messageId が取得できた場合のみ URL を返す。
 */
function buildMessageDeepLink(messageId, threadCtx, channelName) {
  const { threadId, groupId, tenantId } = threadCtx;
  if (!threadId || !messageId) return null;

  const params = new URLSearchParams();
  if (tenantId) params.set('tenantId', tenantId);
  if (groupId) params.set('groupId', groupId);
  params.set('parentMessageId', messageId);
  if (channelName) params.set('channelName', channelName);
  params.set('createdTime', messageId);

  return `https://teams.microsoft.com/l/message/${encodeURIComponent(threadId)}/${messageId}?${params}`;
}

/**
 * 現在開いているチャンネル・チャットのコンテキスト情報を取得
 */
function getCurrentContext() {
  const threadCtx = extractTeamsThreadContext();
  return {
    url: window.location.href,
    channelName: document.querySelector(SELECTORS.channelName)?.textContent?.trim() || null,
    chatTitle: document.querySelector(SELECTORS.chatTitle)?.textContent?.trim() || null,
    pageTitle: document.title,
    threadId: threadCtx.threadId,
    groupId: threadCtx.groupId,
    tenantId: threadCtx.tenantId,
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

// ========== DOM 構造調査 ==========

/**
 * Teams の DOM 構造を調査し、セレクタ調整に必要な情報を返す。
 * メッセージ要素の HTML スニペット、data-* 属性、クラス名を収集する。
 */
function inspectDom() {
  const results = {
    summary: {},
    dataTidElements: [],
    messageCandidate: [],
    senderCandidate: [],
    timestampCandidate: [],
    replyBoxCandidate: [],
    sampleHtml: [],
  };

  // 1. data-tid 属性を持つ全要素を収集
  const tidElements = document.querySelectorAll('[data-tid]');
  const tidMap = {};
  tidElements.forEach(el => {
    const tid = el.getAttribute('data-tid');
    tidMap[tid] = (tidMap[tid] || 0) + 1;
  });
  results.dataTidElements = Object.entries(tidMap)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 50)
    .map(([tid, count]) => ({ tid, count }));

  // 2. メッセージ候補を探す（チャンネル・チャット両対応）
  const msgPatterns = [
    // チャンネル系
    '[data-tid="channel-pane-message"]',
    // チャット系（候補）
    '[data-tid="chat-pane-message"]',
    '[data-tid="chat-item"]',
    '[data-tid="chat-message"]',
    '[data-tid*="chat-pane"]',
    '[data-tid*="chat-message"]',
    '[data-tid*="chat-item"]',
    // 汎用
    '[data-tid*="message"]',
    '[class*="message"]',
    '[class*="Message"]',
    '[data-message-id]',
    '[role="listitem"]',
    '[class*="chat-item"]',
    '[class*="ChatItem"]',
  ];
  msgPatterns.forEach(selector => {
    const els = document.querySelectorAll(selector);
    if (els.length > 0) {
      results.messageCandidate.push({
        selector,
        count: els.length,
        sampleClasses: els[0].className?.split(' ').slice(0, 5) || [],
        sampleDataAttrs: Array.from(els[0].attributes)
          .filter(a => a.name.startsWith('data-'))
          .map(a => `${a.name}="${a.value}"`)
          .slice(0, 5),
      });
    }
  });

  // 3. 送信者名の候補（チャット・チャンネル両対応）
  const senderPatterns = [
    // チャンネル系
    '[data-tid*="author"]',
    '[data-tid*="sender"]',
    '[data-tid*="display-name"]',
    'span[id^="author-"]',
    // チャット系（Fluent UI / aria）
    '[data-tid*="header"]',
    '[data-tid*="name"]',
    '[data-tid*="person"]',
    'span[title]:not([aria-hidden="true"])',
    'button[aria-label]',
    '[class*="fui-Persona"]',
    '[class*="fui-Text"]',
    '[class*="fui-Avatar"]',
    // 汎用
    '[class*="author"]',
    '[class*="sender"]',
    '[class*="displayName"]',
    '[class*="DisplayName"]',
    '[class*="name"][class*="fui"]',
  ];
  senderPatterns.forEach(selector => {
    const els = document.querySelectorAll(selector);
    if (els.length > 0) {
      results.senderCandidate.push({
        selector,
        count: els.length,
        samples: Array.from(els).slice(0, 3).map(e => e.textContent?.trim().slice(0, 30)),
      });
    }
  });

  // 4. タイムスタンプの候補
  const timePatterns = [
    'time',
    '[data-tid*="time"]',
    '[data-tid*="timestamp"]',
    '[class*="timestamp"]',
    '[class*="Timestamp"]',
    '[class*="time"]',
    '[datetime]',
  ];
  timePatterns.forEach(selector => {
    const els = document.querySelectorAll(selector);
    if (els.length > 0) {
      results.timestampCandidate.push({
        selector,
        count: els.length,
        samples: Array.from(els).slice(0, 3).map(e => ({
          text: e.textContent?.trim().slice(0, 30),
          datetime: e.getAttribute('datetime') || null,
        })),
      });
    }
  });

  // 5. 返信ボックスの候補
  const replyPatterns = [
    '[contenteditable="true"]',
    '[data-tid*="ckeditor"]',
    '[data-tid*="editor"]',
    '[role="textbox"]',
    '[class*="editor"]',
    '[class*="Editor"]',
  ];
  replyPatterns.forEach(selector => {
    const els = document.querySelectorAll(selector);
    if (els.length > 0) {
      results.replyBoxCandidate.push({
        selector,
        count: els.length,
        sampleTag: els[0].tagName,
        sampleClasses: els[0].className?.split(' ').slice(0, 5) || [],
      });
    }
  });

  // 6. メッセージらしき要素の HTML サンプル（最初の2件）
  // 最も有望なメッセージ候補の outerHTML を取得
  const bestMsgSelector = results.messageCandidate
    .sort((a, b) => b.count - a.count)[0]?.selector;
  if (bestMsgSelector) {
    const sampleEls = document.querySelectorAll(bestMsgSelector);
    Array.from(sampleEls).slice(0, 2).forEach((el, i) => {
      // 巨大になりすぎないようHTMLを切り詰め
      let html = el.outerHTML;
      if (html.length > 2000) html = html.slice(0, 2000) + '... [truncated]';
      results.sampleHtml.push({
        index: i,
        selector: bestMsgSelector,
        htmlLength: el.outerHTML.length,
        html,
      });
    });
  }

  results.summary = {
    dataTidCount: tidElements.length,
    uniqueTids: Object.keys(tidMap).length,
    messageCandidates: results.messageCandidate.length,
    senderCandidates: results.senderCandidate.length,
    url: window.location.href,
  };

  return results;
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

    case 'INSPECT_DOM': {
      const result = inspectDom();
      log('log', 'DOM構造調査完了:', result.summary);
      sendResponse({ success: true, data: result });
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

// ========== 自動プッシュ ==========

let lastPushHash = '';

/**
 * メッセージを抽出してブリッジサーバーへ自動送信する。
 * 前回送信時と内容が変わった場合のみ送信する。
 */
async function autoPush() {
  const result = extractMessages();
  // メッセージ件数 + 先頭/末尾の本文でハッシュを簡易生成
  const msgs = result.messages;
  const hash = `${msgs.length}:${msgs[0]?.body?.slice(0, 30) || ''}:${msgs[msgs.length - 1]?.body?.slice(0, 30) || ''}`;

  if (hash === lastPushHash) return; // 変更なし → スキップ

  lastPushHash = hash;
  const res = await sendToBridge('/messages', result);
  if (res) {
    log('log', `自動プッシュ: ${msgs.length}件送信`);
  }
}

// ========== 初期化 ==========

log('log', 'Teams コンテンツスクリプト起動 (Phase 1 PoC)');
log('log', 'URL:', window.location.href);

// ページ読み込み完了後に初回取得 + 定期プッシュ開始
window.addEventListener('load', () => {
  // 初回: DOM が安定するまで少し待つ
  setTimeout(() => {
    autoPush();
    // 15秒ごとに自動プッシュ（変更がなければスキップ）
    setInterval(autoPush, 15000);
  }, 3000);
});
