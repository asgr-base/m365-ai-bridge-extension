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
  // === チャンネル用 ===
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

  // === DM / グループチャット用 ===
  // チャットアイテム外枠（メッセージごとの wrapper）
  dmChatItem: '[data-tid="chat-pane-item"]',
  // チャットメッセージ本文（id="message-body-{timestamp}"）
  dmMessageBody: '[data-tid="chat-pane-message"]',
  // 送信者名（DM 専用。チャンネルの span[id^="author-"] に相当）
  dmSenderName: '[data-tid="message-author-name"]',

  // === 共通 ===
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

// ========== ファイル添付抽出 ==========

/**
 * メッセージコンテナ内のファイル添付情報を抽出する
 * @param {Element} container - メッセージのコンテナ要素
 * @returns {Array<{name: string}>} 添付ファイル情報の配列
 */
function extractAttachments(container) {
  const fileRoots = container.querySelectorAll('[data-tid="file-preview-root"]');
  if (fileRoots.length === 0) return [];

  const attachments = [];
  fileRoots.forEach(root => {
    // 1. textContent の1行目からファイル名を取得（最もクリーン）
    const textName = root.textContent?.trim().split('\n')[0]?.trim();

    // 2. button の aria-label からフォールバック
    //    「ファイル XXX の画像プレビュー」形式の場合はファイル名だけ抽出
    let ariaName = null;
    if (!textName) {
      const btn = root.querySelector('button[aria-label]');
      const raw = btn?.getAttribute('aria-label')?.trim();
      if (raw) {
        const m = raw.match(/^ファイル\s+(.+?)\s+の画像プレビュー$/);
        ariaName = m ? m[1] : raw;
      }
    }

    const name = textName || ariaName || null;
    if (name) {
      attachments.push({ name });
    }
  });

  return attachments;
}

// ========== メッセージ取得 ==========

/**
 * DM / グループチャット向け抽出
 * chat-pane-item コンテナ + message-author-name セレクタを使用
 */
function extractDMMessages() {
  const items = document.querySelectorAll(SELECTORS.dmChatItem);
  if (items.length === 0) return null;

  const context = getCurrentContext();
  const messages = [];

  items.forEach((item) => {
    if (messages.length >= CONFIG.maxMessages) return;

    const bodyEl = item.querySelector(SELECTORS.dmMessageBody);
    if (!bodyEl) return; // メッセージ本文のないアイテム（日付区切り等）をスキップ

    const senderEl = item.querySelector(SELECTORS.dmSenderName);
    const timeEl = item.querySelector('[datetime]');

    // messageId: "message-body-{timestamp}" → timestamp
    const rawId = bodyEl.id || '';
    const messageId = rawId.replace(/^message-body-/, '') || null;

    const deepLink = buildMessageDeepLink(
      messageId,
      { threadId: context.threadId, groupId: context.groupId, tenantId: context.tenantId },
      null
    );

    const attachments = extractAttachments(item);

    messages.push({
      index: messages.length,
      sender: senderEl?.textContent?.trim() || 'Unknown',
      body: bodyEl.innerText?.trim() || '',
      timestamp: timeEl?.getAttribute('datetime') || timeEl?.textContent?.trim() || '',
      messageId,
      url: deepLink,
      ...(attachments.length > 0 && { attachments }),
    });
  });

  if (messages.length === 0) return null;
  return {
    context,
    messages,
    extractedAt: new Date().toISOString(),
    method: 'dm-chat-pane',
  };
}

function extractMessages() {
  const messages = [];

  // DM / グループチャット向け: chat-pane-item コンテナを優先
  // （Teams SPA がチャンネル DOM をキャッシュしていても chat-pane-item は DM 専用）
  const dmResult = extractDMMessages();
  if (dmResult) return dmResult;

  // チャンネル向け: channel-pane-message コンテナから取得
  const containers = document.querySelectorAll(SELECTORS.messageContainer);

  if (containers.length === 0) {
    // 最終フォールバック
    return extractMessagesFallback();
  }

  const context = getCurrentContext();

  containers.forEach((container, index) => {
    if (index >= CONFIG.maxMessages) return;

    // メッセージ本文
    const bodyEl = container.querySelector(SELECTORS.messageBody);
    // タイムスタンプ
    const timeEl = container.querySelector(SELECTORS.timestamp);

    // 送信者名: 複数パターンを順次試みる（チャンネル・DM両対応）
    const senderEl = container.querySelector(SELECTORS.senderName)       // span[id^="author-"]  チャンネル用
      || container.querySelector('[data-tid="author-name"]')             // DM 新UI候補
      || container.querySelector('[data-tid*="author"]')                 // author を含む data-tid
      || container.querySelector('[class*="author-name"]')               // クラス名ベース
      || container.querySelector('[class*="displayName"]');              // Fluent UI displayName
    // アバターボタンの aria-label からフォールバック（DM で名前が button に入る場合）
    const avatarAriaLabel = !senderEl
      ? container.querySelector('button[aria-label]:not([aria-label=""])')?.getAttribute('aria-label')
      : null;

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

    const attachments = extractAttachments(container);

    messages.push({
      index,
      sender: senderEl?.textContent?.trim() || avatarAriaLabel || 'Unknown',
      body: bodyEl?.innerText?.trim() || '',
      timestamp: timeEl?.getAttribute('datetime') || timeEl?.textContent?.trim() || '',
      messageId,
      url: deepLink,
      ...(attachments.length > 0 && { attachments }),
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
 * DM チャット向け: [data-tid="message-body"] から親を辿って送信者を探す
 * @param {Element[]|NodeList} [bodyElements] - 対象の message-body 要素（省略時は全件取得）
 */
function extractMessagesByBody(bodyElements) {
  const messageBodies = bodyElements || document.querySelectorAll(SELECTORS.messageBody);
  if (!messageBodies || messageBodies.length === 0) return null;

  const context = getCurrentContext();
  const messages = [];

  messageBodies.forEach((bodyEl, index) => {
    if (index >= CONFIG.maxMessages) return;

    // bodyEl から親を最大6階層辿り、span[id^="author-"] を探す
    let senderEl = null;
    let cur = bodyEl.parentElement;
    for (let depth = 0; depth < 6 && cur; depth++) {
      senderEl = cur.querySelector(SELECTORS.senderName);
      if (senderEl) break;
      cur = cur.parentElement;
    }

    // タイムスタンプ: 同じ親階層内で探す
    const timeEl = cur?.querySelector(SELECTORS.timestamp) || null;

    // messageId: bodyEl の id 属性から取得
    const rawId = bodyEl.id || '';
    const messageId = rawId.replace(/^message-body-/, '') || null;
    const numericId = messageId?.replace(/^content-/, '') || messageId;

    const deepLink = buildMessageDeepLink(
      numericId,
      { threadId: context.threadId, groupId: context.groupId, tenantId: context.tenantId },
      context.channelName
    );

    messages.push({
      index,
      sender: senderEl?.textContent?.trim() || 'Unknown',
      body: bodyEl.innerText?.trim() || '',
      timestamp: timeEl?.getAttribute('datetime') || timeEl?.textContent?.trim() || '',
      messageId,
      url: deepLink,
    });
  });

  return {
    context,
    messages,
    extractedAt: new Date().toISOString(),
    method: 'dm-body-traverse',
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
    fileCandidate: [],
    fileSampleHtml: [],
    dmIdElements: [],
    iframes: [],
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

  // 7. DM メッセージ候補（id*="message" / data-message-id の要素を調査）
  const dmIdCandidates = document.querySelectorAll('[id*="message"]:not([data-tid="message-body"])');
  results.dmIdElements = Array.from(dmIdCandidates)
    .filter(el => el.textContent?.trim().length > 5)
    .slice(0, 15)
    .map(el => ({
      id: el.id?.slice(0, 60),
      tag: el.tagName,
      dataTid: el.getAttribute('data-tid') || null,
      parentDataTid: el.parentElement?.getAttribute('data-tid') || null,
      textSample: el.textContent?.trim().slice(0, 60),
    }));

  // 8. ファイル添付候補の調査
  const filePatterns = [
    '[data-tid="file-preview-root"]',
    '[data-tid="file-attachment-grid"]',
    '[data-tid*="file-preview"]',
    '[data-tid*="file-attachment"]',
    '[data-tid*="file-card"]',
    'a[href*="sharepoint.com"]',
    'a[href*="1drv.ms"]',
    'a[href*=".sharepoint.com"]',
    '[class*="file-preview"]',
    '[class*="FileCard"]',
    '[class*="AttachmentCard"]',
  ];
  results.fileCandidate = [];
  filePatterns.forEach(selector => {
    try {
      const els = document.querySelectorAll(selector);
      if (els.length > 0) {
        results.fileCandidate.push({
          selector,
          count: els.length,
          samples: Array.from(els).slice(0, 3).map(e => ({
            tag: e.tagName,
            href: e.href || e.querySelector('a')?.href || null,
            text: e.textContent?.trim().slice(0, 80),
            dataTid: e.getAttribute('data-tid') || null,
            ariaLabel: e.getAttribute('aria-label')?.slice(0, 80) || null,
          })),
        });
      }
    } catch { /* invalid selector */ }
  });

  // 9. ファイル添付の詳細構造（最大3件）
  const fileRoots = document.querySelectorAll('[data-tid="file-preview-root"]');
  results.fileDetail = Array.from(fileRoots).slice(0, 3).map(root => {
    const imgs = root.querySelectorAll('img');
    const buttons = root.querySelectorAll('button');
    // href / data-url / src を持つ全要素
    const allChildren = root.querySelectorAll('*');
    const urlEls = Array.from(allChildren).filter(el =>
      el.getAttribute('href') || el.getAttribute('data-url') ||
      el.getAttribute('data-href') || el.getAttribute('data-src') ||
      (el.tagName === 'IMG' && el.src)
    );
    return {
      text: root.textContent?.trim().slice(0, 80),
      ariaLabel: root.getAttribute('aria-label')?.slice(0, 80),
      role: root.getAttribute('role'),
      imgs: Array.from(imgs).map(img => ({
        src: img.src?.slice(0, 150),
        alt: img.alt?.slice(0, 60),
      })),
      buttons: Array.from(buttons).slice(0, 3).map(btn => ({
        ariaLabel: btn.getAttribute('aria-label')?.slice(0, 80),
        text: btn.textContent?.trim().slice(0, 40),
      })),
      urlElements: urlEls.slice(0, 5).map(el => ({
        tag: el.tagName,
        href: (el.getAttribute('href') || '')?.slice(0, 150),
        dataUrl: (el.getAttribute('data-url') || '')?.slice(0, 150),
        src: (el.getAttribute('src') || '')?.slice(0, 150),
      })),
    };
  });

  // 10. iframeの検出
  const iframes = document.querySelectorAll('iframe');
  results.iframes = Array.from(iframes).map(f => ({
    src: f.src || '(no src)',
    id: f.id || null,
    name: f.name || null,
  }));

  results.summary = {
    dataTidCount: tidElements.length,
    uniqueTids: Object.keys(tidMap).length,
    messageCandidates: results.messageCandidate.length,
    senderCandidates: results.senderCandidate.length,
    url: window.location.href,
    iframeCount: iframes.length,
    frameType: window === window.top ? 'top-frame' : 'child-frame',
  };

  return results;
}

// ========== トークン調査 ==========

/**
 * Teams ページ内の MSAL / アクセストークンの格納場所を調査する。
 * sessionStorage, localStorage, cookie のキーパターンを検出する。
 * トークン値自体は先頭20文字のみ返す（セキュリティ配慮）。
 */
function inspectTokenStorage() {
  const results = {
    sessionStorage: [],
    localStorage: [],
    cookies: [],
    graphTokenFound: false,
    tokenSummary: {},
  };

  // 1. sessionStorage を調査
  try {
    for (let i = 0; i < sessionStorage.length; i++) {
      const key = sessionStorage.key(i);
      const value = sessionStorage.getItem(key);
      // MSAL 関連キーまたはアクセストークンを検出
      if (key.match(/msal|token|auth|access|bearer|credential/i) ||
          (value && value.length > 100 && value.startsWith('ey'))) {
        results.sessionStorage.push({
          key: key.slice(0, 100),
          valueLength: value?.length || 0,
          valuePreview: value?.slice(0, 20) + '...',
          looksLikeJwt: value?.startsWith('eyJ') || false,
        });
      }
    }
  } catch (e) {
    results.sessionStorage.push({ error: e.message });
  }

  // 2. localStorage を調査
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      const value = localStorage.getItem(key);
      if (key.match(/msal|token|auth|access|bearer|credential/i) ||
          (value && value.length > 100 && value.startsWith('ey'))) {
        results.localStorage.push({
          key: key.slice(0, 100),
          valueLength: value?.length || 0,
          valuePreview: value?.slice(0, 20) + '...',
          looksLikeJwt: value?.startsWith('eyJ') || false,
        });
      }
    }
  } catch (e) {
    results.localStorage.push({ error: e.message });
  }

  // 3. cookie を調査（アクセス可能な範囲）
  try {
    const cookies = document.cookie.split(';').map(c => c.trim());
    cookies.forEach(cookie => {
      const [name, ...rest] = cookie.split('=');
      const value = rest.join('=');
      if (name.match(/msal|token|auth|access|bearer/i) ||
          (value && value.length > 100 && value.startsWith('ey'))) {
        results.cookies.push({
          name: name.trim(),
          valueLength: value?.length || 0,
          valuePreview: value?.slice(0, 20) + '...',
          looksLikeJwt: value?.startsWith('eyJ') || false,
        });
      }
    });
  } catch (e) {
    results.cookies.push({ error: e.message });
  }

  // 4. tmp.auth.v1 形式のトークンを詳細調査（Teams の主要トークン格納形式）
  results.serviceTokens = [];
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      // tmp.auth.v1.*.Token.* パターンを検索
      if (!key.includes('.Token.')) continue;

      const value = localStorage.getItem(key);
      if (!value || value.length < 100) continue;

      try {
        const parsed = JSON.parse(value);
        const token = parsed?.item?.token;
        if (!token || token.length < 50) continue;

        // サービス名をキーから抽出
        const serviceMatch = key.match(/\.Token\.(.+)$/);
        const service = serviceMatch ? serviceMatch[1] : 'unknown';

        const tokenInfo = {
          key: key.slice(0, 120),
          service,
          tokenLength: token.length,
          looksLikeJwt: token.startsWith('eyJ'),
          tokenPreview: token.slice(0, 20) + '...',
          shouldRefresh: parsed?.shouldRefresh || false,
        };

        results.serviceTokens.push(tokenInfo);

        // Graph API トークンを特別に記録
        if (service.includes('GRAPH.MICROSOFT.COM')) {
          results.graphTokenFound = true;
          results.tokenSummary = {
            source: 'localStorage (tmp.auth.v1)',
            key: key.slice(0, 120),
            service,
            tokenLength: token.length,
            looksLikeJwt: token.startsWith('eyJ'),
            tokenPreview: token.slice(0, 20) + '...',
          };
        }
      } catch { /* not JSON */ }
    }
  } catch (e) {
    results.serviceTokens.push({ error: e.message });
  }

  // 5. MSAL キャッシュ構造を調査（credentialType 形式 — 実際のトークンはここ）
  results.msalTokens = [];
  try {
    const storages = [sessionStorage, localStorage];
    for (const storage of storages) {
      const storageName = storage === sessionStorage ? 'sessionStorage' : 'localStorage';
      for (let i = 0; i < storage.length; i++) {
        const key = storage.key(i);
        const value = storage.getItem(key);
        if (!value || value.length < 50) continue;
        try {
          const parsed = JSON.parse(value);
          if (parsed.credentialType === 'AccessToken' && parsed.secret) {
            const target = parsed.target || '';
            const isGraph = target.toLowerCase().includes('graph.microsoft.com');
            const isSharePoint = key.toLowerCase().includes('sharepoint') ||
                                 target.toLowerCase().includes('sharepoint');

            const tokenInfo = {
              source: storageName,
              key: key.slice(0, 120),
              target: target.slice(0, 200),
              expiresOn: parsed.expiresOn,
              realm: parsed.realm,
              tokenLength: parsed.secret.length,
              looksLikeJwt: parsed.secret.startsWith('eyJ'),
              isGraph,
              isSharePoint,
            };
            results.msalTokens.push(tokenInfo);

            // Graph API トークンを特別に記録
            if (isGraph && !results.graphTokenFound) {
              results.graphTokenFound = true;
              results.tokenSummary = {
                source: storageName,
                service: 'GRAPH.MICROSOFT.COM',
                key: key.slice(0, 120),
                target: target.slice(0, 200),
                expiresOn: parsed.expiresOn,
                tokenLength: parsed.secret.length,
                looksLikeJwt: parsed.secret.startsWith('eyJ'),
              };
            }
          }
        } catch { /* not JSON */ }
      }
    }
  } catch (e) {
    results.msalTokens.push({ error: e.message });
  }

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

    case 'INSPECT_TOKEN': {
      const result = inspectTokenStorage();
      log('log', 'トークン調査完了:', result.graphTokenFound ? 'Graph Token 検出' : '未検出');
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
