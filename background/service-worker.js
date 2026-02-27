/**
 * M365 AI Bridge - Background Service Worker
 *
 * Claude Code（外部）からのリクエストを受け付け、
 * content script 経由で Teams を操作する中継役。
 *
 * 通信フロー:
 *   Claude Code
 *     → fetch POST http://localhost:3765/...  (bridge server)
 *     → chrome.runtime.sendMessage            (background)
 *     → chrome.tabs.sendMessage               (content script)
 *     → Teams DOM 操作
 */

'use strict';

// ========== 設定 ==========
const CONFIG = {
  bridgePort: 3765,
  version: '0.1.0',
};

// ========== ユーティリティ ==========

function log(level, ...args) {
  const prefix = '[M365 AI Bridge SW]';
  if (level === 'error') console.error(prefix, ...args);
  else console.log(prefix, ...args);
}

// ========== Teams タブ検索 ==========

/**
 * 現在開いている Teams タブを取得する
 * @returns {Promise<chrome.tabs.Tab|null>}
 */
async function getTeamsTab() {
  const tabs = await chrome.tabs.query({
    url: [
      'https://teams.microsoft.com/*',
      'https://*.teams.microsoft.com/*',
      'https://teams.cloud.microsoft/*',
      'https://*.teams.cloud.microsoft/*',
    ],
  });
  return tabs.length > 0 ? tabs[0] : null;
}

// ========== コンテンツスクリプト通信 ==========

/**
 * Teams タブのコンテンツスクリプトにコマンドを送信する
 * @param {string} command - コマンド名
 * @param {Object} payload - 追加データ
 * @returns {Promise<Object>} レスポンス
 */
async function sendToTeams(command, payload = {}) {
  const tab = await getTeamsTab();
  if (!tab) {
    return { success: false, error: 'Teams tab not found. Please open teams.microsoft.com' };
  }

  try {
    const response = await chrome.tabs.sendMessage(tab.id, {
      command,
      ...payload,
    });
    return response;
  } catch (err) {
    log('error', 'Content script への送信失敗:', err.message);
    return { success: false, error: err.message };
  }
}

// ========== API ハンドラ ==========

/**
 * popup または外部からのメッセージを処理する
 */
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  log('log', 'メッセージ受信:', request.action);

  handleRequest(request)
    .then(sendResponse)
    .catch((err) => sendResponse({ success: false, error: err.message }));

  return true; // 非同期レスポンスを許可
});

async function handleRequest(request) {
  switch (request.action) {
    case 'READ_MESSAGES':
      return await sendToTeams('READ_MESSAGES');

    case 'INSERT_REPLY':
      return await sendToTeams('INSERT_REPLY', { text: request.text });

    case 'PING_TEAMS':
      return await sendToTeams('PING');

    case 'INSPECT_DOM':
      return await sendToTeams('INSPECT_DOM');

    case 'GET_STATUS': {
      const tab = await getTeamsTab();
      return {
        success: true,
        status: {
          version: CONFIG.version,
          teamsTabFound: !!tab,
          teamsUrl: tab?.url || null,
        },
      };
    }

    default:
      return { success: false, error: `Unknown action: ${request.action}` };
  }
}

// ========== 初期化 ==========

chrome.runtime.onInstalled.addListener(() => {
  log('log', `M365 AI Bridge v${CONFIG.version} インストール完了`);
});

log('log', `M365 AI Bridge Service Worker 起動 v${CONFIG.version}`);
