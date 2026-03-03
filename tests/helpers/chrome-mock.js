/**
 * Chrome API モック + Content Script 注入ヘルパー
 *
 * teams-reader.js の実コードを Playwright ページに注入するために、
 * chrome.runtime 等の API をモックする。
 */

const fs = require('fs');
const path = require('path');

const CONTENT_SCRIPT_PATH = path.join(__dirname, '../../content/teams-reader.js');

/**
 * Chrome API モックコード（文字列）を返す。
 * ページに先行注入し、teams-reader.js が chrome.runtime を参照しても動作するようにする。
 */
function getChromeMockCode() {
  return `
    // bridge-server への fetch を抑制（autoPush 等）
    const _origFetch = window.fetch;
    window.fetch = function(url, ...args) {
      if (typeof url === 'string' && url.includes('localhost:3765')) {
        return Promise.resolve(new Response(JSON.stringify({ success: true }), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        }));
      }
      return _origFetch.call(this, url, ...args);
    };

    // Chrome Extension API モック
    window.chrome = {
      runtime: {
        onMessage: {
          addListener(fn) {
            window.__chromeMessageListener = fn;
          },
        },
        sendMessage(msg) {
          return Promise.resolve({ success: true });
        },
      },
      tabs: {
        query() { return Promise.resolve([]); },
        sendMessage() { return Promise.resolve({}); },
      },
      storage: {
        local: {
          get(keys, cb) { if (cb) cb({}); return Promise.resolve({}); },
          set(items, cb) { if (cb) cb(); return Promise.resolve(); },
        },
      },
      scripting: {
        executeScript() { return Promise.resolve([]); },
      },
    };
  `;
}

/**
 * teams-reader.js のソースコードを読み込んで返す。
 */
function getTeamsReaderCode() {
  return fs.readFileSync(CONTENT_SCRIPT_PATH, 'utf-8');
}

/**
 * Playwright ページに Chrome モック + teams-reader.js を注入する。
 * 注入後は window.__chromeMessageListener でコマンドを送信可能。
 */
async function injectTeamsReader(page) {
  const code = getChromeMockCode() + '\n' + getTeamsReaderCode();
  await page.addScriptTag({ content: code });
}

/**
 * chrome.runtime.onMessage ハンドラにコマンドを送信し、レスポンスを返す。
 *
 * @param {import('@playwright/test').Page} page
 * @param {string} command - コマンド名（READ_MESSAGES, INSPECT_DOM 等）
 * @param {Object} payload - 追加パラメータ
 * @returns {Promise<Object>} sendResponse に渡されたオブジェクト
 */
async function sendCommand(page, command, payload = {}) {
  return await page.evaluate(({ command, payload }) => {
    return new Promise((resolve, reject) => {
      const listener = window.__chromeMessageListener;
      if (!listener) {
        reject(new Error('No chrome message listener registered'));
        return;
      }
      listener(
        { command, ...payload },
        {},           // sender (空オブジェクト)
        (response) => resolve(response),
      );
    });
  }, { command, payload });
}

module.exports = { getChromeMockCode, getTeamsReaderCode, injectTeamsReader, sendCommand };
