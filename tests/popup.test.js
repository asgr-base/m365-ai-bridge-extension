/**
 * Popup UI テスト
 *
 * popup.html を直接ロードし、chrome.runtime をモックして
 * ボタンクリック → 出力表示のフローを検証する。
 */

const { test, expect } = require('@playwright/test');
const path = require('path');
const fs = require('fs');

const POPUP_HTML = path.join(__dirname, '../popup/popup.html');
const POPUP_JS = path.join(__dirname, '../popup/popup.js');

/**
 * popup.js をモック付きで注入するために page.route で popup.js を置換する。
 * chrome.runtime.sendMessage と fetch をモックし、各アクションに適切なレスポンスを返す。
 */
function getPopupChromeMock() {
  return `
    window.chrome = {
      runtime: {
        sendMessage(msg) {
          if (msg.action === 'GET_STATUS') {
            return Promise.resolve({
              success: true,
              status: { teamsTabFound: true, version: '0.1.0', teamsUrl: 'https://teams.microsoft.com/test' },
            });
          }
          if (msg.action === 'READ_MESSAGES') {
            return Promise.resolve({
              success: true,
              data: {
                context: { channelName: 'General', chatTitle: null, pageTitle: 'Teams' },
                messages: [
                  { index: 0, sender: 'Test User', body: 'Hello world', timestamp: '2026-02-26T09:00:00Z' },
                  { index: 1, sender: 'Another User', body: 'Test message', timestamp: '2026-02-26T09:05:00Z' },
                ],
                method: 'primary',
              },
            });
          }
          if (msg.action === 'INSPECT_DOM') {
            return Promise.resolve({
              success: true,
              data: {
                summary: { dataTidCount: 10, uniqueTids: 5, messageCandidates: 2, senderCandidates: 1, url: 'https://teams.microsoft.com/test', iframeCount: 0, frameType: 'top-frame' },
                dataTidElements: [{ tid: 'channel-pane-message', count: 3 }],
                messageCandidate: [{ selector: '[data-tid="channel-pane-message"]', count: 3, sampleClasses: [], sampleDataAttrs: [] }],
                senderCandidate: [],
                timestampCandidate: [],
                replyBoxCandidate: [],
                sampleHtml: [],
                fileCandidate: [],
                dmIdElements: [],
                iframes: [],
              },
            });
          }
          if (msg.action === 'INSPECT_TOKEN') {
            return Promise.resolve({
              success: true,
              data: {
                sessionStorage: [],
                localStorage: [],
                cookies: [],
                graphTokenFound: true,
                tokenSummary: {
                  source: 'localStorage',
                  service: 'GRAPH.MICROSOFT.COM',
                  target: 'https://graph.microsoft.com/.default',
                  tokenLength: 3000,
                  looksLikeJwt: true,
                  expiresOn: '9999999999',
                },
                msalTokens: [{
                  source: 'localStorage',
                  key: 'test-token-key',
                  target: 'https://graph.microsoft.com/.default',
                  expiresOn: '9999999999',
                  tokenLength: 3000,
                  looksLikeJwt: true,
                  isGraph: true,
                  isSharePoint: false,
                }],
                serviceTokens: [],
              },
            });
          }
          return Promise.resolve({ success: false, error: 'Unknown action' });
        },
      },
    };

    // fetch をモック（bridge-server ヘルスチェック用）
    const _origFetch = window.fetch;
    window.fetch = function(url, ...args) {
      if (typeof url === 'string' && url.includes('localhost:3765/health')) {
        return Promise.resolve(new Response(JSON.stringify({ status: 'ok' }), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        }));
      }
      if (typeof url === 'string' && url.includes('localhost:3765')) {
        return Promise.resolve(new Response('{}', {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        }));
      }
      return _origFetch.call(this, url, ...args);
    };
  `;
}

test.describe('Popup UI', () => {
  test.beforeEach(async ({ page }) => {
    const popupJsCode = fs.readFileSync(POPUP_JS, 'utf-8');
    const mockCode = getPopupChromeMock();

    // popup.js のリクエストをインターセプトしてモック付きに差し替え
    await page.route('**/popup.js', async (route) => {
      await route.fulfill({
        contentType: 'application/javascript',
        body: mockCode + '\n' + popupJsCode,
      });
    });
  });

  test('ポップアップのタイトルとヘッダーが正しく表示される', async ({ page }) => {
    await page.goto(`file://${POPUP_HTML}`);

    const header = await page.textContent('header h1');
    expect(header).toBe('M365 AI Bridge');

    const badge = await page.textContent('header .badge');
    expect(badge).toBe('PoC');
  });

  test('初期化後にステータスドットが ok になる', async ({ page }) => {
    await page.goto(`file://${POPUP_HTML}`);

    // checkStatus() は非同期なので少し待つ
    await page.waitForFunction(() => {
      const teamsDot = document.getElementById('teams-dot');
      return teamsDot?.classList.contains('ok');
    }, { timeout: 3000 });

    const teamsDotClass = await page.getAttribute('#teams-dot', 'class');
    expect(teamsDotClass).toContain('ok');

    const bridgeDotClass = await page.getAttribute('#bridge-dot', 'class');
    expect(bridgeDotClass).toContain('ok');
  });

  test('「メッセージを読み取る」ボタンでメッセージが表示される', async ({ page }) => {
    await page.goto(`file://${POPUP_HTML}`);

    await page.click('#read-btn');

    // 出力が表示されるまで待つ
    await page.waitForFunction(() => {
      const output = document.getElementById('output');
      return output?.textContent?.includes('General');
    }, { timeout: 3000 });

    const outputText = await page.textContent('#output');
    expect(outputText).toContain('General');
    expect(outputText).toContain('Test User');
    expect(outputText).toContain('メッセージ数: 2件');
  });

  test('「DOM構造を調査」ボタンで調査結果が表示される', async ({ page }) => {
    await page.goto(`file://${POPUP_HTML}`);

    await page.click('#inspect-btn');

    await page.waitForFunction(() => {
      const output = document.getElementById('output');
      return output?.textContent?.includes('DOM 構造調査結果');
    }, { timeout: 3000 });

    const outputText = await page.textContent('#output');
    expect(outputText).toContain('DOM 構造調査結果');
    expect(outputText).toContain('data-tid 要素数');
  });

  test('「トークン調査」ボタンで調査結果が表示される', async ({ page }) => {
    await page.goto(`file://${POPUP_HTML}`);

    await page.click('#token-btn');

    await page.waitForFunction(() => {
      const output = document.getElementById('output');
      return output?.textContent?.includes('トークンストレージ調査結果');
    }, { timeout: 3000 });

    const outputText = await page.textContent('#output');
    expect(outputText).toContain('トークンストレージ調査結果');
    expect(outputText).toContain('Graph Token 検出: YES');
  });

  test('コピーボタンのテキストが「OK」に変わる', async ({ page, context }) => {
    // クリップボード権限を付与
    await context.grantPermissions(['clipboard-write', 'clipboard-read']);

    await page.goto(`file://${POPUP_HTML}`);

    // まず出力を表示させる
    await page.click('#read-btn');
    await page.waitForFunction(() => {
      const wrapper = document.getElementById('output-wrapper');
      return wrapper?.style.display === 'block';
    }, { timeout: 3000 });

    // コピーボタンをクリック
    await page.click('#copy-btn');

    // ボタンテキストが「OK」に変わるまで待つ（非同期処理のため）
    await page.waitForFunction(() => {
      return document.getElementById('copy-btn')?.textContent === 'OK';
    }, { timeout: 3000 });

    const btnText = await page.textContent('#copy-btn');
    expect(btnText).toBe('OK');
  });
});
