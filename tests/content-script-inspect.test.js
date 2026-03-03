/**
 * Content Script DOM調査 + トークン調査テスト
 *
 * inspectDom() と inspectTokenStorage() の動作を検証する。
 */

const { test, expect } = require('@playwright/test');
const path = require('path');
const { injectTeamsReader, sendCommand } = require('./helpers/chrome-mock');

const MOCK_HTML = path.join(__dirname, 'mock/teams-mock.html');

test.describe('DOM 構造調査 (INSPECT_DOM)', () => {
  test('data-tid 要素の数を正しくカウントする', async ({ page }) => {
    await page.goto(`file://${MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'INSPECT_DOM');
    expect(response.success).toBe(true);

    const data = response.data;
    expect(data.summary).toBeDefined();
    expect(data.summary.dataTidCount).toBeGreaterThan(0);
    expect(data.dataTidElements).toBeInstanceOf(Array);

    // channel-pane-message が3件あるはず
    const channelMsg = data.dataTidElements.find(e => e.tid === 'channel-pane-message');
    expect(channelMsg).toBeDefined();
    expect(channelMsg.count).toBe(3);
  });

  test('メッセージ候補を検出する', async ({ page }) => {
    await page.goto(`file://${MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'INSPECT_DOM');
    expect(response.data.messageCandidate).toBeInstanceOf(Array);
    expect(response.data.messageCandidate.length).toBeGreaterThan(0);
  });

  test('送信者候補を検出する', async ({ page }) => {
    await page.goto(`file://${MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'INSPECT_DOM');
    expect(response.data.senderCandidate).toBeInstanceOf(Array);
  });

  test('ファイル添付候補を検出する', async ({ page }) => {
    await page.goto(`file://${MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'INSPECT_DOM');
    expect(response.data.fileCandidate).toBeInstanceOf(Array);
    // teams-mock.html には file-preview-root が1件ある
    const filePreviews = response.data.fileCandidate.find(
      f => f.selector.includes('file-preview-root')
    );
    expect(filePreviews).toBeDefined();
  });
});

test.describe('トークンストレージ調査 (INSPECT_TOKEN)', () => {
  test('トークン調査結果の基本構造が正しい', async ({ page }) => {
    await page.goto(`file://${MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'INSPECT_TOKEN');
    expect(response.success).toBe(true);

    const data = response.data;
    expect(data.sessionStorage).toBeInstanceOf(Array);
    expect(data.localStorage).toBeInstanceOf(Array);
    expect(data.cookies).toBeInstanceOf(Array);
    expect(typeof data.graphTokenFound).toBe('boolean');
  });

  test('モック MSAL トークンを検出できる', async ({ page }) => {
    await page.goto(`file://${MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'INSPECT_TOKEN');
    const data = response.data;

    // teams-mock.html に localStorage の MSAL AccessToken モックがある
    expect(data.msalTokens).toBeInstanceOf(Array);
    expect(data.msalTokens.length).toBeGreaterThan(0);

    // Graph トークンのモックが検出されるはず
    const graphToken = data.msalTokens.find(t => t.isGraph);
    expect(graphToken).toBeDefined();
    expect(graphToken.looksLikeJwt).toBe(true);
  });

  test('空のストレージでもエラーなく動作する', async ({ page }) => {
    // about:blank では sessionStorage がアクセス不可な場合があるが、クラッシュしないことが重要
    await page.goto('about:blank');
    await page.setContent('<html><body></body></html>');
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'INSPECT_TOKEN');
    expect(response.success).toBe(true);
    expect(response.data.graphTokenFound).toBe(false);
    // エラーエントリが含まれる場合もあるが、配列であること
    expect(response.data.msalTokens).toBeInstanceOf(Array);
  });
});
