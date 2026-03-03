/**
 * Content Script Graph トークン取得テスト
 *
 * getGraphToken() と GET_GRAPH_TOKEN コマンドの動作を検証する。
 */

const { test, expect } = require('@playwright/test');
const path = require('path');
const { injectTeamsReader, sendCommand } = require('./helpers/chrome-mock');

const BLANK_HTML = path.join(__dirname, 'mock/blank.html');

test.describe('Graph トークン取得 (GET_GRAPH_TOKEN)', () => {
  test('localStorage の MSAL AccessToken から Graph JWT を取得できる', async ({ page }) => {
    await page.goto(`file://${BLANK_HTML}`);

    // MSAL AccessToken モックを localStorage に設定
    await page.evaluate(() => {
      localStorage.setItem('msal-access-token-graph', JSON.stringify({
        credentialType: 'AccessToken',
        secret: 'eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.mock-graph-token-payload.signature',
        target: 'https://graph.microsoft.com/.default openid profile',
        expiresOn: '9999999999',
      }));
    });

    await injectTeamsReader(page);

    const response = await sendCommand(page, 'GET_GRAPH_TOKEN');
    expect(response.success).toBe(true);
    expect(response.data.token).toContain('eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9');
    expect(response.data.expiresOn).toBe('9999999999');
    expect(response.data.target).toContain('graph.microsoft.com');
  });

  test('Graph トークンがない場合はエラーを返す', async ({ page }) => {
    await page.goto(`file://${BLANK_HTML}`);

    // Graph 以外のトークンのみ設定
    await page.evaluate(() => {
      localStorage.setItem('msal-access-token-spo', JSON.stringify({
        credentialType: 'AccessToken',
        secret: 'eyJhbGciOiJSUzI1NiJ9.sharepoint-token.sig',
        target: 'https://contoso.sharepoint.com/.default',
        expiresOn: '9999999999',
      }));
    });

    await injectTeamsReader(page);

    const response = await sendCommand(page, 'GET_GRAPH_TOKEN');
    expect(response.success).toBe(false);
    expect(response.error).toContain('Graph token not found');
  });
});
