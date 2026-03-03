/**
 * Content Script DM/グループチャット抽出テスト
 *
 * DM 専用の chat-pane-item 構造からメッセージを正しく抽出できることを検証する。
 */

const { test, expect } = require('@playwright/test');
const path = require('path');
const { injectTeamsReader, sendCommand } = require('./helpers/chrome-mock');

const DM_MOCK_HTML = path.join(__dirname, 'mock/teams-dm-mock.html');
const CHANNEL_MOCK_HTML = path.join(__dirname, 'mock/teams-mock.html');

test.describe('DM/グループチャット メッセージ抽出', () => {
  test('DMメッセージ3件を正しく取得できる', async ({ page }) => {
    await page.goto(`file://${DM_MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');

    expect(response.success).toBe(true);
    expect(response.data.method).toBe('dm-chat-pane');
    expect(response.data.messages).toHaveLength(3);
    expect(response.data.messages[0].sender).toBe('佐藤 一郎');
    expect(response.data.messages[1].sender).toBe('中村 太一');
    expect(response.data.messages[2].sender).toBe('佐藤 一郎');
  });

  test('DMメッセージのタイムスタンプが正しく取得される', async ({ page }) => {
    await page.goto(`file://${DM_MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');

    expect(response.data.messages[0].timestamp).toBe('2026-02-27T14:00:00Z');
    expect(response.data.messages[1].timestamp).toBe('2026-02-27T14:05:00Z');
    expect(response.data.messages[2].timestamp).toBe('2026-02-27T14:10:00Z');
  });

  test('DMメッセージの messageId が message-body-{id} から抽出される', async ({ page }) => {
    await page.goto(`file://${DM_MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');

    expect(response.data.messages[0].messageId).toBe('1770100');
    expect(response.data.messages[1].messageId).toBe('1770101');
  });

  test('DMメッセージにファイル添付がある場合 attachments が含まれる', async ({ page }) => {
    await page.goto(`file://${DM_MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');

    // dm-item-1 には shared-doc.pdf が添付
    const msg0 = response.data.messages[0];
    expect(msg0.attachments).toBeDefined();
    expect(msg0.attachments[0].name).toBe('shared-doc.pdf');

    // dm-item-2 にはファイル添付なし
    expect(response.data.messages[1].attachments).toBeUndefined();
  });

  test('aria-label からファイル名が正しく抽出される（フォールバック）', async ({ page }) => {
    await page.goto(`file://${DM_MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');

    // dm-item-3 は button aria-label のみ
    const msg2 = response.data.messages[2];
    expect(msg2.attachments).toBeDefined();
    expect(msg2.attachments[0].name).toBe('presentation.pptx');
  });

  test('チャットタイトルが正しく取得される', async ({ page }) => {
    await page.goto(`file://${DM_MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');
    expect(response.data.context.chatTitle).toBe('佐藤 一郎');
  });

  test('DMメッセージでメンションが正しく抽出される', async ({ page }) => {
    await page.goto(`file://${DM_MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');

    // dm-item-1: メンションあり（TO: 中村 太一）
    const msg0 = response.data.messages[0];
    expect(msg0.mentions).toBeDefined();
    expect(msg0.mentions.to).toEqual(['中村 太一']);
    expect(msg0.mentions.cc).toEqual([]);

    // dm-item-2: メンションなし
    expect(response.data.messages[1].mentions).toBeUndefined();
  });

  test('DMメッセージには replyCount が含まれない', async ({ page }) => {
    await page.goto(`file://${DM_MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');
    // DM にはスレッド構造がないため replyCount は undefined
    expect(response.data.messages[0].replyCount).toBeUndefined();
  });

  test('chat-pane-item がある場合、チャンネルメッセージより DM 抽出が優先される', async ({ page }) => {
    // DM + チャンネルの両方がある HTML を構築
    await page.setContent(`
      <html><body>
        <div data-tid="channelTitle-text">General</div>
        <div data-tid="channel-pane-message" id="msg-ch1">
          <span id="author-ch1">Channel User</span>
          <div data-tid="message-body">channel message</div>
        </div>
        <div data-tid="chat-pane-item" id="dm-1">
          <div data-tid="message-author-name">DM User</div>
          <div data-tid="chat-pane-message" id="message-body-1001">dm message</div>
        </div>
      </body></html>
    `);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');
    expect(response.data.method).toBe('dm-chat-pane');
    expect(response.data.messages[0].sender).toBe('DM User');
  });
});
