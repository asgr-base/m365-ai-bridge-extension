/**
 * Content Script DOM読み取りテスト
 *
 * Chrome Extension の content_scripts は teams.microsoft.com にしか注入されないため、
 * Playwright でモック HTML ページを開き、スクリプトを手動注入してテストする。
 */

const { test, expect } = require('@playwright/test');
const path = require('path');
const fs = require('fs');

const MOCK_HTML = path.join(__dirname, 'mock/teams-mock.html');
const CONTENT_SCRIPT = path.join(__dirname, '../content/teams-reader.js');

// Content Script のコアロジックのみを抽出してテスト（chrome.runtime依存部分を除く）
// セレクタは content/teams-reader.js と同期させること
const extractorCode = `
  const SELECTORS = {
    messageContainer: '[data-tid="channel-pane-message"]',
    messageBody: '[data-tid="message-body"]',
    senderName: 'span[id^="author-"]',
    senderHeader: '[data-tid="post-message-subheader"], [data-tid="reply-message-header"]',
    timestamp: '[data-tid="timestamp"]',
    channelName: '[data-tid="channelTitle-text"]',
    chatTitle: '[data-tid="chat-title"]',
    replyBox: '[data-tid="ckeditor"], [role="textbox"][contenteditable="true"]',
  };

  function extractMessages() {
    const messages = [];
    const containers = document.querySelectorAll(SELECTORS.messageContainer);

    containers.forEach((container, index) => {
      const bodyEl = container.querySelector(SELECTORS.messageBody);
      const senderEl = container.querySelector(SELECTORS.senderName);
      const timeEl = container.querySelector(SELECTORS.timestamp);

      messages.push({
        index,
        sender: senderEl?.textContent?.trim() || 'Unknown',
        body: bodyEl?.innerText?.trim() || '',
        timestamp: timeEl?.getAttribute('datetime') || '',
      });
    });

    return {
      context: {
        channelName: document.querySelector(SELECTORS.channelName)?.textContent?.trim() || null,
        chatTitle: document.querySelector(SELECTORS.chatTitle)?.textContent?.trim() || null,
        pageTitle: document.title,
        url: window.location.href,
      },
      messages,
    };
  }

  window.__extractMessages = extractMessages;
`;

test.describe('Content Script DOM 読み取り', () => {
  test('モックTeamsページからメッセージ3件を取得できる', async ({ page }) => {
    await page.goto(`file://${MOCK_HTML}`);
    await page.addScriptTag({ content: extractorCode });

    const result = await page.evaluate(() => window.__extractMessages());

    expect(result.messages).toHaveLength(3);
    expect(result.messages[0].sender).toBe('田中 太郎');
    expect(result.messages[0].body).toContain('おはようございます');
    expect(result.messages[0].timestamp).toBe('2026-02-26T09:00:00Z');
    expect(result.messages[1].sender).toBe('山田 花子');
    expect(result.messages[2].sender).toBe('都甲 篤史');
  });

  test('チャンネル名が正しく取得できる', async ({ page }) => {
    await page.goto(`file://${MOCK_HTML}`);
    await page.addScriptTag({ content: extractorCode });

    const result = await page.evaluate(() => window.__extractMessages());
    expect(result.context.channelName).toBe('General');
  });

  test('返信フォームへのテキスト挿入が動作する', async ({ page }) => {
    await page.goto(`file://${MOCK_HTML}`);

    const replyText = 'テスト返信です。';
    const replyBox = page.locator('[data-tid="ckeditor"]');

    await replyBox.click();
    await replyBox.fill(replyText);

    const content = await replyBox.textContent();
    expect(content).toBe(replyText);
  });

  test('メッセージが0件のページでは空配列を返す', async ({ page }) => {
    // 空のHTMLページ
    await page.setContent('<html><body><div data-tid="channelTitle-text">Empty</div></body></html>');
    await page.addScriptTag({ content: extractorCode });

    const result = await page.evaluate(() => window.__extractMessages());
    expect(result.messages).toHaveLength(0);
    expect(result.context.channelName).toBe('Empty');
  });
});
