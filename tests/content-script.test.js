/**
 * Content Script DOM読み取りテスト（チャンネルメッセージ）
 *
 * teams-reader.js の実コードを Chrome API モック付きで注入し、
 * 本番同等のコードパスをテストする。
 */

const { test, expect } = require('@playwright/test');
const path = require('path');
const { injectTeamsReader, sendCommand } = require('./helpers/chrome-mock');

const MOCK_HTML = path.join(__dirname, 'mock/teams-mock.html');

test.describe('Content Script DOM 読み取り', () => {
  test('モックTeamsページからメッセージ3件を取得できる', async ({ page }) => {
    await page.goto(`file://${MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');

    expect(response.success).toBe(true);
    expect(response.data.messages).toHaveLength(3);
    expect(response.data.messages[0].sender).toBe('田中 太郎');
    expect(response.data.messages[0].body).toContain('おはようございます');
    expect(response.data.messages[0].timestamp).toBe('2026-02-26T09:00:00Z');
    expect(response.data.messages[1].sender).toBe('山田 花子');
    expect(response.data.messages[2].sender).toBe('中村 太一');
  });

  test('チャンネル名が正しく取得できる', async ({ page }) => {
    await page.goto(`file://${MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');
    expect(response.data.context.channelName).toBe('General');
  });

  test('返信フォームへのテキスト挿入が動作する', async ({ page }) => {
    await page.goto(`file://${MOCK_HTML}`);
    const replyBox = page.locator('[data-tid="ckeditor"]');
    await replyBox.click();
    await replyBox.fill('テスト返信です。');
    const content = await replyBox.textContent();
    expect(content).toBe('テスト返信です。');
  });

  test('メッセージが0件のページでは空配列を返す', async ({ page }) => {
    await page.setContent('<html><body><div data-tid="channelTitle-text">Empty</div></body></html>');
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');
    expect(response.success).toBe(true);
    expect(response.data.messages).toHaveLength(0);
  });

  test('メッセージにファイル添付がある場合 attachments が含まれる', async ({ page }) => {
    await page.goto(`file://${MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');
    // msg-1770002 にファイル添付がある
    const msgWithAttachment = response.data.messages[1];
    expect(msgWithAttachment.attachments).toBeDefined();
    expect(msgWithAttachment.attachments).toHaveLength(1);
    expect(msgWithAttachment.attachments[0].name).toBe('quarterly-report.xlsx');

    // msg-1770001 にはファイル添付がない
    expect(response.data.messages[0].attachments).toBeUndefined();
  });

  test('メンション（TO/CC）が正しく抽出される', async ({ page }) => {
    await page.goto(`file://${MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');

    // msg-1770001: TO: 山田 花子, CC: 中村 太一
    const msg0 = response.data.messages[0];
    expect(msg0.mentions).toBeDefined();
    expect(msg0.mentions.to).toEqual(['山田 花子']);
    expect(msg0.mentions.cc).toEqual(['中村 太一']);

    // msg-1770002: TO: 田中 太郎, CC なし
    const msg1 = response.data.messages[1];
    expect(msg1.mentions).toBeDefined();
    expect(msg1.mentions.to).toEqual(['田中 太郎']);
    expect(msg1.mentions.cc).toEqual([]);

    // msg-1770003: メンションなし
    expect(response.data.messages[2].mentions).toBeUndefined();
  });

  test('返信数と返信者が正しく取得される', async ({ page }) => {
    await page.goto(`file://${MOCK_HTML}`);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');

    // msg-1770001: 返信2件（山田 花子, 中村 太一）
    const msg0 = response.data.messages[0];
    expect(msg0.replyCount).toBe(2);
    expect(msg0.replySenders).toContain('山田 花子');
    expect(msg0.replySenders).toContain('中村 太一');

    // msg-1770002: 返信なし
    expect(response.data.messages[1].replyCount).toBe(0);

    // msg-1770003: 返信なし
    expect(response.data.messages[2].replyCount).toBe(0);
  });

  test('NBSPテキストノードで隣接メンションが正しくグループ化される（Teams実DOM構造）', async ({ page }) => {
    // Teams の実 DOM: 各メンション単語は個別の <div> ラッパーに入っている。
    // 同一人物の単語間は \u00a0 (NBSP) テキストノードで結合。
    // 異なる人物の単語間はテキストノードなし（空文字）→ 別人として分離 (Rule 5)
    // <span>Suzuki</span>&nbsp;<span>Ichiro</span><span>YAMAMOTO</span>&nbsp;<span>KENJI</span>
    await page.setContent(`
      <html><body>
        <div data-tid="channelTitle-text">TestChannel</div>
        <div data-tid="channel-pane-message" id="msg-ts-1">
          <div data-tid="post-message-subheader">
            <div data-tid="post-message-header-avatar"></div>
            <span id="author-ts-1" class="fui-StyledText">送信者D</span>
          </div>
          <div data-tid="timestamp" datetime="2026-03-01T13:00:00Z">13:00</div>
          <div id="message-body-ts-1" data-tid="message-body"><div><span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">Suzuki</span>&nbsp;<span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">Ichiro</span><span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">YAMAMOTO</span>&nbsp;<span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">KENJI</span>
テストメッセージです。</div></div>
        </div>
      </body></html>
    `);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');
    const msg = response.data.messages[0];
    expect(msg.mentions).toBeDefined();
    expect(msg.mentions.to).toEqual(['Suzuki Ichiro', 'YAMAMOTO KENJI']);
    expect(msg.mentions.cc).toEqual([]);
  });

  test('NBSPテキストノード方式で3人のメンションが正しく分離される', async ({ page }) => {
    // [SIS]&nbsp;[高橋]&nbsp;[京子][Sato]&nbsp;[Jiro][Suzuki]&nbsp;[Ichiro]
    //                        ↑ empty=boundary      ↑ empty=boundary
    await page.setContent(`
      <html><body>
        <div data-tid="channelTitle-text">TestChannel</div>
        <div data-tid="channel-pane-message" id="msg-ts-2">
          <div data-tid="post-message-subheader">
            <div data-tid="post-message-header-avatar"></div>
            <span id="author-ts-2" class="fui-StyledText">送信者E</span>
          </div>
          <div data-tid="timestamp" datetime="2026-03-01T14:00:00Z">14:00</div>
          <div id="message-body-ts-2" data-tid="message-body"><div><span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">SIS</span>&nbsp;<span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">高橋</span>&nbsp;<span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">京子</span><span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">Sato</span>&nbsp;<span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">Jiro</span><span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">Suzuki</span>&nbsp;<span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">Ichiro</span>
メッセージです。</div></div>
        </div>
      </body></html>
    `);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');
    const msg = response.data.messages[0];
    expect(msg.mentions).toBeDefined();
    expect(msg.mentions.to).toEqual(['SIS 高橋 京子', 'Sato Jiro', 'Suzuki Ichiro']);
    expect(msg.mentions.cc).toEqual([]);
  });

  test('テキストノードなしの隣接メンションが別人として分離される（ラッパー経由）', async ({ page }) => {
    // ラッパー要素内は空白テキストノードで同一人物、ラッパー間はテキストなし → 別人 (Rule 5)
    await page.setContent(`
      <html><body>
        <div data-tid="channelTitle-text">TestChannel</div>
        <div data-tid="channel-pane-message" id="msg-wrap-1">
          <div data-tid="post-message-subheader">
            <div data-tid="post-message-header-avatar"></div>
            <span id="author-wrap-1" class="fui-StyledText">送信者B</span>
          </div>
          <div data-tid="timestamp" datetime="2026-03-01T11:00:00Z">11:00</div>
          <div id="message-body-wrap-1" data-tid="message-body"><a class="mention-entity"><span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">高橋</span> <span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">京子</span></a><a class="mention-entity"><span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">伊藤</span> <span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">健</span></a>

確認お願いします。</div>
        </div>
      </body></html>
    `);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');
    const msg = response.data.messages[0];
    expect(msg.mentions).toBeDefined();
    expect(msg.mentions.to).toEqual(['高橋 京子', '伊藤 健']);
  });

  test('全角スペースで区切られたメンションが別人として分離される', async ({ page }) => {
    await page.setContent(`
      <html><body>
        <div data-tid="channelTitle-text">TestChannel</div>
        <div data-tid="channel-pane-message" id="msg-fw-1">
          <div data-tid="post-message-subheader">
            <div data-tid="post-message-header-avatar"></div>
            <span id="author-fw-1" class="fui-StyledText">送信者C</span>
          </div>
          <div data-tid="timestamp" datetime="2026-03-01T12:00:00Z">12:00</div>
          <div id="message-body-fw-1" data-tid="message-body"><span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">佐藤</span> <span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">太郎</span>\u3000<span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">鈴木</span> <span dir="auto" itemtype="http://schema.skype.com/Mention" class="fui-StyledText">花子</span>

メッセージです。</div>
        </div>
      </body></html>
    `);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');
    const msg = response.data.messages[0];
    expect(msg.mentions).toBeDefined();
    expect(msg.mentions.to).toEqual(['佐藤 太郎', '鈴木 花子']);
  });
});
