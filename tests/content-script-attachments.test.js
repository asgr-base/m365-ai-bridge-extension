/**
 * Content Script ファイル添付抽出テスト
 *
 * extractAttachments() がファイル名を正しく取得することを検証する。
 */

const { test, expect } = require('@playwright/test');
const { injectTeamsReader, sendCommand } = require('./helpers/chrome-mock');

test.describe('ファイル添付抽出', () => {
  test('textContent からファイル名を取得できる', async ({ page }) => {
    await page.setContent(`
      <html><body>
        <div data-tid="channel-pane-message">
          <span id="author-1">User A</span>
          <div data-tid="message-body">メッセージ本文</div>
          <div data-tid="file-preview-root">
            <span>report.xlsx</span>
          </div>
        </div>
      </body></html>
    `);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');
    const msg = response.data.messages[0];
    expect(msg.attachments).toHaveLength(1);
    expect(msg.attachments[0].name).toBe('report.xlsx');
  });

  test('button の aria-label からファイル名をフォールバック取得できる', async ({ page }) => {
    await page.setContent(`
      <html><body>
        <div data-tid="channel-pane-message">
          <span id="author-1">User A</span>
          <div data-tid="message-body">メッセージ本文</div>
          <div data-tid="file-preview-root">
            <button aria-label="ファイル slides.pptx の画像プレビュー"></button>
          </div>
        </div>
      </body></html>
    `);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');
    const msg = response.data.messages[0];
    expect(msg.attachments).toHaveLength(1);
    expect(msg.attachments[0].name).toBe('slides.pptx');
  });

  test('ファイル添付がないメッセージでは attachments が含まれない', async ({ page }) => {
    await page.setContent(`
      <html><body>
        <div data-tid="channel-pane-message">
          <span id="author-1">User A</span>
          <div data-tid="message-body">添付なしメッセージ</div>
        </div>
      </body></html>
    `);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');
    expect(response.data.messages[0].attachments).toBeUndefined();
  });

  test('複数ファイルが添付されたメッセージで全件取得できる', async ({ page }) => {
    await page.setContent(`
      <html><body>
        <div data-tid="channel-pane-message">
          <span id="author-1">User A</span>
          <div data-tid="message-body">複数添付メッセージ</div>
          <div data-tid="file-preview-root"><span>file1.pdf</span></div>
          <div data-tid="file-preview-root"><span>file2.docx</span></div>
          <div data-tid="file-preview-root"><span>file3.png</span></div>
        </div>
      </body></html>
    `);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');
    const msg = response.data.messages[0];
    expect(msg.attachments).toHaveLength(3);
    expect(msg.attachments.map(a => a.name)).toEqual(['file1.pdf', 'file2.docx', 'file3.png']);
  });

  test('file-preview-root 内に名前がない場合はスキップされる', async ({ page }) => {
    await page.setContent(`
      <html><body>
        <div data-tid="channel-pane-message">
          <span id="author-1">User A</span>
          <div data-tid="message-body">メッセージ</div>
          <div data-tid="file-preview-root"></div>
        </div>
      </body></html>
    `);
    await injectTeamsReader(page);

    const response = await sendCommand(page, 'READ_MESSAGES');
    // 空の file-preview-root からは名前が取れないので attachments なし
    expect(response.data.messages[0].attachments).toBeUndefined();
  });
});
