'use strict';

const BRIDGE_URL = 'http://localhost:3765';

// ========== DOM 要素 ==========
const teamsDot = document.getElementById('teams-dot');
const teamsStatus = document.getElementById('teams-status');
const bridgeDot = document.getElementById('bridge-dot');
const bridgeStatus = document.getElementById('bridge-status');
const readBtn = document.getElementById('read-btn');
const output = document.getElementById('output');

// ========== 状態表示 ==========

function setStatus(dotEl, labelEl, state, text) {
  dotEl.className = `status-dot ${state}`;
  labelEl.textContent = text;
}

function showOutput(text, type = 'normal') {
  output.style.display = 'block';
  output.className = `output ${type}`;
  output.textContent = text;
}

// ========== 初期化: 接続状態チェック ==========

async function checkStatus() {
  // Teams タブの確認
  try {
    const res = await chrome.runtime.sendMessage({ action: 'GET_STATUS' });
    if (res?.status?.teamsTabFound) {
      setStatus(teamsDot, teamsStatus, 'ok', '接続済み');
    } else {
      setStatus(teamsDot, teamsStatus, 'error', '未検出');
    }
  } catch (err) {
    setStatus(teamsDot, teamsStatus, 'error', 'エラー');
  }

  // ブリッジサーバーの確認
  try {
    const res = await fetch(`${BRIDGE_URL}/health`, { signal: AbortSignal.timeout(2000) });
    if (res.ok) {
      setStatus(bridgeDot, bridgeStatus, 'ok', '起動中');
    } else {
      setStatus(bridgeDot, bridgeStatus, 'error', `HTTP ${res.status}`);
    }
  } catch {
    setStatus(bridgeDot, bridgeStatus, 'error', '未起動');
  }
}

// ========== メッセージ読み取り ==========

readBtn.addEventListener('click', async () => {
  readBtn.disabled = true;
  readBtn.textContent = '読み取り中...';
  showOutput('Teams からメッセージを取得中...', 'normal');

  try {
    const res = await chrome.runtime.sendMessage({ action: 'READ_MESSAGES' });
    if (res?.success && res?.data) {
      const { context, messages, method } = res.data;
      const summary = [
        `取得方法: ${method}`,
        `コンテキスト: ${context.channelName || context.chatTitle || context.pageTitle}`,
        `メッセージ数: ${messages.length}件`,
        '',
        ...messages.slice(0, 5).map((m, i) =>
          `[${i + 1}] ${m.sender}: ${m.body.slice(0, 60)}${m.body.length > 60 ? '...' : ''}`
        ),
        messages.length > 5 ? `... 他 ${messages.length - 5} 件` : '',
      ].filter(Boolean).join('\n');

      showOutput(summary, 'success');

      // ブリッジサーバーへ転送（未起動でも問題なし）
      fetch(`${BRIDGE_URL}/messages`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(res.data),
      }).catch(() => {});
    } else {
      showOutput(`エラー: ${res?.error || 'メッセージ取得失敗'}`, 'error');
    }
  } catch (err) {
    showOutput(`例外: ${err.message}`, 'error');
  } finally {
    readBtn.disabled = false;
    readBtn.textContent = 'メッセージを読み取る';
  }
});

// ========== DOM 構造調査 ==========

const inspectBtn = document.getElementById('inspect-btn');

inspectBtn.addEventListener('click', async () => {
  inspectBtn.disabled = true;
  inspectBtn.textContent = '調査中...';
  showOutput('Teams の DOM 構造を調査中...', 'normal');

  try {
    const res = await chrome.runtime.sendMessage({ action: 'INSPECT_DOM' });
    if (res?.success && res?.data) {
      const d = res.data;
      const lines = [
        `=== DOM 構造調査結果 ===`,
        `data-tid 要素数: ${d.summary.dataTidCount} (ユニーク: ${d.summary.uniqueTids})`,
        '',
        `--- メッセージ候補 ---`,
        ...d.messageCandidate.map(c =>
          `  ${c.selector}: ${c.count}件 classes=[${c.sampleClasses.join(', ')}]`
        ),
        '',
        `--- 送信者候補 ---`,
        ...d.senderCandidate.map(c =>
          `  ${c.selector}: ${c.count}件 例=${c.samples.join(', ')}`
        ),
        '',
        `--- タイムスタンプ候補 ---`,
        ...d.timestampCandidate.map(c =>
          `  ${c.selector}: ${c.count}件 例=${c.samples.map(s => s.text || s.datetime).join(', ')}`
        ),
        '',
        `--- 返信ボックス候補 ---`,
        ...d.replyBoxCandidate.map(c =>
          `  ${c.selector}: ${c.count}件 tag=${c.sampleTag}`
        ),
      ];

      showOutput(lines.join('\n'), 'success');

      // ブリッジサーバーにも送信（起動していれば）
      fetch(`${BRIDGE_URL}/dom-inspect`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(res.data),
      }).catch(() => {});

      // コンソールにも完全なデータを出力
      console.log('[M365 AI Bridge] DOM調査結果:', JSON.stringify(res.data, null, 2));
    } else {
      showOutput(`エラー: ${res?.error || 'DOM調査失敗'}`, 'error');
    }
  } catch (err) {
    showOutput(`例外: ${err.message}`, 'error');
  } finally {
    inspectBtn.disabled = false;
    inspectBtn.textContent = 'DOM構造を調査';
  }
});

// ========== 起動 ==========
checkStatus();
