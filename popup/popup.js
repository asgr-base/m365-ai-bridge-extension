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

const outputWrapper = document.getElementById('output-wrapper');
const copyBtn = document.getElementById('copy-btn');

function showOutput(text, type = 'normal') {
  outputWrapper.style.display = 'block';
  output.className = `output ${type}`;
  output.textContent = text;
}

// ========== クリップボードコピー ==========

copyBtn.addEventListener('click', async () => {
  try {
    await navigator.clipboard.writeText(output.textContent);
    copyBtn.textContent = 'OK';
    copyBtn.classList.add('copied');
    setTimeout(() => {
      copyBtn.textContent = 'コピー';
      copyBtn.classList.remove('copied');
    }, 1500);
  } catch {
    // フォールバック: execCommand
    const range = document.createRange();
    range.selectNodeContents(output);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    document.execCommand('copy');
    sel.removeAllRanges();
    copyBtn.textContent = 'OK';
    setTimeout(() => { copyBtn.textContent = 'コピー'; }, 1500);
  }
});

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
        ...messages.slice(0, 5).map((m, i) => {
          let line = `[${i + 1}] ${m.sender}: ${m.body.slice(0, 60)}${m.body.length > 60 ? '...' : ''}`;
          if (m.attachments?.length > 0) {
            line += `\n     添付: ${m.attachments.map(a => a.name).join(', ')}`;
          }
          return line;
        }),
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
        `--- data-tid 上位50件 ---`,
        ...(d.dataTidElements || []).slice(0, 50).map(e => `  ${e.tid}: ${e.count}件`),
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
        '',
        `--- ファイル添付候補 ---`,
        ...(d.fileCandidate || []).map(c =>
          `  ${c.selector}: ${c.count}件\n` +
          c.samples.map(s => `    href=${(s.href || '').slice(0, 60)} | ${s.text?.slice(0, 50) || ''}`).join('\n')
        ),
        ...(d.fileDetail || []).flatMap((f, i) => [
          `  --- file[${i}]: ${f.text?.slice(0, 50) || '?'} ---`,
          `    aria: ${f.ariaLabel || 'null'} role: ${f.role || 'null'}`,
          ...(f.imgs || []).map(im => `    img src=${im.src?.slice(0, 80)} alt=${im.alt}`),
          ...(f.buttons || []).map(b => `    btn aria=${b.ariaLabel} | ${b.text}`),
          ...(f.urlElements || []).map(u => `    [${u.tag}] href=${u.href} url=${u.dataUrl} src=${u.src}`),
          f.urlElements?.length === 0 ? '    (URLを持つ子要素なし)' : '',
        ]),
        '',
        `--- メンション候補 ---`,
        ...(d.mentionCandidate || []).map(c =>
          `  ${c.selector} [${c.scope}]: ${c.count}件\n` +
          (c.samples || []).map(s =>
            `    text="${s.text}" classes="${s.classes}" attrs=[${s.attrs}]` +
            (s.parentTag ? ` parent=${s.parentTag}(${s.parentTid || ''})` : '')
          ).join('\n')
        ),
        ...(d.messageBodySpans || []).length > 0 ? [
          '',
          `--- message-body[0] 内 span 要素 ---`,
          ...(d.messageBodySpans || []).map(s =>
            `  text="${s.text}" classes="${s.classes}" id="${s.id}" attrs=[${s.attrs}] children=${s.childCount}`
          ),
        ] : [],
        ...(d.messageBodySpans2 || []).length > 0 ? [
          '',
          `--- message-body[1] 内 span 要素 ---`,
          ...(d.messageBodySpans2 || []).map(s =>
            `  text="${s.text}" classes="${s.classes}" attrs=[${s.attrs}]`
          ),
        ] : [],
        '',
        `--- DM id候補（上位15件） ---`,
        ...(d.dmIdElements || []).map(e =>
          `  [${e.tag}] id=${e.id} tid=${e.dataTid} parent=${e.parentDataTid} | ${e.textSample}`
        ),
        '',
        `--- フレーム情報 ---`,
        `  URL: ${d.summary.url || '不明'}`,
        `  frameType: ${d.summary.frameType || '不明'}`,
        `  iframes: ${d.summary.iframeCount ?? '?'}件`,
        ...(d.iframes || []).slice(0, 10).map(f =>
          `    src=${f.src.slice(0, 80)} id=${f.id}`
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

// ========== トークン調査 ==========

const tokenBtn = document.getElementById('token-btn');

tokenBtn.addEventListener('click', async () => {
  tokenBtn.disabled = true;
  tokenBtn.textContent = '調査中...';
  showOutput('Teams のトークンストレージを調査中...', 'normal');

  try {
    const res = await chrome.runtime.sendMessage({ action: 'INSPECT_TOKEN' });
    if (res?.success && res?.data) {
      const d = res.data;
      const lines = [
        `=== トークンストレージ調査結果 ===`,
        `Graph Token 検出: ${d.graphTokenFound ? 'YES' : 'NO'}`,
        '',
        `--- sessionStorage (${d.sessionStorage.length}件) ---`,
        ...d.sessionStorage.map(e =>
          e.error ? `  エラー: ${e.error}` :
          `  ${e.key}\n    長さ=${e.valueLength} JWT=${e.looksLikeJwt} preview=${e.valuePreview}`
        ),
        '',
        `--- localStorage (${d.localStorage.length}件) ---`,
        ...d.localStorage.map(e =>
          e.error ? `  エラー: ${e.error}` :
          `  ${e.key}\n    長さ=${e.valueLength} JWT=${e.looksLikeJwt} preview=${e.valuePreview}`
        ),
        '',
        `--- cookies (${d.cookies.length}件) ---`,
        ...d.cookies.map(e =>
          e.error ? `  エラー: ${e.error}` :
          `  ${e.name}\n    長さ=${e.valueLength} JWT=${e.looksLikeJwt}`
        ),
      ];

      if (d.msalTokens?.length > 0) {
        lines.push(
          '',
          `--- MSAL AccessToken (${d.msalTokens.length}件) ---`,
          ...d.msalTokens.map(t =>
            t.error ? `  エラー: ${t.error}` :
            `  ${t.isGraph ? '[Graph]' : t.isSharePoint ? '[SPO]' : '[Other]'} 長さ=${t.tokenLength} JWT=${t.looksLikeJwt}\n` +
            `    target: ${t.target?.slice(0, 80)}\n` +
            `    expires: ${t.expiresOn}`
          ),
        );
      }

      if (d.graphTokenFound && d.tokenSummary) {
        const t = d.tokenSummary;
        lines.push(
          '',
          `--- Graph Token 詳細 ---`,
          `  source: ${t.source}`,
          `  service: ${t.service}`,
          `  target: ${t.target?.slice(0, 100)}`,
          `  tokenLength: ${t.tokenLength}`,
          `  JWT: ${t.looksLikeJwt}`,
          `  expires: ${t.expiresOn}`,
        );
      }

      showOutput(lines.join('\n'), d.graphTokenFound ? 'success' : 'normal');

      console.log('[M365 AI Bridge] トークン調査結果:', JSON.stringify(res.data, null, 2));
    } else {
      showOutput(`エラー: ${res?.error || 'トークン調査失敗'}`, 'error');
    }
  } catch (err) {
    showOutput(`例外: ${err.message}`, 'error');
  } finally {
    tokenBtn.disabled = false;
    tokenBtn.textContent = 'トークン調査';
  }
});

// ========== 起動 ==========
checkStatus();
