# M365 AI Bridge Extension

Microsoft 365（Teams、SharePoint等）と AI アシスタント（Claude Code等）を接続する Chrome Extension。

管理者権限・API登録不要。ブラウザ上で動作するユーザー主体のアーキテクチャ。

## アーキテクチャ

```
Claude Code (AI)
  ↕  MCP stdio（JSON-RPC）
native/mcp-server.js  ─── MCP Tools: teams_read_messages / teams_queue_reply / teams_get_status
  ↕  HTTP (localhost:3765, 内部通信)
Chrome Extension (background service worker)
  ↕  chrome.runtime.sendMessage
Content Script (teams-reader.js)
  ↕  DOM 操作
Teams (browser)
```

`native/mcp-server.js` は 1 プロセスで 2 つのインターフェースを担う:
- **MCP stdio サーバー**: Claude Code がツールとして呼び出す
- **HTTP サーバー** (localhost:3765): Chrome Extension がデータを投稿・ポーリングする

## フロー

1. **読み取り**: Teams のチャット画面を開く → Extension ポップアップで「メッセージを読み取る」→ MCP ツール `teams_read_messages()` で Claude Code が取得
2. **返信**: Claude Code が返信案を生成 → MCP ツール `teams_queue_reply(text)` → Extension が返信フォームに自動挿入 → ユーザーが送信ボタンを押す

## セットアップ

### 1. Extension をインストール

1. Chrome で `chrome://extensions` を開く
2. 「デベロッパーモード」をオン
3. 「パッケージ化されていない拡張機能を読み込む」→ このフォルダを選択

### 2. Claude Code の .mcp.json に登録

```json
{
  "mcpServers": {
    "teams": {
      "type": "stdio",
      "command": "node",
      "args": ["/path/to/m365-ai-bridge-extension/native/mcp-server.js"]
    }
  }
}
```

Claude Code 起動時に MCP サーバーが自動起動する。

### 3. Teams を開く

`https://teams.microsoft.com` を開き、読み取りたいチャネル/チャットを表示する。

### 4. Extension でメッセージ取得

ツールバーの Extension アイコンをクリック → 「メッセージを読み取る」

### 5. Claude Code からアクセス

```
teams_read_messages()          # メッセージ一覧取得
teams_read_messages(limit=10)  # 最新10件に絞る
teams_queue_reply("返信文")    # 返信フォームに挿入
teams_get_status()             # 接続状態確認
```

## MCP ツールリファレンス

| ツール名 | 引数 | 説明 |
|---------|------|------|
| `teams_read_messages` | `limit?` (number) | 現在開いているチャンネル/チャットのメッセージ取得 |
| `teams_queue_reply` | `text` (string) | 返信フォームにテキストを挿入（送信はユーザーが行う） |
| `teams_get_status` | なし | MCP サーバー・Extension の接続状態確認 |

## HTTP API（Extension との内部通信）

Extension が直接通信するエンドポイント（Claude Code からは MCP ツール経由で使うこと）。

| Method | Endpoint | 説明 |
|--------|----------|------|
| GET | `/health` | ヘルスチェック |
| POST | `/messages` | Extension からメッセージデータを受信 |
| GET | `/pending-reply` | Extension が返信テキストをポーリング取得 |
| GET | `/status` | サーバー状態とバッファ情報 |

## テスト

```bash
npm install
npx playwright test
```

19 件のテストがすべて通ることを確認:
- `bridge-server.test.js`: HTTP API 6件
- `content-script.test.js`: DOM読み取り 4件
- `mcp-server.test.js`: MCP ツール呼び出し 9件

## ロードマップ

### Phase 1（現在）: PoC
- [x] Content Script による Teams DOM 読み取り
- [x] MCP サーバー（stdio）で Claude Code と接続
- [x] Extension ポップアップ UI
- [x] Playwright テスト（19件）
- [x] Teams DOM セレクタの実機検証・調整（2026-02-27: teams.cloud.microsoft 新UI対応）
- [ ] エンドツーエンドの動作確認（MCP経由でClaude Codeからメッセージ取得）

### Phase 2: 安定化
- [ ] Extension → MCP サーバーへのリアルタイムプッシュ（ポーリング廃止）
- [ ] 複数チャンネル・チャットの切り替え対応
- [ ] メッセージのページネーション（スクロールで過去取得）
- [ ] ファイル添付・リンクの取得

### Phase 3: 拡張
- [ ] SharePoint ドキュメントライブラリの読み取り
- [ ] Outlook カレンダーとの連携

## 技術メモ

### なぜ Chrome Extension か

- **管理者不要**: Azure Entra ID へのアプリ登録・管理者同意が不要
- **ゲストテナント対応**: ゲストとして参加している組織の Teams でも動作
- **ToS リスクなし**: ブラウザの正規 UI を使うため Playwright スクレイピングと異なりリスクが低い
- **ユーザー主体**: 送信ボタンは必ずユーザーが押す。AI が勝手に送信しない

### DOM セレクタについて

Teams の UI 更新でセレクタが壊れることがある。`content/teams-reader.js` の `SELECTORS` オブジェクトを更新することで対応。ブラウザの DevTools で確認可能。

### セキュリティ

- MCP サーバーの HTTP 部分は `127.0.0.1`（ローカルホスト）のみリッスン
- ネットワーク外部からのアクセス不可
- Teams の認証トークンは取得・送信しない
