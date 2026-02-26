# M365 AI Bridge Extension

Microsoft 365（Teams、SharePoint等）と AI アシスタント（Claude Code等）を接続する Chrome Extension。

管理者権限・API登録不要。ブラウザ上で動作するユーザー主体のアーキテクチャ。

## アーキテクチャ

```
Teams (browser)
  ↕  DOM操作
Content Script (teams-reader.js)
  ↕  chrome.runtime.sendMessage
Service Worker (background)
  ↕  HTTP POST/GET
Bridge Server (localhost:3765)
  ↕  fetch / CLI
Claude Code (AI)
```

## フロー

1. **読み取り**: Teams のチャット画面を開く → Extension ポップアップで「メッセージを読み取る」→ ブリッジサーバー経由で Claude Code が取得
2. **返信**: Claude Code が返信案を生成 → `POST /reply` → Extension が返信フォームに挿入 → ユーザーが送信ボタンを押す

## セットアップ（Phase 1 PoC）

### 1. Extension をインストール

1. Chrome で `chrome://extensions` を開く
2. 「デベロッパーモード」をオン
3. 「パッケージ化されていない拡張機能を読み込む」→ このフォルダを選択

### 2. ブリッジサーバーを起動

```bash
node native/bridge-server.js
```

### 3. Teams を開く

`https://teams.microsoft.com` を開き、読み取りたいチャネル/チャットを表示する。

### 4. Extension でメッセージ取得

ツールバーの Extension アイコンをクリック → 「メッセージを読み取る」

### 5. Claude Code からアクセス

```bash
# メッセージ取得
curl http://localhost:3765/messages | jq '.messages[] | {sender, body}'

# 返信テキストを送信（Extension が自動挿入）
curl -X POST http://localhost:3765/reply \
  -H 'Content-Type: application/json' \
  -d '{"text": "ご確認いただきありがとうございます。"}'
```

## API リファレンス

| Method | Endpoint | 説明 |
|--------|----------|------|
| GET | `/health` | ヘルスチェック |
| GET | `/messages` | 最新の Teams メッセージ取得 |
| POST | `/messages` | Extension からメッセージ受信（内部用） |
| POST | `/reply` | 返信テキストをキューに追加 |
| GET | `/pending-reply` | Extension が返信テキストをポーリング取得（内部用） |
| GET | `/status` | サーバー状態とバッファ情報 |

## ロードマップ

### Phase 1（現在）: PoC
- [x] Content Script による Teams DOM 読み取り
- [x] ローカルブリッジサーバー（HTTP）
- [x] Extension ポップアップ UI
- [ ] Teams DOM セレクタの検証・調整
- [ ] 実際のメッセージ取得テスト

### Phase 2: 安定化
- [ ] WebSocket によるリアルタイムプッシュ（ポーリング廃止）
- [ ] 複数チャンネル・チャットの切り替え対応
- [ ] メッセージのページネーション（スクロールで過去取得）
- [ ] ファイル添付・リンクの取得

### Phase 3: 拡張
- [ ] SharePoint ドキュメントライブラリの読み取り
- [ ] Outlook カレンダーとの連携
- [ ] Claude Code MCP サーバーとの直接統合

## 技術メモ

### なぜ Chrome Extension か

- **管理者不要**: Azure Entra ID へのアプリ登録・管理者同意が不要
- **ゲストテナント対応**: ゲストとして参加している組織の Teams でも動作
- **ToS リスクなし**: ブラウザの正規UI を使うため Playwright スクレイピングと異なりリスクが低い
- **ユーザー主体**: 送信ボタンは必ずユーザーが押す。AI が勝手に送信しない

### DOM セレクタについて

Teams の UI 更新でセレクタが壊れることがある。`content/teams-reader.js` の `SELECTORS` オブジェクトを更新することで対応。ブラウザの DevTools で確認可能。

### セキュリティ

- ブリッジサーバーは `127.0.0.1`（ローカルホスト）のみリッスン
- ネットワーク外部からのアクセス不可
- Teams の認証トークンは取得・送信しない
