# HARIBOW単独公演 興味関心アンケート

HARIBOW初の単独公演（2026/9/29・めぐろパーシモンホール）に向けた、観客候補向けの興味関心アンケートWebアプリ（Google Apps Script / HtmlService）。

他のHARIBOW関連リポジトリ（オーディション系）とは無関係の、独立したツールです。

## 構成

- `Code.js` — GAS本体。`doGet`でページ出し分け、`google.script.run`経由でSheetsに回答を書き込む
- `solo_interest.html` — 国内向け日本語フォーム
- `solo_interest_en.html` — 海外SNSフォロワー向け簡易英語フォーム（配信の需要調査用）
- `appsscript.json` — マニフェスト（Webアプリとして誰でもアクセス可能に設定）
- `Config.js` — **リポジトリには含まれません**（`.gitignore`対象）。書き込み先スプレッドシートIDを自分で設定してください

## セットアップ

```bash
npm install -g @google/clasp   # 未インストールなら
clasp login

cp Config.example.js Config.js
# Config.js を開いて SOLO_SS_ID に書き込み先スプレッドシートIDを入れる

clasp create --type standalone --title "任意のタイトル" --rootDir .
# 既存のGASプロジェクトに繋ぐ場合は .clasp.json の scriptId を書き換える

clasp push -f
clasp deploy --description "初回公開"
```

書き込み先スプレッドシートには、以下のヘッダー構成のタブが必要です。

**国内向け（`興味関心_国内`）**: A6以降が回答データ領域(〜105行目)、9列（氏名/経由/認知度/興味度/価格帯/学生か/観たい公演/居住地/回答日）

**海外向け（`興味関心_海外EN`）**: A6以降が回答データ領域(〜55行目)、5列（Name/Heard/Watch/Location/Response date）

## 初回デプロイ後の注意

新規作成したGASプロジェクトは、匿名アクセス（`ANYONE_ANONYMOUS`）が有効になるまでに、オーナーが一度Webアプリのexec URLをブラウザで開いて権限を承認する必要がある場合があります。
