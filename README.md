# HARIBOW単独公演 興味関心アンケート

HARIBOW初の単独公演（2026/9/29・めぐろパーシモンホール）に向けた、観客候補向けの興味関心アンケートWebアプリ（Google Apps Script + GitHub Pages）。

他のHARIBOW関連リポジトリ（オーディション系）とは無関係の、独立したツールです。

## 配布用URL（実際にはこちらを案内する）

- 国内向け（日本語）: https://junyaunitydev.github.io/haribow-solo-show-survey/
- 海外向け（英語・配信需要調査）: https://junyaunitydev.github.io/haribow-solo-show-survey/en.html

GASのWebアプリ直URL(`.../exec`)は一部ユーザーでアクセスエラーが多発するため、外部配布は上記のGitHub Pages URLを使う。GAS側の`.../exec`は`doPost`エンドポイントとしてのみ使用（`index.html`/`en.html`が内部で`fetch()`している）。

## 構成

- `Code.js` — GAS本体。`doGet`は内部確認用（`solo_interest.html`/`solo_interest_en.html`をHtmlServiceで直配信）、`doPost`が本番の受け口（GitHub Pagesの静的サイトから`fetch()`で送られてくるJSONを受けてSheetsに書き込む）
- `index.html` — **GitHub Pagesで配信される国内向け日本語フォーム本体**（`fetch()`で`doPost`にPOST）
- `en.html` — **GitHub Pagesで配信される海外向け英語フォーム本体**（同上）
- `solo_interest.html` / `solo_interest_en.html` — GAS直配信版（`google.script.run`使用・内部確認用のフォールバック、外部配布には使わない）
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

**国内向け（`興味関心_国内`）**: A6以降が回答データ領域(〜105行目)、10列（氏名/経由/認知度/興味度/価格帯/学生か/観たい公演/居住地/配信視聴意向/回答日）

**海外向け（`興味関心_海外EN`）**: A6以降が回答データ領域(〜55行目)、6列（Name/Heard/Watch(live)/Watch(recorded)/Location/Response date）

## 初回デプロイ後の注意

新規作成したGASプロジェクトは、匿名アクセス（`ANYONE_ANONYMOUS`）が有効になるまでに、オーナーが一度Webアプリのexec URLをブラウザで開いて権限を承認する必要がある場合があります。
