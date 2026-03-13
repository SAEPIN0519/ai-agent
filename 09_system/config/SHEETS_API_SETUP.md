# Google Sheets API セットアップ手順

週次日報分析スクリプトがGoogle Sheetsを読み取るために必要な設定です。
**所要時間: 約10分**

---

## 手順

### 1. Google Cloud Console でプロジェクト作成

1. https://console.cloud.google.com/ にアクセス
2. 上部の「プロジェクトを選択」→「新しいプロジェクト」
3. プロジェクト名: `siftai-report-analyzer`（任意）
4. 「作成」をクリック

### 2. Google Sheets API を有効化

1. 左メニュー「APIとサービス」→「ライブラリ」
2. 「Google Sheets API」を検索して「有効にする」

### 3. サービスアカウント作成

1. 左メニュー「APIとサービス」→「認証情報」
2. 「認証情報を作成」→「サービスアカウント」
3. サービスアカウント名: `sheets-reader`（任意）
4. 「作成して続行」→ ロールは不要 →「完了」

### 4. JSONキーをダウンロード

1. 作成したサービスアカウントをクリック
2. 「キー」タブ →「鍵を追加」→「新しい鍵を作成」
3. 種類: **JSON** を選択 →「作成」
4. ダウンロードされたJSONファイルを以下に配置:

```
09_system/config/google_service_account.json
```

### 5. スプレッドシートにサービスアカウントを共有

1. ダウンロードしたJSONファイルを開く
2. `"client_email"` の値をコピー（例: `sheets-reader@xxx.iam.gserviceaccount.com`）
3. 会員カルテのスプレッドシートを開く
4. 右上の「共有」→ コピーしたメールアドレスを追加
5. 権限: **閲覧者**（読み取り専用でOK）

### 6. 動作確認

```bash
python 09_system/週次日報分析_discord.py
```

dry-run（送信なし）で分析結果が表示されればOK！

---

## セキュリティ注意

- `google_service_account.json` は **絶対にGitにコミットしない**
- `.gitignore` に追加済み（確認してください）
