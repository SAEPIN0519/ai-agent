# AIエージェント定例業務：SIFTAI 週次レポート生成・配信

## 役割

毎週月曜の朝、会員カルテから最新データを取得し、週次分析レポート（PRO / PREMIUM / 合算 / チャットログ分析）を生成してSlackに配信する。

---

## 実行スケジュール

| 順序 | 時刻 | スクリプト | 内容 |
|---|---|---|---|
| 1 | 7:00 | `sync_member_data.py` | ローカルMD（member_roster / kpi_progress）をカルテから最新化 |
| 2 | 9:32 | `weekly_member_analysis.py --combined` | HTML生成 + Slack投稿（4ファイル） |

※ タスクスケジューラ名: `SIFTAI_Weekly_Member_Analysis`（Mac PCで実行）

---

## 前提条件

- Python 3.10+
- 必須ライブラリ: `google-auth`, `google-auth-httplib2`, `requests`
- 認証ファイル: `09_system/config/google_service_account.json`（サービスアカウント）
- Slackトークン: `09_system/config/slack_user_token.txt`（冴香名義のユーザートークン）
- データソース: Google Sheets 会員カルテ（ID: `1EHKgmE1d7T5N9GbT_QJV72Q8Dh91iNysHHkz6LEWJy4`）

---

## ステップ1: データ同期（sync_member_data.py）

```bash
cd 09_system/scripts
python sync_member_data.py
```

**処理内容:**
1. Google Sheets「DB_Member」シートから会員一覧を取得
2. `member_roster.md` を最新のメンバー情報で上書き
3. Google Sheets「DB_KPI」シートからKPIデータを取得
4. `kpi_progress.md` を最新のKPI進捗で上書き

**出力ファイル:**
- `03_clients/SIFTAI/プロプレミアムTEAM/会員管理/member_roster.md`
- `03_clients/SIFTAI/プロプレミアムTEAM/コーチング/kpi_progress.md`

**正常完了の確認:**
- 「同期完了」のログが出力される
- 上記2ファイルの更新日時が実行時刻になっている

---

## ステップ2: 週次レポート生成 + Slack投稿（weekly_member_analysis.py）

```bash
cd 09_system/scripts
python weekly_member_analysis.py --combined
```

**処理内容:**
1. Google Sheets から日報データ（DB_Daily-Report）を取得
2. 先週月曜〜日曜の期間を自動算出
3. PRO / PREMIUM それぞれの分析HTMLを生成
4. `--combined` フラグにより合算レポート（重複除外）も生成
5. `channel_log_analysis.py` を内部importしてPREMIUM専用ルームのチャットログ分析HTMLを生成
6. 4つのHTMLファイルをSlackにアップロード
7. パーマリンク付きメッセージを `community_ss_pro-premium_plan` チャンネルに投稿

**生成される4ファイル:**

| ファイル名 | 内容 |
|---|---|
| `PRO_{期間}.html` | PROプラン会員の週次分析 |
| `PREMIUM_{期間}.html` | PREMIUMプラン会員の週次分析 |
| `合算_{期間}.html` | PRO+PREMIUM合算（重複除外） |
| `チャンネルログ分析.html` | PREMIUM専用ルーム チャットログ分析（全期間） |

**保存先:** `03_clients/SIFTAI/プロプレミアムTEAM/会員管理/週次レポート/`

**Slack投稿先:** `community_ss_pro-premium_plan`（ID: C0AC8404FPE）

**Slack投稿形式:**
- 冴香名義（ユーザートークン）で投稿
- 週次レポート（PRO / PREMIUM / 合算）のリンク + チャットログ分析リンクを分けて表示

**正常完了の確認:**
- ターミナルに「Slack投稿完了」が表示される
- Slackチャンネルにメッセージ + 4ファイルが届いている
- ローカルの週次レポートフォルダに4つのHTMLが保存されている

---

## テスト実行（dry-run）

Slack投稿せずにHTML生成だけ確認したい場合:

```bash
python weekly_member_analysis.py --combined --dry-run
```

---

## エラー時の対処

| エラー | 原因 | 対処 |
|---|---|---|
| `google_service_account.json が見つかりません` | 認証ファイル未配置 | `09_system/config/` にサービスアカウントJSONを配置 |
| `Sheets APIエラー 403` | スプレッドシートへの権限なし | サービスアカウントのメールをカルテに共有設定 |
| `slack_user_token.txt が見つかりません` | Slackトークン未配置 | `09_system/config/` にユーザートークンを配置 |
| `Slack投稿エラー` | トークン期限切れ or チャンネルID変更 | トークン更新 or チャンネルID確認 |
| `channel_log_analysis import エラー` | 依存スクリプトの問題 | `channel_log_analysis.py` の構文エラーを確認 |

---

## 週次チェックリスト（AIエージェント用）

毎週月曜、以下を順番に確認:

- [ ] `sync_member_data.py` が正常完了したか（member_roster.md / kpi_progress.md 更新）
- [ ] `weekly_member_analysis.py --combined` が正常完了したか
- [ ] Slackにメッセージ + 4ファイルが投稿されたか
- [ ] HTMLの内容に異常がないか（会員数0、期間ズレなど）
- [ ] エラーログがあれば冴香に報告

---

## 関連ファイル

- スクリプト: [sync_member_data.py](../../../../09_system/scripts/sync_member_data.py)
- スクリプト: [weekly_member_analysis.py](../../../../09_system/scripts/weekly_member_analysis.py)
- スクリプト: [channel_log_analysis.py](../../../../09_system/scripts/channel_log_analysis.py)
- 認証: `09_system/config/google_service_account.json`
- Slackトークン: `09_system/config/slack_user_token.txt`
- 出力先: `03_clients/SIFTAI/プロプレミアムTEAM/会員管理/週次レポート/`
- Slackチャンネル: `community_ss_pro-premium_plan`（C0AC8404FPE）

---

## 次のアクション（拡張予定）

- [ ] Phase 2: 日報・面談ログとの突合分析（チャット活用度 × 会員成長フェーズの相関）
- [ ] チャットゼロの会員へのフォロー施策を自動化
- [ ] 週次レポートに前週比較（増減トレンド）を追加
