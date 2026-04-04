# AIエージェント指示書：週次報告 自動生成・Slack投稿エージェント

## 発令者
SAEPIN（COO）

## 担当エージェント
SHIFTAIレポートエージェント

## ミッション
毎週月曜9:30に、プロプレミアムTEAMの週次総括をSlackに自動投稿する。
冴香（野村冴香）名義で投稿すること。

---

## 毎週のタスクフロー

### ステップ1：カルテ同期（日曜夜 or 月曜朝）
```bash
python3 09_system/カルテ同期.py
```
- Google Sheets（カルテ）から最新の会員データを取得
- `member_roster.md` と `kpi_progress.md` が自動更新される

### ステップ2：週次データ集計
カルテから以下のデータを集計する（対象：直前の月曜〜日曜）

| データ | 取得元シート | 集計内容 |
|---|---|---|
| 総稼働時間 | DB_Daily-Report | 稼働時間の合計・前週比 |
| 日報提出 | DB_Daily-Report | 提出者数・件数・前週比 |
| 日報提出数TOP5 | DB_Daily-Report | 同率は横並び |
| 稼働時間TOP5 | DB_Daily-Report | 同率は横並び |
| FB会参加数TOP5 | DB_FB-Report | 同率は横並び |
| アクティブ傾向率 | DB_Daily-Report | アクティブ/停滞の比率 |
| 非アクティブ一覧 | DB_Daily-Report × DB_Pro-Members | 今週日報・FB会ともにゼロのPRO会員 |
| 今週のマネタイズ | DB_Weekly-Report | 今週分の収益報告（金額・内容） |
| 累計マネタイズTOP5 | DB_Weekly-Report | 全期間の累計ランキング |

### ステップ3：週報テキスト生成
以下のフォーマットで `03_clients/SIFTAI/プロプレミアムTEAM/週次報告/YYYY-MM-DD.md` を作成する。

```
【週次総括】PRO + PREMIUM 合算（M/DD〜M/DD）

━━━━━━━━━━━━━━
*数字のハイライト*
━━━━━━━━━━━━━━

- 総稼働時間：XXXh（前週比 ±XX%）
- 日報提出：XX名 / XXX件（前週比 ±XX件）
- マネタイズ累計：XX,XXX円（前週比 ±X,XXX円）
- アクティブ傾向率：XX%（前週比 ±Xpt）
- 非アクティブ：XX名 / XX名（XX%）

[1〜2行の総評コメント]

*日報提出数TOP:*
  [同率は横並び]

*稼働時間TOP5:*
  [5位まで]

*実践フィードバック会参加TOP5:*
  [同率は横並び]

━━━━━━━━━━━━━━
*今週のマネタイズ成果*
━━━━━━━━━━━━━━

[今週の収益報告者リスト]

━━━━━━━━━━━━━━
*累計マネタイズTOP5*
━━━━━━━━━━━━━━

[累計ランキング]

━━━━━━━━━━━━━━
*今週のアクション（担当別）*
━━━━━━━━━━━━━━

*【全員共通】*
[全体向けの連絡事項]

*【安永さん（上位コンシェルジュ）】*
[全体進捗管理の依頼]

*【安部さん】*
[担当会員: 坂井俊章、矢三やよい ほかPREMIUM担当]

*【山本さん】*
[担当会員: 石渡一雄 ほかPREMIUM担当]

*【岩間さん】*
[担当会員: 藤木智咲、村上真樹子、築山優希、内藤友香子、中澤みどり、片寄孝康、立田素子]

*【冴香・秘書チーム】*
[担当会員: 伊藤紘二、吉田比奈、山田真由美、島田英敏、池田勝典、甲村順子、谷山真一]

*【組織課題（次回MTGで議論）】*
[制度的な課題・検討事項]
```

### ステップ4：Slack予約投稿
```bash
python3 09_system/slack_weekly_post.py --schedule
```
- 投稿先: `community_ss_pro-premium_plan`（ID: C0AC8404FPE）
- 投稿時間: 毎週月曜 9:30
- 投稿者: 野村冴香（ユーザートークン使用）
- トークン: `09_system/config/slack_user_token.txt`

---

## コンシェルジュ担当マッピング（DB_Interviewsより）

| コンシェルジュ | 役割 | 担当会員 |
|---|---|---|
| 安永智也 | 上位コンシェルジュ | 安部・山本の進捗管理 |
| 安部友博 | コンシェルジュ | 坂井俊章、矢三やよい + PREMIUM担当 |
| 山本大輔 | コンシェルジュ | 石渡一雄 + PREMIUM担当 |
| 岩間慧大 | 秘書チーム | 藤木智咲、村上真樹子、築山優希、内藤友香子、中澤みどり、片寄孝康、立田素子 |
| 野村冴香 | 秘書チーム | 伊藤紘二、吉田比奈、山田真由美、島田英敏、池田勝典、甲村順子、谷山真一 |

※ 担当変更があった場合は DB_Interviews の最新データを参照して更新する

---

## 注意事項

- 「退会意思確認」という表現は使わない
- マネタイズは「今週」と「累計」の両方を出す
- TOP5ランキングは同率を横並びにする
- コンシェルジュと会員の紐付けは DB_Interviews に基づく（勝手に変えない）
- 週報テキストのSlack書式（*太字*、━━ライン）を維持する

---

## 関連ファイル

- カルテ同期: `09_system/カルテ同期.py`
- Slack投稿: `09_system/slack_weekly_post.py`
- OAuth認証: `09_system/google_auth_setup.py`
- 週次報告保存先: `03_clients/SIFTAI/プロプレミアムTEAM/週次報告/`
- 過去レポート: `03_clients/SIFTAI/プロプレミアムTEAM/週次報告/archive/`

---

---

## 週次業務フロー（改定版 2026-03-31）

### 全体スケジュール

| 曜日・時間 | 内容 | コマンド |
|---|---|---|
| 日曜 21:00 | HTMLレポート生成 + git push + 冴香にLINE通知 | `python3 weekly_member_analysis.py --combined --generate-only` |
| 月曜（冴香の指示後） | Slackに週次レポート投稿 | `python3 weekly_member_analysis.py --combined --slack-only` |
| 月曜 10:00 | Discord用 週間MVP投稿文を生成 + 冴香にLINE通知 | （後述のDiscord MVP生成スクリプト） |
| 月曜（冴香の指示後） | DiscordにMVP投稿 | 冴香がDiscordに手動投稿 or Bot投稿 |

### 日曜 21:00 — HTMLレポート生成
1. Google Sheetsから最新データ取得
2. PRO / PREMIUM / 合算 / チャンネルログ分析の4つのHTMLを生成
3. git push（GitHub Pagesに反映）
4. 冴香のLINEにGitHub PagesのURLを送信して確認依頼

### 月曜 10:00 — Discord週間MVP投稿文の生成
以下のランキングを自動生成し、冴香のLINEに送信する:

**投稿内容:**
- 稼働時間 TOP3（同率は同じメダル）
- 日報提出 TOP3（7回以上＝毎日提出で1位扱い）
- FB会参加 TOP3（同率は同じメダル）

**ルール:**
- 氏名は `DB_All-Members` の `Discord ID`（col[6]）を使い `<@ID>` メンション形式にする
- 氏名の表記揺れ対策として全角/半角スペースをstripする
- PRO/PREMIUM両方に所属するメンバーは重複カウントしない
- 冒頭に `@PRO`（ロールメンション）を付ける（ロールID確認後）
- 冴香が確認・修正してからDiscordに投稿する

### 通知手段について
- LINE Notifyは2025年3月末でサービス終了済み
- 代替手段を冴香と相談して決定する（LINE Messaging API / Discord DM / メール）

---

## ステップ5：HTMLレポート生成（4月以降追加）

週報テキスト（MD）に加えて、HTMLビジュアルレポートも自動生成する。

### 生成するHTMLファイル
- `週次報告/PRO_M月D日-M月D日.html` — PRO会員詳細レポート
- `週次報告/PREMIUM_M月D日-M月D日.html` — PREMIUM会員詳細レポート
- `週次報告/会員分析（日報）M月D日-M月D日.html` — 日報分析レポート
- `週次報告/会員分析（日報）M月D日-M月D日.md` — 同内容のMD版

### HTML投稿フロー
1. HTMLファイルを生成
2. Slackに週報テキスト投稿（9:30）
3. HTMLファイルをSlackにアップロード（同スレッドに添付）

### Slackファイルアップロード
```python
from slack_sdk import WebClient
client = WebClient(token=user_token)
# スレッドに添付
client.files_upload_v2(
    channel=CHANNEL_ID,
    file=html_path,
    title="週次レポート PRO M/DD〜M/DD",
    thread_ts=post_ts,  # 週報テキストのスレッドに添付
)
```

---

## Windows連携ルール

### git pull / push の運用
- **作業開始前に必ず `git pull origin main`** を実行する
- 週報生成後は `git add` → `git commit` → `git push origin main` で同期する
- Windows側も同じリポジトリを使用しているため、コンフリクト防止のために必ずpull→作業→pushの順序を守る

### ファイル配置ルール
| ファイル | パス | git管理 |
|---|---|---|
| 週報テキスト | `週次報告/YYYY-MM-DD.md` | ○ |
| HTMLレポート | `週次報告/PRO_*.html` 等 | ○ |
| 過去レポート | `週次報告/archive/YYYY-MM/` | ○ |
| OAuth認証 | `09_system/config/google_oauth_*.json/.pickle` | × (.gitignore) |
| Slackトークン | `09_system/config/slack_*_token.txt` | × (.gitignore) |

### 月替わり処理
月が変わったら、前月のHTMLレポートを `週次報告/archive/YYYY-MM/` に移動する。

---

発令日: 2026-03-29
更新日: 2026-03-31（週次フロー改定・Discord MVP投稿追加・Slack投稿前確認フロー導入）
発令者: SAEPIN（COO）
