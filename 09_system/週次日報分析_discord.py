"""
週次日報分析レポート — 毎週月曜8:00にDiscord/Slackへ配信

機能:
  1. Google Sheets「Form_Responses3」から直近1週間の日報を取得
  2. 提出率TOP3の会員を抽出（1人あたり平均稼働時間も算出）
  3. ポジティブ／ネガティブ発言比率を分析（生の声も引用）
  4. 提出3件以下の非アクティブ会員を検出
  5. Discord Webhookで週次レポートを配信
  6. 非アクティブ会員へのリマインドDMリストを出力

スケジュール:
  - 月曜 8:00  → 週次分析レポート配信（Discord）
  - 月曜 20:00 → 非アクティブ会員リマインドDM（Discord）
  - 火曜       → Slack週次報告（Slack Webhook）
"""

import json
import random
import re
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict

import gspread
from google.oauth2.service_account import Credentials

# Windows環境でのUTF-8出力対応
os.environ["PYTHONIOENCODING"] = "utf-8"
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

try:
    import requests
except ImportError:
    print("エラー: requestsライブラリが必要です。pip install requests を実行してください。")
    sys.exit(1)

# ==========================================
# 設定
# ==========================================
PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = PROJECT_ROOT / "09_system" / "config"
REPORTS_DIR = PROJECT_ROOT / "02_clients" / "SIFTAI" / "プロプレミアムTEAM" / "会員管理" / "週次レポート"

# Google Sheets
SPREADSHEET_ID = "1EHKgmE1d7T5N9GbT_QJV72Q8Dh91iNysHHkz6LEWJy4"
SHEET_NAME = "DB_Daily-Report"
CREDENTIALS_PATH = CONFIG_DIR / "google_service_account.json"

# 列名マッピング（スプレッドシートのヘッダーに合わせる）
COL_TIMESTAMP = "タイムスタンプ"
COL_NAME = "氏名"
COL_DATE = "日付"
COL_HOURS = "稼働時間"
COL_LEARNING = "今日の学び"
COL_TOMORROW = "明日やること"

# Discord Webhook（週次レポート配信用）
DISCORD_WEBHOOK_URL = "https://discord.com/api/webhooks/1480602170368856275/DCDEaDwYhIIH0xGioPjanNbTocElv0FxtbnGo7HTqSDfpR_bn_IPO6Rm2PNhC7p0gUB4"

# Slack Webhook（火曜の週次報告用 — 設定後に有効化）
SLACK_WEBHOOK_URL = ""

# 非アクティブ判定の閾値
MIN_SUBMISSIONS = 3  # 週間提出がこの件数以下で非アクティブ判定

# 感情分析キーワード
POSITIVE_KEYWORDS = [
    "楽しい", "嬉しい", "できた", "理解できた", "進んだ", "成長",
    "面白い", "学べた", "発見", "ありがとう", "感謝", "達成",
    "やりがい", "充実", "順調", "いい感じ", "手応え", "うまくいった",
    "スッキリ", "モチベ", "やる気", "わくわく", "自信",
]
NEGATIVE_KEYWORDS = [
    "難しい", "わからない", "つまずい", "不安", "焦り", "迷い",
    "うまくいかない", "止まっ", "苦戦", "時間がかかる", "挫折",
    "落ち込", "しんどい", "疲れ", "無理", "できない", "混乱",
    "モヤモヤ", "悩", "困っ", "辛い",
]

TODAY = datetime.now()
TODAY_STR = TODAY.strftime("%Y-%m-%d")

# 月曜〜日曜の週サイクルを計算
# weekday(): 月=0, 火=1, ..., 日=6
_current_weekday = TODAY.weekday()
THIS_MONDAY = TODAY - timedelta(days=_current_weekday)
LAST_MONDAY = THIS_MONDAY - timedelta(days=7)
LAST_SUNDAY = THIS_MONDAY - timedelta(days=1)

# ==========================================
# Google Sheets 接続
# ==========================================
def connect_sheets():
    """Google Sheetsに接続してシートを返す"""
    if not CREDENTIALS_PATH.exists():
        print(f"エラー: サービスアカウントの認証ファイルが見つかりません")
        print(f"  → {CREDENTIALS_PATH}")
        print(f"  → セットアップ手順は 09_system/config/SHEETS_API_SETUP.md を参照")
        sys.exit(1)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
    ]
    creds = Credentials.from_service_account_file(str(CREDENTIALS_PATH), scopes=scopes)
    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    sheet = spreadsheet.worksheet(SHEET_NAME)
    return sheet


def connect_spreadsheet():
    """Google Sheetsスプレッドシート全体に接続（複数シート用）"""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
    ]
    creds = Credentials.from_service_account_file(str(CREDENTIALS_PATH), scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


def fetch_weekly_data(sheet):
    """直近1週間の日報データを取得"""
    all_records = sheet.get_all_records()

    # 前週の月曜〜日曜を対象期間とする
    week_start = LAST_MONDAY
    week_end = LAST_SUNDAY

    weekly_data = []
    for row in all_records:
        date_str = str(row.get(COL_DATE, "")).strip()
        if not date_str:
            continue

        # 日付をパース（YYYY/MM/DD形式）
        try:
            row_date = datetime.strptime(date_str, "%Y/%m/%d")
        except ValueError:
            try:
                row_date = datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                continue

        if week_start <= row_date <= week_end:
            weekly_data.append({
                "name": str(row.get(COL_NAME, "")).strip(),
                "date": row_date,
                "hours": str(row.get(COL_HOURS, "")).strip(),
                "learning": str(row.get(COL_LEARNING, "")).strip(),
                "tomorrow": str(row.get(COL_TOMORROW, "")).strip(),
            })

    return weekly_data


# ==========================================
# 分析ロジック
# ==========================================
def parse_hours(hours_str):
    """稼働時間の文字列を数値（時間）に変換"""
    if not hours_str:
        return 0.0

    # 「3:00:00」「2:30:00」形式
    match = re.match(r'(\d+):(\d+)', hours_str)
    if match:
        h = int(match.group(1))
        m = int(match.group(2))
        return h + m / 60.0

    # 「3時間」「2.5時間」形式
    match = re.match(r'([\d.]+)\s*時間', hours_str)
    if match:
        return float(match.group(1))

    # 数値のみ
    try:
        return float(hours_str)
    except ValueError:
        return 0.0


def analyze_submissions(weekly_data):
    """提出回数を会員ごとに集計し、TOP3と非アクティブを抽出"""
    submission_count = defaultdict(int)
    total_hours = defaultdict(float)

    for row in weekly_data:
        name = row["name"]
        if not name:
            continue
        submission_count[name] += 1
        total_hours[name] += parse_hours(row["hours"])

    # 提出率TOP3
    sorted_members = sorted(submission_count.items(), key=lambda x: x[1], reverse=True)
    top3 = []
    for name, count in sorted_members[:3]:
        avg_hours = total_hours[name] / count if count > 0 else 0
        top3.append({
            "name": name,
            "count": count,
            "avg_hours": round(avg_hours, 1),
            "total_hours": round(total_hours[name], 1),
        })

    # 非アクティブ会員（提出3件以下）
    inactive = []
    for name, count in sorted_members:
        if count <= MIN_SUBMISSIONS:
            inactive.append({"name": name, "count": count})

    return {
        "top3": top3,
        "inactive": inactive,
        "all_members": dict(submission_count),
        "total_submissions": sum(submission_count.values()),
        "unique_members": len(submission_count),
    }


def analyze_sentiment(weekly_data):
    """日報テキストからポジティブ／ネガティブ発言を分析"""
    positive_voices = []
    negative_voices = []
    positive_count = 0
    negative_count = 0

    for row in weekly_data:
        text = row["learning"]
        if not text:
            continue

        # ポジティブキーワード検出
        is_positive = any(kw in text for kw in POSITIVE_KEYWORDS)
        is_negative = any(kw in text for kw in NEGATIVE_KEYWORDS)

        if is_positive:
            positive_count += 1
            # 代表的な声を収集（最大5件）
            if len(positive_voices) < 5:
                # テキストが長すぎる場合は80文字で切る
                snippet = text[:80] + "..." if len(text) > 80 else text
                positive_voices.append({"name": row["name"], "voice": snippet})

        if is_negative:
            negative_count += 1
            if len(negative_voices) < 5:
                snippet = text[:80] + "..." if len(text) > 80 else text
                negative_voices.append({"name": row["name"], "voice": snippet})

    total = positive_count + negative_count
    pos_ratio = round(positive_count / total * 100) if total > 0 else 0
    neg_ratio = round(negative_count / total * 100) if total > 0 else 0

    return {
        "positive_count": positive_count,
        "negative_count": negative_count,
        "pos_ratio": pos_ratio,
        "neg_ratio": neg_ratio,
        "positive_voices": positive_voices,
        "negative_voices": negative_voices,
    }


# ==========================================
# レポート生成
# ==========================================
def format_weekly_report(submissions, sentiment, weekly_data):
    """Discord送信用の週次レポートを生成"""
    period_start = LAST_MONDAY.strftime("%m/%d（月）")
    period_end = LAST_SUNDAY.strftime("%m/%d（日）")

    lines = []
    lines.append(f"# 📊 会員日報ウィークリー分析レポート")
    lines.append(f"**期間: {period_start} 〜 {period_end}**")
    lines.append("")

    # 全体サマリー
    lines.append(f"## 📈 全体サマリー")
    lines.append(f"- 提出総数: **{submissions['total_submissions']}件**")
    lines.append(f"- 提出者数: **{submissions['unique_members']}名**")
    lines.append("")

    # 提出率TOP3
    lines.append(f"## 🏆 提出率TOP3")
    medals = ["🥇", "🥈", "🥉"]
    for i, member in enumerate(submissions["top3"]):
        medal = medals[i] if i < len(medals) else "  "
        lines.append(
            f"{medal} **{member['name']}** — "
            f"{member['count']}件提出 / 平均稼働 {member['avg_hours']}h"
        )
    lines.append("")

    # 感情分析
    lines.append(f"## 💬 会員様の声（感情分析）")
    lines.append(
        f"ポジティブ **{sentiment['pos_ratio']}%** ｜ "
        f"ネガティブ **{sentiment['neg_ratio']}%**"
    )
    lines.append("")

    if sentiment["positive_voices"]:
        lines.append("**✨ ポジティブな声:**")
        for v in sentiment["positive_voices"]:
            lines.append(f"> 「{v['voice']}」 — {v['name']}")
        lines.append("")

    if sentiment["negative_voices"]:
        lines.append("**⚠️ 気になる声:**")
        for v in sentiment["negative_voices"]:
            lines.append(f"> 「{v['voice']}」 — {v['name']}")
        lines.append("")

    # 非アクティブ会員
    lines.append(f"## 🔔 非アクティブ会員（提出{MIN_SUBMISSIONS}件以下）")
    if submissions["inactive"]:
        for member in submissions["inactive"]:
            lines.append(f"- **{member['name']}** — {member['count']}件")
        lines.append("")
        lines.append(f"> 📩 月曜20:00にリマインドDMを送信予定")
    else:
        lines.append("なし — 全員アクティブです！ 🎉")
    lines.append("")

    lines.append("---")
    lines.append(f"> 💬 COO SAEPINより：社長、今週の分析レポートです！")

    return "\n".join(lines)


def format_reminder_list(inactive_members):
    """非アクティブ会員へのリマインドDMリストを生成"""
    if not inactive_members:
        return None

    lines = []
    lines.append("# 📩 日報リマインドDM送信レポート")
    lines.append(f"**送信日時: {TODAY.strftime('%Y/%m/%d %H:%M')}**")
    lines.append("")

    for member in inactive_members:
        lines.append(f"## → {member['name']} さん")
        lines.append(f"今週の提出: {member['count']}件")
        lines.append("送信メッセージ:")
        lines.append(f"> {member['name']}さん、お疲れさまです！")
        lines.append(f"> 今週の日報提出が{member['count']}件でした。")
        lines.append("> 小さなことでもOKなので、学びの記録を残していきましょう。")
        lines.append("> 困っていることがあればいつでも相談してくださいね！")
        lines.append("")

    lines.append("---")
    lines.append("> このリストをもとにDiscord DMを送信してください")

    return "\n".join(lines)


# ==========================================
# 送信
# ==========================================
def send_to_discord(message, webhook_url=None):
    """DiscordのWebhookにメッセージを送信"""
    url = webhook_url or DISCORD_WEBHOOK_URL

    # Discordは2000文字制限があるため分割送信
    chunks = []
    current = ""
    for line in message.split("\n"):
        if len(current) + len(line) + 1 > 1900:
            chunks.append(current)
            current = line
        else:
            current += "\n" + line if current else line
    if current:
        chunks.append(current)

    for chunk in chunks:
        payload = {"content": chunk}
        response = requests.post(url, json=payload)
        if response.status_code == 204:
            print("✅ Discord送信成功")
        else:
            print(f"❌ Discord送信失敗: {response.status_code} - {response.text}")


def send_to_slack(message):
    """SlackのWebhookにメッセージを送信"""
    if not SLACK_WEBHOOK_URL:
        print("⏭️ Slack Webhook未設定のためスキップ")
        return

    payload = {"text": message}
    response = requests.post(SLACK_WEBHOOK_URL, json=payload)
    if response.status_code == 200:
        print("✅ Slack送信成功")
    else:
        print(f"❌ Slack送信失敗: {response.status_code} - {response.text}")


def fetch_fb_data(spreadsheet):
    """実践フィードバック会の参加データを前週（月〜日）分取得"""
    ws = spreadsheet.worksheet("DB_FB-Report")
    all_records = ws.get_all_records()

    fb_data = []
    for row in all_records:
        date_str = str(row.get("開催日：", "")).strip()
        if not date_str:
            continue
        try:
            row_date = datetime.strptime(date_str, "%Y/%m/%d")
        except ValueError:
            try:
                row_date = datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                continue

        if LAST_MONDAY <= row_date <= LAST_SUNDAY:
            fb_data.append({
                "name": str(row.get("お名前", "")).strip(),
                "discord": str(row.get("Discordアカウント名", "")).strip(),
                "date": row_date,
            })
    return fb_data


def build_discord_name_map(weekly_data, fb_data, spreadsheet):
    """日報の氏名 → Discordアカウント名の対応表を作成（FB-Reportから取得）"""
    # FB-Reportには氏名とDiscord名の両方がある
    ws = spreadsheet.worksheet("DB_FB-Report")
    all_records = ws.get_all_records()
    name_map = {}
    for row in all_records:
        name = str(row.get("お名前", "")).strip()
        discord = str(row.get("Discordアカウント名", "")).strip()
        if name and discord:
            name_map[name] = discord
    return name_map


def format_award_message(weekly_data, fb_data, name_map):
    """会員表彰メッセージを生成（毎週異なる冒頭メッセージ）"""
    period_start = f"{LAST_MONDAY.month}月{LAST_MONDAY.day}日"
    period_end = f"{LAST_SUNDAY.month}月{LAST_SUNDAY.day}日"

    # 週替わりの冒頭メッセージ（毎週異なる内容。敬意ある温かさで全員に向けた言葉）
    opening_messages = [
        ("みなさま、1週間お疲れさまでした。\n"
         "お忙しい毎日の中で、学びの時間を確保し続けていること——\n"
         "その姿勢そのものが、すでに大きな一歩です。"),
        ("みなさま、今週もありがとうございました。\n"
         "日報に記録された一つひとつの言葉から、\n"
         "着実に前進されている様子が伝わってまいります。"),
        ("みなさま、今週も本当にお疲れさまでした。\n"
         "完璧でなくても、歩みを止めずに続けること。\n"
         "その積み重ねの価値を、私たちは確かに見ております。"),
        ("みなさま、充実した1週間でしたね。\n"
         "日報を書くという小さな習慣が、\n"
         "みなさまの成長を確実に後押ししています。"),
        ("みなさま、1週間お疲れさまでした。\n"
         "「今日も一歩だけ前に進もう」と行動されたその日々が、\n"
         "着実に成果として表れています。"),
        ("みなさま、今週もたくさんの挑戦がありましたね。\n"
         "思い通りにいった日も、そうでなかった日も、\n"
         "そのすべてがみなさまの力になっています。"),
        ("みなさま、1週間ありがとうございました。\n"
         "行動された方だけが見える景色があります。\n"
         "みなさまは今、その景色に立っています。"),
        ("みなさま、今週もお疲れさまでした。\n"
         "「自分はまだまだ…」と感じることがあるかもしれません。\n"
         "でも数字は、確かな前進を示しています。"),
        ("みなさま、素晴らしい1週間でした。\n"
         "小さな行動の積み重ねが、大きな変化を生み出します。\n"
         "その証を、今週もお届けいたします。"),
        ("みなさま、今週もお疲れさまでした。\n"
         "学びを止めずに進み続ける方は、必ず成果にたどり着きます。\n"
         "そんなみなさまの取り組みをご紹介いたします。"),
    ]

    # 週番号をシードにして毎週違うメッセージを選出（同じ週なら同じ結果）
    week_number = LAST_MONDAY.isocalendar()[1]
    year = LAST_MONDAY.year
    rng = random.Random(year * 100 + week_number)
    opening = rng.choice(opening_messages)

    # 日報提出TOP3
    submission_count = defaultdict(int)
    for row in weekly_data:
        name = row["name"]
        if name:
            submission_count[name] += 1
    sorted_report = sorted(submission_count.items(), key=lambda x: x[1], reverse=True)
    report_top3 = sorted_report[:3]

    # 実践FB会参加TOP3
    fb_count = defaultdict(int)
    fb_discord = {}
    for row in fb_data:
        name = row["name"]
        if name:
            fb_count[name] += 1
            fb_discord[name] = row["discord"]
    sorted_fb = sorted(fb_count.items(), key=lambda x: x[1], reverse=True)
    fb_top3 = sorted_fb[:3]

    # 表彰メッセージを組み立て
    lines = []
    lines.append("━━━━━━━━━━━━━━━━━━━━━━━━")
    lines.append(f"**Weekly MVP発表！ （{period_start}〜{period_end}）**")
    lines.append("━━━━━━━━━━━━━━━━━━━━━━━━")
    lines.append("")
    lines.append(opening)
    lines.append("")

    # 日報提出TOP3
    lines.append("**【日報提出回数 TOP3】**")
    lines.append("毎日の記録が、着実な成長につながっています。")
    lines.append("")
    medals = ["🥇", "🥈", "🥉"]
    for i, (name, count) in enumerate(report_top3):
        discord_name = name_map.get(name, name)
        lines.append(f"{medals[i]} **{discord_name}** さん — {count}回提出")
    lines.append("")

    # 実践FB会TOP3
    if fb_top3:
        lines.append("**【実践フィードバック会 参加回数 TOP3】**")
        lines.append("学びを実践に変える行動力、見事です。")
        lines.append("")
        for i, (name, count) in enumerate(fb_top3):
            discord_name = fb_discord.get(name, name_map.get(name, name))
            lines.append(f"{medals[i]} **{discord_name}** さん — {count}回参加")
        lines.append("")
    else:
        lines.append("（今週は実践フィードバック会の開催がありませんでした）")
        lines.append("")

    lines.append("━━━━━━━━━━━━━━━━━━━━━━━━")
    lines.append("")
    lines.append("日々の積み重ねが、")
    lines.append("半年後のみなさまを大きく変えていきます。")
    lines.append("")
    lines.append("来週もみなさまの挑戦を、心より応援しております。")
    lines.append("")
    lines.append("**TOP3に選ばれたみなさま、おめでとうございます。**")
    lines.append("ぜひリアクションで、お互いの頑張りを称え合いましょう！ 🎉👏🔥")

    return "\n".join(lines)


def save_report(report, report_type="weekly"):
    """レポートをファイルに保存（タイトル: 会員分析（日報）●月●日-●月●日）"""
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    start_str = f"{LAST_MONDAY.month}月{LAST_MONDAY.day}日"
    end_str = f"{LAST_SUNDAY.month}月{LAST_SUNDAY.day}日"
    filename = f"会員分析（日報）{start_str}-{end_str}.md"
    filepath = REPORTS_DIR / filename
    filepath.write_text(report, encoding="utf-8")
    print(f"💾 レポート保存: {filepath}")
    return filepath


# ==========================================
# メイン処理
# ==========================================
def main(mode="report"):
    """
    mode:
      "report"   — 週次分析レポート配信（月曜8:00）
      "remind"   — 非アクティブ会員リマインドリスト（月曜20:00）
      "award"    — 週間MVP表彰投稿（月曜19:00）
      "slack"    — Slack週次報告（火曜）
      "dry-run"  — 送信せずにレポートを表示のみ
    """
    print(f"📊 会員日報ウィークリー分析 開始... ({TODAY_STR})")
    print(f"  モード: {mode}")
    print("")

    # Google Sheets接続
    print("🔗 Google Sheets接続中...")
    if mode == "award":
        spreadsheet = connect_spreadsheet()
        sheet = spreadsheet.worksheet(SHEET_NAME)
    else:
        sheet = connect_sheets()
        spreadsheet = None

    # データ取得
    print("📥 直近1週間のデータ取得中...")
    weekly_data = fetch_weekly_data(sheet)
    print(f"  → {len(weekly_data)}件の日報を取得")

    if not weekly_data:
        print("⚠️ 直近1週間の日報データがありません")
        return

    # 分析実行
    print("🔍 分析中...")
    submissions = analyze_submissions(weekly_data)
    sentiment = analyze_sentiment(weekly_data)

    print(f"  → 提出者: {submissions['unique_members']}名")
    print(f"  → 総提出数: {submissions['total_submissions']}件")
    print(f"  → 感情: ポジ{sentiment['pos_ratio']}% / ネガ{sentiment['neg_ratio']}%")
    print(f"  → 非アクティブ: {len(submissions['inactive'])}名")
    print("")

    # モード別処理
    if mode == "report" or mode == "dry-run":
        report = format_weekly_report(submissions, sentiment, weekly_data)
        save_report(report, "weekly")
        print("--- 週次レポート ---")
        print(report)
        print("--- ここまで ---\n")

        if mode == "report":
            send_to_discord(report)

    elif mode == "remind":
        reminder = format_reminder_list(submissions["inactive"])
        if reminder:
            save_report(reminder, "remind")
            print("--- リマインドリスト ---")
            print(reminder)
            print("--- ここまで ---\n")
            send_to_discord(reminder)
        else:
            print("✨ 非アクティブ会員なし — リマインド不要です")

    elif mode == "award":
        # 実践FB会データ取得
        print("📥 実践FB会データ取得中...")
        fb_data = fetch_fb_data(spreadsheet)
        print(f"  → {len(fb_data)}件のFB会参加を取得")

        # Discord名の対応表を作成
        name_map = build_discord_name_map(weekly_data, fb_data, spreadsheet)

        # 表彰メッセージ生成
        award_msg = format_award_message(weekly_data, fb_data, name_map)
        print("--- 表彰メッセージ ---")
        print(award_msg)
        print("--- ここまで ---\n")
        send_to_discord(award_msg)

    elif mode == "slack":
        report = format_weekly_report(submissions, sentiment, weekly_data)
        save_report(report, "slack")
        send_to_slack(report)


if __name__ == "__main__":
    # コマンドライン引数でモード切替
    # python 週次日報分析_discord.py          → dry-run（デフォルト）
    # python 週次日報分析_discord.py report   → Discord配信
    # python 週次日報分析_discord.py remind   → リマインドリスト配信
    # python 週次日報分析_discord.py award    → 週間MVP表彰（Discord）
    # python 週次日報分析_discord.py slack    → Slack配信
    mode = sys.argv[1] if len(sys.argv) > 1 else "dry-run"
    main(mode)
