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
REPORTS_DIR = PROJECT_ROOT / "03_clients" / "SIFTAI" / "プロプレミアムTEAM" / "会員管理" / "週次レポート"

# Google Sheets
SPREADSHEET_ID = "1EHKgmE1d7T5N9GbT_QJV72Q8Dh91iNysHHkz6LEWJy4"
SHEET_NAME = "DB_Daily-Report"
PRO_SHEET_NAME = "DB_Pro-Members"
PREMIUM_SHEET_NAME = "DB_Premium-Members"
WEEKLY_SHEET_NAME = "DB_Weekly-Report"
CREDENTIALS_PATH = CONFIG_DIR / "google_service_account.json"

# 列名マッピング（スプレッドシートのヘッダーに合わせる）
COL_TIMESTAMP = "タイムスタンプ"
COL_NAME = "氏名"
COL_DATE = "日付"
COL_HOURS = "稼働時間"
COL_LEARNING = "今日の学び"
COL_TOMORROW = "明日やること"
COL_MONETIZE = "今週のマネタイズ成果"
COL_MONETIZE_PHASE = "現在のマネタイズ進捗フェーズ"

# Discord Webhook（週次レポート配信用）
DISCORD_WEBHOOK_URL = "https://discord.com/api/webhooks/1480602170368856275/DCDEaDwYhIIH0xGioPjanNbTocElv0FxtbnGo7HTqSDfpR_bn_IPO6Rm2PNhC7p0gUB4"

# Slack Webhook（火曜の週次報告用 — 設定後に有効化）
SLACK_WEBHOOK_URL = ""

# 非アクティブ判定の閾値
MIN_SUBMISSIONS = 3  # 週間提出がこの件数以下で非アクティブ判定

# 運営メンバー（データ集計から除外）
EXCLUDED_MEMBERS = ["安部友博"]

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
# --week-ago N で N週前を対象にできる（デフォルト: 1 = 先週）
_week_ago = 1
for i, arg in enumerate(sys.argv):
    if arg == "--week-ago" and i + 1 < len(sys.argv):
        _week_ago = int(sys.argv[i + 1])
_current_weekday = TODAY.weekday()
THIS_MONDAY = TODAY - timedelta(days=_current_weekday)
LAST_MONDAY = THIS_MONDAY - timedelta(days=7 * _week_ago)
LAST_SUNDAY = LAST_MONDAY + timedelta(days=6)

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
            name = str(row.get(COL_NAME, "")).strip()
            if name in EXCLUDED_MEMBERS:
                continue
            weekly_data.append({
                "name": name,
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
def analyze_categories(weekly_data):
    """日報テキストからカテゴリー分類（行動面・心理的・スキル・時間環境・前進実感）"""
    categories = {
        "行動面の課題": {"keywords": ["応募", "行動", "手が止", "後回し", "進められ", "できなかった", "やれなかった"], "members": set()},
        "心理的不安": {"keywords": ["不安", "怖い", "躊躇", "自信がない", "迷い", "恐怖", "落ち込"], "members": set()},
        "スキル不足": {"keywords": ["わからな", "難し", "できない", "使い方", "操作", "プロンプト", "理解でき"], "members": set()},
        "時間・環境要因": {"keywords": ["時間がな", "忙し", "体調", "仕事が", "時短", "間に合", "疲れ"], "members": set()},
        "前進実感・成功体験": {"keywords": ["できた", "達成", "成功", "受注", "納品", "マネタイズ", "初めて", "完了", "0→1"], "members": set()},
    }

    for row in weekly_data:
        text = row["learning"] + " " + row.get("tomorrow", "")
        name = row["name"]
        if not text.strip() or not name:
            continue
        for cat_name, cat_info in categories.items():
            if any(kw in text for kw in cat_info["keywords"]):
                cat_info["members"].add(name)

    total_members = len(set(r["name"] for r in weekly_data if r["name"]))
    result = {}
    for cat_name, cat_info in categories.items():
        members = list(cat_info["members"])
        pct = round(len(members) / total_members * 100) if total_members > 0 else 0
        result[cat_name] = {"count": len(members), "pct": pct, "members": members}
    return result


def analyze_monetize_phase(weekly_data):
    """マネタイズフェーズ分類"""
    phases = {
        "フェーズ1（操作不安・全体像）": {
            "keywords": ["使い方", "操作", "わからな", "初心者", "基礎", "全体像"],
            "members": set()
        },
        "フェーズ2（優先順位の迷い）": {
            "keywords": ["何から", "優先", "迷い", "どれを", "選べない"],
            "members": set()
        },
        "フェーズ3（応募・継続の壁）": {
            "keywords": ["応募", "案件", "提案", "営業", "継続", "クライアント"],
            "members": set()
        },
        "フェーズ4（質の追求・伸び悩み）": {
            "keywords": ["クオリティ", "品質", "伸び悩", "単価", "ポートフォリオ", "差別化"],
            "members": set()
        },
    }

    for row in weekly_data:
        text = row["learning"] + " " + row.get("tomorrow", "")
        name = row["name"]
        if not text.strip() or not name:
            continue
        for phase_name, phase_info in phases.items():
            if any(kw in text for kw in phase_info["keywords"]):
                phase_info["members"].add(name)

    result = {}
    for phase_name, phase_info in phases.items():
        result[phase_name] = {"count": len(phase_info["members"]), "members": list(phase_info["members"])}
    return result


def parse_yen_amount(text):
    """テキストから日本円の金額を抽出して合計を返す（円単位）

    対応フォーマット:
      300円, 5万円, 10,000円, ¥5000, 1.5万円, 50万, ５万円（全角数字）
    """
    if not text:
        return 0

    total = 0

    # 全角数字を半角に変換
    zen = "０１２３４５６７８９"
    han = "0123456789"
    for z, h in zip(zen, han):
        text = text.replace(z, h)

    # パターン1: ●万円 / ●万（小数対応）
    for m in re.finditer(r'([\d,.]+)\s*万\s*円?', text):
        try:
            val = float(m.group(1).replace(",", ""))
            total += int(val * 10000)
        except ValueError:
            pass

    # パターン2: ●円（万円でないもの）
    for m in re.finditer(r'([\d,]+)\s*円', text):
        # 万円の一部でないかチェック
        start = m.start()
        if start > 0 and text[start-1] == '万':
            continue
        try:
            total += int(m.group(1).replace(",", ""))
        except ValueError:
            pass

    # パターン3: ¥●●●
    for m in re.finditer(r'[¥￥]([\d,]+)', text):
        try:
            total += int(m.group(1).replace(",", ""))
        except ValueError:
            pass

    return total


def fetch_monetize_data(spreadsheet):
    """週報からマネタイズ成果データを前週（月〜日）分取得"""
    ws = spreadsheet.worksheet(WEEKLY_SHEET_NAME)
    all_records = ws.get_all_records()

    data = []
    for row in all_records:
        date_str = str(row.get(COL_DATE, "")).strip()
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
            name = str(row.get(COL_NAME, "")).strip()
            monetize_text = str(row.get(COL_MONETIZE, "")).strip()
            phase = str(row.get(COL_MONETIZE_PHASE, "")).strip()
            amount = parse_yen_amount(monetize_text)

            if name and name != "テスト" and name not in EXCLUDED_MEMBERS:
                data.append({
                    "name": name,
                    "amount": amount,
                    "text": monetize_text,
                    "phase": phase,
                    "date": row_date,
                })
    return data


def analyze_monetize_ranking(monetize_data):
    """マネタイズ成果ランキングを分析"""
    # 会員ごとに金額を合算（同じ週に複数週報がある場合）
    member_total = defaultdict(int)
    member_text = {}
    for row in monetize_data:
        member_total[row["name"]] += row["amount"]
        if row["text"]:
            member_text[row["name"]] = row["text"]

    # 金額ありの会員だけ抽出してランキング
    ranked = sorted(
        [(name, amt) for name, amt in member_total.items() if amt > 0],
        key=lambda x: x[1],
        reverse=True,
    )

    # 統計
    amounts = [amt for _, amt in ranked]
    if amounts:
        avg_amount = sum(amounts) / len(amounts)
        max_entry = ranked[0]   # (name, amount)
        min_entry = ranked[-1]  # (name, amount)
    else:
        avg_amount = 0
        max_entry = None
        min_entry = None

    return {
        "ranking": ranked[:10],  # TOP10
        "total_members": len(member_total),
        "earners": len(ranked),
        "total_amount": sum(amounts),
        "avg_amount": avg_amount,
        "max_entry": max_entry,
        "min_entry": min_entry,
    }


def format_yen(amount):
    """金額を見やすくフォーマット（万円単位も使用）"""
    if amount >= 10000:
        man = amount / 10000
        if man == int(man):
            return f"{int(man)}万円"
        return f"{man:.1f}万円"
    return f"{amount:,}円"


def generate_recommendations(categories, phases, sentiment, submissions):
    """運営への提言を自動生成"""
    recs = []

    # 個別フォロー推奨（心理的不安 or スキル不足の会員）
    follow_up = set()
    if categories["心理的不安"]["members"]:
        follow_up.update(categories["心理的不安"]["members"])
    if categories["スキル不足"]["members"]:
        follow_up.update(categories["スキル不足"]["members"])

    if follow_up:
        recs.append("個別フォロー推奨：")
        for name in list(follow_up)[:5]:
            reasons = []
            if name in categories["心理的不安"]["members"]:
                reasons.append("心理的な壁あり")
            if name in categories["スキル不足"]["members"]:
                reasons.append("スキル面でのつまずき")
            recs.append(f"  {name}様：{'・'.join(reasons)}。個別声かけや1on1でのフォローが有効です。")

    # FB会・サービス改善案
    recs.append("")
    recs.append("FB会・サービス改善案：")
    if categories["スキル不足"]["count"] >= 3:
        recs.append(f"  スキル面の課題を抱える会員が{categories['スキル不足']['count']}名。"
                    "FB会でのハンズオン形式や、プロンプト配布の徹底が有効です。")
    if phases["フェーズ1（操作不安・全体像）"]["count"] >= 3:
        recs.append(f"  フェーズ1（操作不安）の会員が{phases['フェーズ1（操作不安・全体像）']['count']}名。"
                    "初心者向けオンボーディング資料の充実を推奨します。")
    if categories["行動面の課題"]["count"] >= 3:
        recs.append(f"  行動面の課題を抱える会員が{categories['行動面の課題']['count']}名。"
                    "応募文作成会やポートフォリオ作成会など、短時間で形にする集中ワークショップの開催が有効です。")
    if sentiment["neg_ratio"] > 30:
        recs.append(f"  マイナス傾向率が{sentiment['neg_ratio']}％と高めです。"
                    "成功体験の共有会や、モチベーション施策の強化を検討してください。")

    return recs


def normalize_to_pro(name, pro_names):
    """日報の氏名をPROメンバー名簿の名前に正規化"""
    name = name.strip()
    if name in pro_names:
        return name
    for pn in pro_names:
        if pn in name or name in pn:
            return pn
    return name


def analyze_pro_inactive(spreadsheet, weekly_data, fb_data):
    """非アクティブPRO会員を過去4週の傾向付きで分析"""
    # PROメンバー名簿を取得
    pro_ws = spreadsheet.worksheet(PRO_SHEET_NAME)
    pro_records = pro_ws.get_all_records()
    pro_names = set(
        str(r.get("お名前", "")).strip()
        for r in pro_records
        if str(r.get("お名前", "")).strip()
    )

    # 全日報データ取得（傾向分析用）
    daily_ws = spreadsheet.worksheet(SHEET_NAME)
    all_daily = daily_ws.get_all_records()

    # 全FB会データ取得
    fb_ws = spreadsheet.worksheet("DB_FB-Report")
    all_fb = fb_ws.get_all_records()

    def parse_date_to_date(s):
        s = str(s).strip()
        for fmt in ["%Y/%m/%d", "%Y-%m-%d"]:
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                pass
        return None

    def get_week_monday(d):
        return d - timedelta(days=d.weekday())

    # 週基準（dateオブジェクト）
    today = datetime.now().date()
    this_monday = today - timedelta(days=today.weekday())
    last_monday = this_monday - timedelta(days=7)
    last_sunday = last_monday + timedelta(days=6)

    # 過去4週のmonday
    weeks = [this_monday - timedelta(days=7 * i) for i in range(4, 0, -1)]

    # 日報の週別集計
    daily_by_week = defaultdict(lambda: defaultdict(int))
    for row in all_daily:
        name = normalize_to_pro(str(row.get(COL_NAME, "")), pro_names)
        dt = parse_date_to_date(row.get(COL_DATE, ""))
        if name and dt and name in pro_names:
            monday = get_week_monday(dt)
            daily_by_week[name][monday] += 1

    # FB会の週別集計
    fb_by_week = defaultdict(lambda: defaultdict(int))
    for row in all_fb:
        name = normalize_to_pro(str(row.get("お名前", "")), pro_names)
        dt = parse_date_to_date(row.get("開催日：", ""))
        if name and dt and name in pro_names:
            monday = get_week_monday(dt)
            fb_by_week[name][monday] += 1

    # 非アクティブ判定（直近週で日報3件以下）
    members = []
    zero_4weeks = []
    declining = []

    for name in sorted(pro_names):
        rc = daily_by_week[name].get(last_monday, 0)
        fc = fb_by_week[name].get(last_monday, 0)
        if rc <= 3:
            rp = round(rc / 7 * 100)
            fp = round(fc / 7 * 100)
            trends = []
            week_totals = []
            for w in weeks:
                wd = daily_by_week[name].get(w, 0)
                wf = fb_by_week[name].get(w, 0)
                trends.append(f"{wd}/{wf}")
                week_totals.append(wd + wf)

            members.append({
                "name": name,
                "report": rc,
                "report_pct": rp,
                "fb": fc,
                "fb_pct": fp,
                "trends": trends,
            })

            # 4週連続ゼロ
            if all(t == 0 for t in week_totals):
                zero_4weeks.append(name)

            # 急降下（前半2週の合計 > 後半2週の合計 の50%以上減）
            first_half = sum(week_totals[:2])
            second_half = sum(week_totals[2:])
            if first_half >= 4 and second_half <= first_half * 0.5:
                declining.append(name)

    week_labels = [{"label": f"{w.month}/{w.day}週"} for w in weeks]

    return {
        "members": members,
        "weeks": week_labels,
        "inactive_count": len(members),
        "total_count": len(pro_names),
        "zero_4weeks": zero_4weeks,
        "declining": declining,
    }


def format_weekly_report(submissions, sentiment, weekly_data, fb_data=None, name_map=None, spreadsheet=None):
    """週次報告書を生成（成功習慣行動率100％フォーマット準拠）"""
    period_start = LAST_MONDAY.strftime("%Y年%m月%d日")
    period_end = LAST_SUNDAY.strftime("%m月%d日")

    # 追加分析
    categories = analyze_categories(weekly_data)
    phases = analyze_monetize_phase(weekly_data)

    # 稼働時間集計
    member_hours = defaultdict(float)
    member_count = defaultdict(int)
    for row in weekly_data:
        name = row["name"]
        if name:
            h = parse_hours(row["hours"])
            member_hours[name] += h
            member_count[name] += 1
    total_hours = sum(member_hours.values())
    avg_hours = round(total_hours / len(member_hours), 1) if member_hours else 0
    hours_top3 = sorted(member_hours.items(), key=lambda x: x[1], reverse=True)[:3]

    lines = []
    lines.append("週次報告書　")
    lines.append("成功習慣行動率100％")
    lines.append("")

    # ① 基本統計データ
    lines.append("① 基本統計データ（全体・稼働状況）")
    lines.append(f"集計期間：{period_start}〜{period_end}")
    lines.append(f"総稼働時間：{round(total_hours, 1)}時間（日報データより集計）")
    lines.append(f"平均稼働時間：{avg_hours}時間（1人あたり平均）")
    lines.append("稼働時間TOP3：")
    for i, (name, hours) in enumerate(hours_top3, 1):
        avg = round(hours / member_count[name], 1) if member_count[name] > 0 else 0
        lines.append(f"  {i}位：{name}（総稼働時間：{round(hours, 1)}h / 平均：{avg}h）")
    lines.append("日報提出数TOP3：")
    top3_str = "、".join(
        f"{i}位：{m['name']}（{m['count']}回）"
        for i, m in enumerate(submissions["top3"], 1)
    )
    lines.append(f"  {top3_str}")
    # 提出者統計（ダッシュボードHTML用）
    report_2plus = sum(1 for c in submissions["all_members"].values() if c >= 2)
    lines.append(f"日報提出者：{submissions['unique_members']}名（提出{submissions['total_submissions']}件）")
    lines.append(f"日報2回以上提出者：{report_2plus}名（提出率{round(report_2plus / 38 * 100)}%）")
    lines.append("")

    # ② 実践フィードバック会分析
    lines.append("② 実践フィードバック会（FB会）分析")
    if fb_data:
        # 曜日別集計
        weekday_names = ["月", "火", "水", "木", "金", "土", "日"]
        weekday_count = defaultdict(int)
        fb_member_count = defaultdict(int)
        for row in fb_data:
            wd = row["date"].weekday()
            weekday_count[wd] += 1
            fb_member_count[row["name"]] += 1

        wd_str = "、".join(f"{weekday_names[i]}：{weekday_count.get(i, 0)}人" for i in range(7))
        lines.append(f"曜日別参加者数：{wd_str}")

        fb_top3 = sorted(fb_member_count.items(), key=lambda x: x[1], reverse=True)[:3]
        lines.append("参加回数TOP3：")
        fb_top3_str = "、".join(f"{i}位：{name}（{count}回）" for i, (name, count) in enumerate(fb_top3, 1))
        lines.append(f"  {fb_top3_str}")
    else:
        lines.append("（今週はFB会データなし）")
    lines.append("")

    # ③ 感情・傾向占有率
    lines.append("③ 感情・傾向占有率（全体）")
    lines.append(f"プラス傾向率：{sentiment['pos_ratio']}％（前進、自信、成功体験、学びの習得）")
    lines.append(f"マイナス傾向率：{sentiment['neg_ratio']}％（不安、停滞、環境課題、操作への迷い）")

    # 分析コメント生成
    if sentiment['pos_ratio'] >= 60:
        comment = "全体的にポジティブな反応が多く、コミュニティの健康状態は良好です。"
    elif sentiment['pos_ratio'] >= 40:
        comment = "プラス傾向とマイナス傾向が拮抗しています。停滞会員へのフォローを強化すると全体の活力が上がります。"
    else:
        comment = "マイナス傾向が目立ちます。個別フォローと成功体験の共有を優先的に行うことを推奨します。"
    lines.append(f"【分析コメント】：{comment}")
    lines.append("")

    # ④ カテゴリー分析
    lines.append("④ カテゴリー分析（意味分類 ＋ 該当会員名）")
    for cat_name, cat_data in categories.items():
        if cat_data["members"]:
            members_str = "、".join(cat_data["members"][:5])
            if len(cat_data["members"]) > 5:
                members_str += "..."
            lines.append(f"{cat_name}：{cat_data['count']}人（{cat_data['pct']}％）"
                        f"［該当者：{members_str}］")
        else:
            lines.append(f"{cat_name}：0人（0％）")
    lines.append("")

    # ⑤ マネタイズフェーズ × 課題分析
    lines.append("⑤ マネタイズフェーズ × 課題分析（該当会員名）")
    for phase_name, phase_data in phases.items():
        if phase_data["members"]:
            members_str = "、".join(phase_data["members"][:5])
            lines.append(f"{phase_name}：{phase_data['count']}人［該当：{members_str}］")
        else:
            lines.append(f"{phase_name}：0人")
    lines.append("")

    # ⑥ マネタイズ成果ランキング
    if spreadsheet:
        monetize_data = fetch_monetize_data(spreadsheet)
        monetize_ranking = analyze_monetize_ranking(monetize_data)

        lines.append("⑥ マネタイズ成果ランキング")

        if monetize_ranking["earners"] > 0:
            # サマリー統計
            weekly_sub_count = monetize_ranking['total_members']
            weekly_rate_pct = round(weekly_sub_count / 38 * 100)
            lines.append(f"週報提出者: {weekly_sub_count}名 / 収益報告者: {monetize_ranking['earners']}名")
            lines.append(f"週報提出率: {weekly_rate_pct}%（{weekly_sub_count}/38名）")
            lines.append(f"合計マネタイズ額: {format_yen(monetize_ranking['total_amount'])}")
            lines.append("")

            # 最上位・平均・最下位
            max_e = monetize_ranking["max_entry"]
            min_e = monetize_ranking["min_entry"]
            lines.append(f"  最上位: {max_e[0]}（{format_yen(max_e[1])}）")
            lines.append(f"  平均: {format_yen(int(monetize_ranking['avg_amount']))}")
            if min_e and min_e[0] != max_e[0]:
                lines.append(f"  最下位: {min_e[0]}（{format_yen(min_e[1])}）")
            lines.append("")

            # TOP10ランキング
            lines.append("マネタイズTOP10：")
            medals = ["🥇", "🥈", "🥉"]
            for i, (name, amount) in enumerate(monetize_ranking["ranking"]):
                mark = medals[i] if i < 3 else f"  {i+1}位"
                lines.append(f"  {mark} {name} — {format_yen(amount)}")
        else:
            lines.append("（今週は収益報告なし）")
    else:
        lines.append("⑥ マネタイズ成果ランキング")
        lines.append("（週報データ未取得）")
    lines.append("")

    # ⑦ 重要インサイト：会員様の生の声
    lines.append("⑦ 重要インサイト：会員様の生の声")
    if sentiment["positive_voices"]:
        lines.append("前進を支える「プラス傾向」の声")
        for v in sentiment["positive_voices"]:
            lines.append(f"  {v['name']}：「{v['voice']}」")
    if sentiment["negative_voices"]:
        lines.append("停滞を招く「マイナス傾向」の声")
        for v in sentiment["negative_voices"]:
            lines.append(f"  {v['name']}：「{v['voice']}」")
    lines.append("")

    # ⑧ 運営への提言
    lines.append("⑧ 運営への提言")
    recs = generate_recommendations(categories, phases, sentiment, submissions)
    for rec in recs:
        lines.append(rec)
    lines.append("")

    # ⑨ 非アクティブPRO会員（過去4週傾向付き）
    if spreadsheet:
        lines.append("⑨ アクションすべき非アクティブPRO会員（過去4週傾向）")
        pro_inactive = analyze_pro_inactive(spreadsheet, weekly_data, fb_data)
        if pro_inactive["members"]:
            lines.append(f"非アクティブ: {pro_inactive['inactive_count']}名 / PRO全体: {pro_inactive['total_count']}名")
            lines.append("")

            # 表ヘッダー
            weeks_header = " | ".join(f"{w['label']}" for w in pro_inactive["weeks"])
            lines.append(f"| 会員名 | 今週日報 | 提出率 | 今週FB | 参加率 | {weeks_header} |")

            for m in pro_inactive["members"]:
                trend = " | ".join(m["trends"])
                lines.append(f"| {m['name']} | {m['report']}/7 | {m['report_pct']}% "
                            f"| {m['fb']}/7 | {m['fb_pct']}% | {trend} |")

            # 特記事項
            lines.append("")
            if pro_inactive["zero_4weeks"]:
                names = "、".join(pro_inactive["zero_4weeks"][:8])
                lines.append(f"⚠️ 完全未稼働（4週連続ゼロ）: {names}")
            if pro_inactive["declining"]:
                names = "、".join(pro_inactive["declining"][:5])
                lines.append(f"📉 急降下中: {names}")
        else:
            lines.append("全PRO会員がアクティブです")

    return "\n".join(lines)


def analyze_dm_targets(spreadsheet):
    """2週連続で日報提出率80%以下のPRO会員を抽出（Discord名付き）

    ルール定義:
      - 対象: PROサービス対象者（DB_Pro-Members）
      - 条件: 直近2週間（先々週＋先週）ともに日報提出率が80%以下（7日中5日以下）
      - Discord名: DB_FB-Reportの「Discordアカウント名」列から取得
      - DM担当:
        - PRO only → 秘書チーム（門田遥・浅野・川瀬）※仲山さんは除く
        - PRO兼PREMIUM → コンシェルジュ（山本大輔・安部友博）
    """
    # PROメンバー名簿
    pro_ws = spreadsheet.worksheet(PRO_SHEET_NAME)
    pro_records = pro_ws.get_all_records()
    pro_names = set(
        str(r.get("お名前", "")).strip()
        for r in pro_records
        if str(r.get("お名前", "")).strip()
    )

    # PREMIUMメンバー名簿
    premium_ws = spreadsheet.worksheet(PREMIUM_SHEET_NAME)
    premium_records = premium_ws.get_all_records()
    premium_names = set(
        str(r.get("お名前", "")).strip()
        for r in premium_records
        if str(r.get("お名前", "")).strip()
    )

    # 全日報データ
    daily_ws = spreadsheet.worksheet(SHEET_NAME)
    all_daily = daily_ws.get_all_records()

    # Discord名マッピング
    fb_ws = spreadsheet.worksheet("DB_FB-Report")
    all_fb = fb_ws.get_all_records()
    discord_map = {}
    for row in all_fb:
        name = str(row.get("お名前", "")).strip()
        discord = str(row.get("Discordアカウント名", "")).strip()
        if name and discord:
            discord_map[name] = discord

    def parse_date(s):
        s = str(s).strip()
        for fmt in ["%Y/%m/%d", "%Y-%m-%d"]:
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                pass
        return None

    def get_week_monday(d):
        return d - timedelta(days=d.weekday())

    today = datetime.now().date()
    this_monday = today - timedelta(days=today.weekday())
    last_monday = this_monday - timedelta(days=7)      # 先週月曜
    prev_monday = this_monday - timedelta(days=14)      # 先々週月曜
    last_sunday = last_monday + timedelta(days=6)
    prev_sunday = prev_monday + timedelta(days=6)

    # 週別日報提出数
    daily_by_week = defaultdict(lambda: defaultdict(int))
    for row in all_daily:
        name = normalize_to_pro(str(row.get(COL_NAME, "")), pro_names)
        dt = parse_date(row.get(COL_DATE, ""))
        if name and dt and name in pro_names:
            monday = get_week_monday(dt)
            daily_by_week[name][monday] += 1

    # 2週連続80%以下（7日中5日以下）を抽出
    threshold = 0.8  # 80%
    max_days = 7
    min_submissions = int(max_days * threshold)  # 5.6 → 5

    # DM担当者
    secretary_senders = ["門田遥", "浅野", "川瀬"]       # PRO only
    concierge_senders = ["山本大輔", "安部友博"]          # PRO兼PREMIUM
    sec_idx = 0
    con_idx = 0

    targets = []
    for name in sorted(pro_names):
        week1_count = daily_by_week[name].get(prev_monday, 0)
        week2_count = daily_by_week[name].get(last_monday, 0)
        week1_pct = round(week1_count / max_days * 100)
        week2_pct = round(week2_count / max_days * 100)

        # 2週連続で80%以下
        if week1_count <= min_submissions and week2_count <= min_submissions:
            discord_name = discord_map.get(name, "")
            # 部分一致もチェック
            if not discord_name:
                for dn_name, dn_val in discord_map.items():
                    if dn_name in name or name in dn_name:
                        discord_name = dn_val
                        break

            # PREMIUM判定 → 担当者振り分け
            is_premium = name in premium_names
            if not is_premium:
                # 部分一致でもチェック
                for pn in premium_names:
                    if pn in name or name in pn:
                        is_premium = True
                        break

            if is_premium:
                sender = concierge_senders[con_idx % len(concierge_senders)]
                con_idx += 1
                plan = "PRO+PREMIUM"
            else:
                sender = secretary_senders[sec_idx % len(secretary_senders)]
                sec_idx += 1
                plan = "PRO"

            targets.append({
                "name": name,
                "discord": discord_name,
                "week1_count": week1_count,
                "week1_pct": week1_pct,
                "week2_count": week2_count,
                "week2_pct": week2_pct,
                "plan": plan,
                "sender": sender,
            })

    return {
        "targets": targets,
        "total_pro": len(pro_names),
        "premium_count": sum(1 for t in targets if t["plan"] == "PRO+PREMIUM"),
        "prev_period": f"{prev_monday.month}/{prev_monday.day}〜{prev_sunday.month}/{prev_sunday.day}",
        "last_period": f"{last_monday.month}/{last_monday.day}〜{last_sunday.month}/{last_sunday.day}",
    }


def format_reminder_list(dm_data):
    """2週連続80%以下の会員へのDMリストを生成"""
    if not dm_data or not dm_data["targets"]:
        return None

    lines = []
    lines.append("# 📩 日報リマインドDM送信リスト")
    lines.append(f"**作成日時: {TODAY.strftime('%Y/%m/%d %H:%M')}**")
    lines.append("")
    lines.append("## ルール")
    lines.append("- 対象: 2週連続で日報提出率80%以下のPRO会員")
    lines.append(f"- 先々週: {dm_data['prev_period']}")
    lines.append(f"- 先週: {dm_data['last_period']}")
    lines.append(f"- 該当者: {len(dm_data['targets'])}名 / PRO全体: {dm_data['total_pro']}名")
    lines.append(f"  - うちPRO+PREMIUM: {dm_data['premium_count']}名")
    lines.append("- DM担当:")
    lines.append("  - PRO only → 秘書チーム（門田遥 / 浅野 / 川瀬）")
    lines.append("  - PRO+PREMIUM → コンシェルジュ（山本大輔 / 安部友博）")
    lines.append("")

    # テーブル形式で一覧
    lines.append("| # | 会員名 | Discord名 | プラン | 先々週 | 先週 | DM担当 |")
    lines.append("|---|--------|-----------|--------|--------|------|--------|")
    for i, t in enumerate(dm_data["targets"]):
        discord_display = t["discord"] if t["discord"] else "（未登録）"
        lines.append(
            f"| {i+1} | {t['name']} | {discord_display} "
            f"| {t['plan']} "
            f"| {t['week1_count']}/7（{t['week1_pct']}%） "
            f"| {t['week2_count']}/7（{t['week2_pct']}%） "
            f"| {t['sender']} |"
        )
    lines.append("")

    # 送信メッセージテンプレート
    lines.append("## 送信メッセージ（テンプレート）")
    lines.append("```")
    lines.append("[名前]さん、お疲れさまです。")
    lines.append("最近、日報の提出が少なくなっているようです。")
    lines.append("小さなことでも構いませんので、学びの記録を残していきましょう。")
    lines.append("困っていることや、何か気になることがあれば")
    lines.append("いつでもお声がけください。")
    lines.append("```")

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
            name = str(row.get("お名前", "")).strip()
            if name in EXCLUDED_MEMBERS:
                continue
            fb_data.append({
                "name": name,
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


def _esc(text):
    """HTMLエスケープ"""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _build_week_body(md_text, week_monday, prev_md_text=None):
    """1週分のダッシュボードbody HTMLを生成（タブコンテンツ用）"""
    week_sunday = week_monday + timedelta(days=6)
    period_start_str = f"{week_monday.month}月{week_monday.day}日"
    period_end_str = f"{week_sunday.month}月{week_sunday.day}日"
    weekday_names = ['月','火','水','木','金','土','日']
    period_full = f"{week_monday.strftime('%Y年%m月%d日')}（{weekday_names[week_monday.weekday()]}）〜 {week_sunday.strftime('%m月%d日')}（{weekday_names[week_sunday.weekday()]}）"

    d = _parse_md_report(md_text)
    prev = _parse_md_report(prev_md_text) if prev_md_text else None

    # --- 前週との比較 ---
    def _delta(cur, prv, suffix="", fmt_func=None):
        if prv is None:
            return ""
        diff = cur - prv
        sign = "+" if diff > 0 else ""
        val = fmt_func(diff) if fmt_func else f"{sign}{diff}"
        if not fmt_func:
            val = f"{sign}{diff}{suffix}"
        return val

    # --- 日報2回以上・週報データ（Google Sheetsから直接取得） ---
    report_2plus_this = d.get("report_2plus", "—")
    report_2plus_prev = prev.get("report_2plus", "—") if prev else "—"
    report_rate_this = d.get("report_rate", "—")
    report_rate_prev = prev.get("report_rate", "—") if prev else "—"
    weekly_submitters_this = d.get("weekly_submitters", "—")
    weekly_submitters_prev = prev.get("weekly_submitters", "—") if prev else "—"
    weekly_rate_this = d.get("weekly_rate", "—")
    weekly_rate_prev = prev.get("weekly_rate", "—") if prev else "—"

    # --- 各セクションHTML ---
    inactive_table_html = _parse_inactive_table(md_text)
    alerts_html = _parse_alerts(md_text)
    voices_html = _parse_voices(md_text)
    proposals_html = _parse_proposals(md_text)
    categories_html = _parse_categories_bars(md_text)
    phases_html = _parse_phase_cards(md_text)
    monetize_html = _parse_monetize(md_text)
    fb_chart_html = _parse_fb_chart(md_text)
    fb_top3_html = _parse_fb_top3(md_text)
    hours_top3_html = _parse_hours_top3(md_text)
    submit_top3_html = _parse_submit_top3(md_text)

    # --- KPIカード ---
    total_hours = d.get("total_hours", 0)
    avg_hours = d.get("avg_hours", 0)
    submitters = d.get("submitters", 0)
    total_submissions = d.get("total_submissions", 0)
    pos_ratio = d.get("pos_ratio", 0)
    monetize_total = d.get("monetize_total", "—")
    monetize_sub = d.get("monetize_sub", "")
    inactive_count = d.get("inactive_count", 0)
    total_pro = d.get("total_pro", 38)

    prev_hours = prev.get("total_hours", None) if prev else None
    prev_avg = prev.get("avg_hours", None) if prev else None
    prev_pos = prev.get("pos_ratio", None) if prev else None
    prev_inactive = prev.get("inactive_count", None) if prev else None

    hours_change = f'+{round((total_hours/prev_hours-1)*100)}% (前週 {prev_hours}h)' if prev_hours and prev_hours > 0 else ""
    avg_change = f'+{round((avg_hours/prev_avg-1)*100)}% (前週 {prev_avg}h)' if prev_avg and prev_avg > 0 else ""
    pos_change = f'+{pos_ratio - prev_pos}pt (前週 {prev_pos}%)' if prev_pos is not None else ""
    inactive_change = f'改善 (前週 {prev_inactive}名)' if prev_inactive is not None else ""

    # 感情ゲージの前週比較
    prev_neg = prev.get("neg_ratio", None) if prev else None
    neg_ratio = d.get("neg_ratio", 0)
    sentiment_compare = ""
    if prev_pos is not None:
        diff = pos_ratio - prev_pos
        sentiment_compare = f"""
        <div style="display:flex;gap:16px;margin-top:16px;">
          <div style="flex:1;text-align:center;padding:12px;background:#f8fafc;border-radius:8px;">
            <div style="font-size:11px;color:#94a3b8;font-weight:600;">前週</div>
            <div style="display:flex;gap:4px;justify-content:center;margin-top:4px;">
              <span style="color:#16a34a;font-weight:700;font-size:18px;">{prev_pos}%</span>
              <span style="color:#94a3b8;font-size:14px;">/</span>
              <span style="color:#dc2626;font-weight:700;font-size:18px;">{prev_neg}%</span>
            </div>
          </div>
          <div style="flex:1;text-align:center;padding:12px;background:#f8fafc;border-radius:8px;">
            <div style="font-size:11px;color:#94a3b8;font-weight:600;">今週</div>
            <div style="display:flex;gap:4px;justify-content:center;margin-top:4px;">
              <span style="color:#16a34a;font-weight:700;font-size:18px;">{pos_ratio}%</span>
              <span style="color:#94a3b8;font-size:14px;">/</span>
              <span style="color:#dc2626;font-weight:700;font-size:18px;">{neg_ratio}%</span>
            </div>
          </div>
          <div style="flex:1;text-align:center;padding:12px;background:#f8fafc;border-radius:8px;">
            <div style="font-size:11px;color:#94a3b8;font-weight:600;">変化</div>
            <div style="color:{'#16a34a' if diff >= 0 else '#dc2626'};font-weight:700;font-size:18px;margin-top:4px;">{'+' if diff >= 0 else ''}{diff}pt</div>
          </div>
        </div>"""

    # 分析コメント
    if pos_ratio >= 60:
        sentiment_comment = "全体的にポジティブな反応が多く、コミュニティの健康状態は良好。"
    elif pos_ratio >= 40:
        sentiment_comment = "プラスとマイナスが拮抗。停滞会員へのフォロー強化が有効。"
    else:
        sentiment_comment = "マイナス傾向が目立つ。個別フォローと成功体験の共有を優先。"
    if prev_pos is not None and pos_ratio > prev_pos:
        sentiment_comment += "前週比でポジティブ率が改善。"

    # --- 非アクティブ比較カード ---
    active_rate = round((total_pro - inactive_count) / total_pro * 100) if total_pro > 0 else 0
    prev_active_rate = round((total_pro - prev_inactive) / total_pro * 100) if prev_inactive is not None and total_pro > 0 else "—"

    # body contentのみ返す（ページ外殻は generate_monthly_dashboard が担当）
    return f"""<div style="font-size:13px;color:var(--gray-600);margin-bottom:20px;">
  <strong style="font-size:18px;color:var(--gray-800);">{period_full}</strong>
</div>
<div class="kpi-grid">
  <div class="kpi-card"><div class="kpi-label">総稼働時間</div><div class="kpi-value">{total_hours}h</div><div class="kpi-change up">{hours_change}</div></div>
  <div class="kpi-card"><div class="kpi-label">平均稼働時間</div><div class="kpi-value">{avg_hours}h</div><div class="kpi-change up">{avg_change}</div></div>
  <div class="kpi-card"><div class="kpi-label">日報提出者</div><div class="kpi-value">{submitters}名</div><div class="kpi-sub">提出{total_submissions}件</div></div>
  <div class="kpi-card"><div class="kpi-label">ポジティブ率</div><div class="kpi-value">{pos_ratio}%</div><div class="kpi-change up">{pos_change}</div></div>
  <div class="kpi-card"><div class="kpi-label">マネタイズ合計</div><div class="kpi-value">{monetize_total}</div><div class="kpi-sub">{monetize_sub}</div></div>
  <div class="kpi-card"><div class="kpi-label">非アクティブPRO</div><div class="kpi-value">{inactive_count}名</div><div class="kpi-change up">{inactive_change}</div><div class="kpi-sub">PRO全体 {total_pro}名</div></div>
</div>
<div class="section"><div class="section-header" onclick="this.parentElement.classList.toggle('collapsed')">
  <span class="section-num">1</span><span class="section-title">基本統計データ（稼働状況）</span><span class="section-toggle">▼</span>
</div><div class="section-body">
  <h4 style="font-size:14px;color:var(--gray-600);margin-bottom:8px;">稼働時間 TOP3</h4>{hours_top3_html}
  <h4 style="font-size:14px;color:var(--gray-600);margin:16px 0 8px;">日報提出数 TOP3</h4>{submit_top3_html}
</div></div>
<div class="section"><div class="section-header" onclick="this.parentElement.classList.toggle('collapsed')">
  <span class="section-num">2</span><span class="section-title">実践フィードバック会（FB会）分析</span><span class="section-toggle">▼</span>
</div><div class="section-body">
  <h4 style="font-size:14px;color:var(--gray-600);margin-bottom:4px;">曜日別参加者数</h4>{fb_chart_html}
  <h4 style="font-size:14px;color:var(--gray-600);margin:16px 0 8px;">参加回数 TOP3</h4>{fb_top3_html}
</div></div>
<div class="section"><div class="section-header" onclick="this.parentElement.classList.toggle('collapsed')">
  <span class="section-num">3</span><span class="section-title">感情・傾向占有率</span><span class="section-toggle">▼</span>
</div><div class="section-body">
  <div class="sentiment-gauge">
    <div class="sentiment-pos" style="width:{pos_ratio}%;">{pos_ratio}%</div>
    <div class="sentiment-neg" style="width:{neg_ratio}%;">{neg_ratio}%</div>
  </div>
  <div style="display:flex;justify-content:space-between;font-size:12px;color:var(--gray-600);">
    <span>前進・自信・成功体験・学び</span><span>不安・停滞・環境課題・操作の迷い</span>
  </div>
  {sentiment_compare}
  <p style="font-size:13px;color:var(--gray-600);margin-top:12px;">{sentiment_comment}</p>
</div></div>
<div class="section"><div class="section-header" onclick="this.parentElement.classList.toggle('collapsed')">
  <span class="section-num">4</span><span class="section-title">カテゴリー分析（意味分類）</span><span class="section-toggle">▼</span>
</div><div class="section-body">{categories_html}</div></div>
<div class="section"><div class="section-header" onclick="this.parentElement.classList.toggle('collapsed')">
  <span class="section-num">5</span><span class="section-title">マネタイズフェーズ × 課題分析</span><span class="section-toggle">▼</span>
</div><div class="section-body">{phases_html}</div></div>
<div class="section"><div class="section-header" onclick="this.parentElement.classList.toggle('collapsed')">
  <span class="section-num">6</span><span class="section-title">マネタイズ成果ランキング</span><span class="section-toggle">▼</span>
</div><div class="section-body">{monetize_html}</div></div>
<div class="section"><div class="section-header" onclick="this.parentElement.classList.toggle('collapsed')">
  <span class="section-num">7</span><span class="section-title">会員様の生の声</span><span class="section-toggle">▼</span>
</div><div class="section-body">{voices_html}</div></div>
<div class="section"><div class="section-header" onclick="this.parentElement.classList.toggle('collapsed')">
  <span class="section-num">8</span><span class="section-title">運営への提言</span><span class="section-toggle">▼</span>
</div><div class="section-body">{proposals_html}</div></div>
<div class="section"><div class="section-header" onclick="this.parentElement.classList.toggle('collapsed')">
  <span class="section-num">9</span><span class="section-title">非アクティブPRO会員（過去4週傾向）</span><span class="section-toggle">▼</span>
</div><div class="section-body">
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px;">
    <div style="background:var(--accent-light);border-radius:10px;padding:20px;">
      <div style="font-size:13px;font-weight:700;color:var(--accent);margin-bottom:12px;">今週（{period_start_str}–{period_end_str}）</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;">
        <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">非アクティブ</div><div style="font-size:28px;font-weight:800;color:var(--danger);">{inactive_count}名</div></div>
        <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">アクティブ率</div><div style="font-size:28px;font-weight:800;color:var(--success);">{active_rate}%</div></div>
        <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">日報提出者（2回以上）</div><div style="font-size:28px;font-weight:800;color:var(--primary);">{report_2plus_this}名</div><div style="font-size:11px;color:var(--gray-400);">/ PRO {total_pro}名</div></div>
        <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">日報提出率</div><div style="font-size:28px;font-weight:800;color:var(--primary);">{report_rate_this}%</div></div>
        <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">週報提出者</div><div style="font-size:28px;font-weight:800;color:var(--accent);">{weekly_submitters_this}名</div><div style="font-size:11px;color:var(--gray-400);">/ PRO {total_pro}名</div></div>
        <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">週報提出率</div><div style="font-size:28px;font-weight:800;color:var(--accent);">{weekly_rate_this}%</div><div style="font-size:11px;color:var(--gray-400);">目標 80%</div></div>
      </div>
    </div>
    <div style="background:var(--gray-50);border-radius:10px;padding:20px;">
      <div style="font-size:13px;font-weight:700;color:var(--gray-400);margin-bottom:12px;">前週</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;">
        <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">非アクティブ</div><div style="font-size:28px;font-weight:800;color:var(--gray-400);">{prev_inactive if prev_inactive is not None else '—'}名</div></div>
        <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">アクティブ率</div><div style="font-size:28px;font-weight:800;color:var(--gray-400);">{prev_active_rate}%</div></div>
        <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">日報提出者（2回以上）</div><div style="font-size:28px;font-weight:800;color:var(--gray-400);">{report_2plus_prev}名</div></div>
        <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">日報提出率</div><div style="font-size:28px;font-weight:800;color:var(--gray-400);">{report_rate_prev}%</div></div>
        <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">週報提出者</div><div style="font-size:28px;font-weight:800;color:var(--gray-400);">{weekly_submitters_prev}名</div></div>
        <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">週報提出率</div><div style="font-size:28px;font-weight:800;color:var(--gray-400);">{weekly_rate_prev}%</div></div>
      </div>
    </div>
  </div>
  {inactive_table_html}
  {alerts_html}
</div></div>"""


# ==========================================
# MDレポートパーサー（ダッシュボード用）
# ==========================================
def _parse_md_report(md):
    """MDレポートから主要数値を抽出"""
    d = {}
    for line in md.split("\n"):
        l = line.strip()
        # 総稼働時間
        m = re.search(r"総稼働時間：([\d.]+)時間", l)
        if m: d["total_hours"] = float(m.group(1))
        # 平均稼働時間
        m = re.search(r"平均稼働時間：([\d.]+)時間", l)
        if m: d["avg_hours"] = float(m.group(1))
        # ポジティブ率
        m = re.search(r"プラス傾向率：(\d+)％", l)
        if m: d["pos_ratio"] = int(m.group(1))
        # ネガティブ率
        m = re.search(r"マイナス傾向率：(\d+)％", l)
        if m: d["neg_ratio"] = int(m.group(1))
        # 非アクティブ
        m = re.search(r"非アクティブ: (\d+)名 / PRO全体: (\d+)名", l)
        if m:
            d["inactive_count"] = int(m.group(1))
            d["total_pro"] = int(m.group(2))
        # マネタイズ合計
        m = re.search(r"合計マネタイズ額: (.+)", l)
        if m: d["monetize_total"] = m.group(1).strip()
        # 週報提出者 / 収益報告者
        m = re.search(r"週報提出者: (\d+)名 / 収益報告者: (\d+)名", l)
        if m: d["monetize_sub"] = f"収益報告 {m.group(2)}名 / 週報 {m.group(1)}名"
    # 提出者数・提出件数
    for line in md.split("\n"):
        l = line.strip()
        m = re.search(r"日報提出者：(\d+)名（提出(\d+)件）", l)
        if m:
            d["submitters"] = int(m.group(1))
            d["total_submissions"] = int(m.group(2))
        m = re.search(r"日報2回以上提出者：(\d+)名（提出率(\d+)%）", l)
        if m:
            d["report_2plus"] = int(m.group(1))
            d["report_rate"] = int(m.group(2))
        m = re.search(r"週報提出率: (\d+)%（(\d+)/(\d+)名）", l)
        if m:
            d["weekly_rate"] = int(m.group(1))
            d["weekly_submitters"] = int(m.group(2))
    return d


def _load_prev_week_data(path=None):
    """前週レポートのMDファイルを読み込んでパース"""
    if path and Path(path).exists():
        md = Path(path).read_text(encoding="utf-8")
        return _parse_md_report(md)
    # 自動検索: 前週のMDファイルを探す
    prev_mon = LAST_MONDAY - timedelta(days=7)
    prev_sun = LAST_MONDAY - timedelta(days=1)
    start_str = f"{prev_mon.month}月{prev_mon.day}日"
    end_str = f"{prev_sun.month}月{prev_sun.day}日"
    prev_path = REPORTS_DIR / f"会員分析（日報）{start_str}-{end_str}.md"
    if prev_path.exists():
        md = prev_path.read_text(encoding="utf-8")
        return _parse_md_report(md)
    return None




def _parse_inactive_table(md):
    """非アクティブテーブルをHTMLに変換"""
    lines = md.split("\n")
    table_lines = [l.strip() for l in lines if l.strip().startswith("|")]
    if not table_lines:
        return ""
    html = ['<div style="overflow-x:auto;"><table><thead><tr>']
    header = [c.strip() for c in table_lines[0].strip("|").split("|")]
    for h in header:
        html.append(f"<th>{_esc(h)}</th>")
    html.append("</tr></thead><tbody>")
    for row_str in table_lines[1:]:
        cells = [c.strip() for c in row_str.strip("|").split("|")]
        html.append("<tr>")
        for c in cells:
            cls = ""
            if c in ("0/7", "0%", "0/0"):
                cls = ' class="zero"'
            elif "14%" in c or "29%" in c:
                cls = ' class="low"'
            html.append(f"<td{cls}>{_esc(c)}</td>")
        html.append("</tr>")
    html.append("</tbody></table></div>")
    return "\n".join(html)


def _parse_alerts(md):
    """アラート行をHTMLに変換"""
    html = []
    for line in md.split("\n"):
        l = line.strip()
        if l.startswith("⚠️"):
            content = l.replace("⚠️ ", "").replace("⚠️", "")
            html.append(f'<div class="alert danger"><span class="alert-icon">⚠️</span><div>{_esc(content)}</div></div>')
        elif l.startswith("📉"):
            content = l.replace("📉 ", "").replace("📉", "")
            html.append(f'<div class="alert warning"><span class="alert-icon">📉</span><div>{_esc(content)}</div></div>')
    return "\n".join(html)


def _parse_voices(md):
    """生の声セクションをHTMLカードに変換"""
    lines = md.split("\n")
    in_voices = False
    pos_voices = []
    neg_voices = []
    current_list = None
    for line in lines:
        l = line.strip()
        if "⑦" in l or "重要インサイト" in l:
            in_voices = True
            continue
        if in_voices and l and l[0] in "⑧⑨":
            break
        if not in_voices:
            continue
        if "プラス傾向" in l:
            current_list = pos_voices
            continue
        if "マイナス傾向" in l:
            current_list = neg_voices
            continue
        if current_list is not None and line.startswith("  ") and "：「" in l:
            parts = l.split("：「", 1)
            name = parts[0].strip()
            voice = parts[1].rstrip("」").rstrip("」") if len(parts) > 1 else ""
            if len(voice) > 120:
                voice = voice[:120] + "…"
            current_list.append((name, voice))

    pos_html = '<div class="voice-section-label pos">前進を支える声</div>\n'
    for name, voice in pos_voices:
        pos_html += f'<div class="voice-card pos"><div class="voice-name">{_esc(name)}</div>{_esc(voice)}</div>\n'
    neg_html = '<div class="voice-section-label neg">停滞を招く声</div>\n'
    seen = set()
    for name, voice in neg_voices:
        key = f"{name}:{voice[:30]}"
        if key in seen:
            continue
        seen.add(key)
        neg_html += f'<div class="voice-card neg"><div class="voice-name">{_esc(name)}</div>{_esc(voice)}</div>\n'

    return f'<div class="voice-grid"><div>{pos_html}</div><div>{neg_html}</div></div>'


def _parse_proposals(md):
    """提言セクションをHTMLに変換"""
    lines = md.split("\n")
    in_proposals = False
    follow_html = []
    improve_html = []
    current = None
    for line in lines:
        l = line.strip()
        if "⑧" in l or "運営への提言" in l:
            in_proposals = True
            continue
        if in_proposals and l and l[0] in "⑨":
            break
        if not in_proposals:
            continue
        if "個別フォロー" in l:
            current = "follow"
            continue
        if "FB会" in l or "改善案" in l:
            current = "improve"
            continue
        if current == "follow" and line.startswith("  ") and l:
            follow_html.append(f'<div class="proposal-item"><div class="proposal-icon follow">👤</div><div>{_esc(l)}</div></div>')
        elif current == "improve" and line.startswith("  ") and l:
            improve_html.append(f'<div class="proposal-item"><div class="proposal-icon improve">💡</div><div>{_esc(l)}</div></div>')

    html = '<h4 style="font-size:14px;color:var(--gray-600);margin-bottom:8px;">個別フォロー推奨</h4>\n'
    html += '<div class="proposal-list">' + "\n".join(follow_html) + '</div>\n'
    html += '<h4 style="font-size:14px;color:var(--gray-600);margin:20px 0 8px;">FB会・サービス改善案</h4>\n'
    html += '<div class="proposal-list">' + "\n".join(improve_html) + '</div>'
    return html


def _parse_categories_bars(md):
    """カテゴリー分析をバーチャートHTMLに変換"""
    colors = {"前進実感": "green", "スキル不足": "orange", "行動面": "blue", "時間・環境": "purple", "心理的": "red"}
    bars = []
    for line in md.split("\n"):
        l = line.strip()
        if not l or l[0] == "④":
            continue
        m = re.search(r"(.+?)：(\d+)人（(\d+)％）", l)
        if m and ("④" in md.split(l)[0][-50:] if len(md.split(l)) > 1 else True):
            label = m.group(1).strip()
            count = m.group(2)
            pct = int(m.group(3))
            color = "blue"
            for key, c in colors.items():
                if key in label:
                    color = c
                    break
            short_label = label[:8] if len(label) > 8 else label
            bars.append(f'<div class="bar-row"><span class="bar-label">{_esc(short_label)}</span>'
                       f'<div class="bar-track"><div class="bar-fill {color}" style="width:{pct}%;">{count}人</div></div>'
                       f'<span class="bar-value">{pct}%</span></div>')
    if not bars:
        return "<p>（データなし）</p>"
    return '<div class="bar-chart">' + "\n".join(bars) + '</div>'


def _parse_phase_cards(md):
    """フェーズ分析をカードHTMLに変換"""
    cards = []
    phase_classes = ["p1", "p2", "p3", "p4"]
    idx = 0
    for line in md.split("\n"):
        l = line.strip()
        m = re.search(r"フェーズ(\d)（(.+?)）：(\d+)人[\[［]該当[：:](.+?)[\]］]", l)
        if m:
            phase_num = m.group(1)
            phase_name = m.group(2)
            count = m.group(3)
            members = m.group(4).strip()
            cls = phase_classes[idx % 4]
            cards.append(f'<div class="phase-card {cls}">'
                        f'<div style="font-size:11px;font-weight:700;color:var(--gray-400);">PHASE {phase_num}</div>'
                        f'<div style="font-size:14px;font-weight:700;margin:2px 0;">{_esc(phase_name)}</div>'
                        f'<div style="font-size:24px;font-weight:800;color:var(--primary);">{count}人</div>'
                        f'<div style="font-size:12px;color:var(--gray-600);margin-top:4px;">{_esc(members)}</div></div>')
            idx += 1
    if not cards:
        return "<p>（データなし）</p>"
    return '<div class="phase-grid">' + "\n".join(cards) + '</div>'


def _parse_monetize(md):
    """マネタイズセクションをHTMLに変換"""
    lines = md.split("\n")
    in_section = False
    ranking = []
    total_amount = ""
    avg_amount = ""
    for line in lines:
        l = line.strip()
        if "⑥" in l:
            in_section = True
            continue
        if in_section and l and l[0] in "⑦⑧⑨":
            break
        if not in_section:
            continue
        m = re.search(r"合計マネタイズ額: (.+)", l)
        if m:
            total_amount = m.group(1).strip()
        m = re.search(r"平均: (.+)", l)
        if m:
            avg_amount = m.group(1).strip()
        m = re.search(r"[🥇🥈🥉]\s*(.+?)\s*—\s*(.+)", l)
        if m:
            ranking.append((m.group(1).strip(), m.group(2).strip()))
        m2 = re.search(r"\d+位\s+(.+?)\s*—\s*(.+)", l)
        if m2:
            ranking.append((m2.group(1).strip(), m2.group(2).strip()))

    if not ranking:
        return "<p>（今週は収益報告なし）</p>"

    # 表彰台
    podium = ""
    if len(ranking) >= 3:
        podium = f"""<div class="monetize-podium">
          <div class="podium-item silver"><div class="podium-medal">🥈</div><div class="podium-name">{_esc(ranking[1][0])}</div><div class="podium-amount">{_esc(ranking[1][1])}</div></div>
          <div class="podium-item gold"><div class="podium-medal">🥇</div><div class="podium-name">{_esc(ranking[0][0])}</div><div class="podium-amount">{_esc(ranking[0][1])}</div></div>
          <div class="podium-item bronze"><div class="podium-medal">🥉</div><div class="podium-name">{_esc(ranking[2][0])}</div><div class="podium-amount">{_esc(ranking[2][1])}</div></div>
        </div>"""

    # 4位以下テーブル
    rest_html = ""
    if len(ranking) > 3:
        rest_html = '<table><thead><tr><th>順位</th><th>会員名</th><th>金額</th></tr></thead><tbody>'
        for i, (name, amount) in enumerate(ranking[3:], 4):
            rest_html += f'<tr><td>{i}位</td><td>{_esc(name)}</td><td>{_esc(amount)}</td></tr>'
        rest_html += '</tbody></table>'

    return f"""<div style="display:flex;gap:16px;margin-bottom:16px;">
      <div style="flex:1;text-align:center;padding:12px;background:var(--gray-50);border-radius:8px;">
        <div style="font-size:11px;color:var(--gray-400);font-weight:600;">合計額</div>
        <div style="font-size:22px;font-weight:800;color:var(--primary);">{_esc(total_amount)}</div>
      </div>
      <div style="flex:1;text-align:center;padding:12px;background:var(--gray-50);border-radius:8px;">
        <div style="font-size:11px;color:var(--gray-400);font-weight:600;">平均額</div>
        <div style="font-size:22px;font-weight:800;color:var(--primary);">{_esc(avg_amount)}</div>
      </div>
    </div>{podium}{rest_html}"""


def _parse_fb_chart(md):
    """FB会曜日別データをバーチャートHTMLに変換"""
    for line in md.split("\n"):
        m = re.search(r"曜日別参加者数：(.+)", line.strip())
        if m:
            parts = m.group(1).split("、")
            days_data = []
            for p in parts:
                dm = re.search(r"(.)：(\d+)人", p)
                if dm:
                    days_data.append((dm.group(1), int(dm.group(2))))
            if not days_data:
                return ""
            max_val = max(d[1] for d in days_data) or 1
            total = sum(d[1] for d in days_data)
            bars = ""
            for day_name, count in days_data:
                h = int(count / max_val * 100) if count > 0 else 0
                style = f'height:{h}px;' if count > 0 else 'height:0px;background:#cbd5e1;'
                bars += f'<div class="day-col"><div class="day-bar" style="{style}">{count if count > 0 else ""}</div><div class="day-label">{day_name}</div></div>\n'
            return f'<div class="day-chart">{bars}</div><div style="font-size:12px;color:var(--gray-400);text-align:center;">合計 {total}件</div>'
    return ""


def _parse_fb_top3(md):
    """FB会TOP3をHTMLに変換"""
    medals = ["🥇", "🥈", "🥉"]
    items = []
    in_fb = False
    for line in md.split("\n"):
        l = line.strip()
        if "参加回数TOP3" in l:
            in_fb = True
            continue
        if in_fb and line.startswith("  "):
            # "1位：川西秀樹（7回）、2位：..." のような形式
            for part in l.split("、"):
                m = re.search(r"(\d)位[：:](.+?)（(\d+)回）", part)
                if m:
                    idx = int(m.group(1)) - 1
                    medal = medals[idx] if idx < 3 else ""
                    items.append(f'<div class="top3-item"><span class="rank">{medal}</span><div><div class="top3-name">{_esc(m.group(2))}</div><div class="top3-detail">{m.group(3)}回</div></div></div>')
            break
    return '<div class="top3">' + "\n".join(items) + '</div>' if items else ""


def _parse_hours_top3(md):
    """稼働時間TOP3をHTMLに変換"""
    medals = ["🥇", "🥈", "🥉"]
    items = []
    in_section = False
    for line in md.split("\n"):
        l = line.strip()
        if "稼働時間TOP3" in l:
            in_section = True
            continue
        if in_section and "日報提出数" in l:
            break
        if in_section and line.startswith("  "):
            m = re.search(r"(\d)位[：:](.+?)（総稼働時間[：:](.+?)h\s*/\s*平均[：:](.+?)h）", l)
            if m:
                idx = int(m.group(1)) - 1
                medal = medals[idx] if idx < 3 else ""
                items.append(f'<div class="top3-item"><span class="rank">{medal}</span><div>'
                           f'<div class="top3-name">{_esc(m.group(2))}</div>'
                           f'<div class="top3-detail">{m.group(3)}h（平均 {m.group(4)}h/日）</div></div></div>')
    return '<div class="top3">' + "\n".join(items) + '</div>' if items else ""


def _parse_submit_top3(md):
    """日報提出TOP3をHTMLに変換"""
    medals = ["🥇", "🥈", "🥉"]
    items = []
    for line in md.split("\n"):
        l = line.strip()
        if "日報提出数TOP3" in l:
            continue
        if line.startswith("  ") and "位：" in l and "回）" in l:
            for part in l.split("、"):
                m = re.search(r"(\d)位[：:](.+?)（(\d+)回）", part)
                if m:
                    idx = int(m.group(1)) - 1
                    medal = medals[idx] if idx < 3 else ""
                    items.append(f'<div class="top3-item"><span class="rank">{medal}</span><div>'
                               f'<div class="top3-name">{_esc(m.group(2))}</div>'
                               f'<div class="top3-detail">{m.group(3)}回</div></div></div>')
            break
    return '<div class="top3">' + "\n".join(items) + '</div>' if items else ""


def save_report(report, report_type="weekly"):
    """レポートをMDで保存し、月次ダッシュボードを自動更新"""
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    start_str = f"{LAST_MONDAY.month}月{LAST_MONDAY.day}日"
    end_str = f"{LAST_SUNDAY.month}月{LAST_SUNDAY.day}日"
    filename = f"会員分析（日報）{start_str}-{end_str}"

    # MD保存
    md_path = REPORTS_DIR / f"{filename}.md"
    md_path.write_text(report, encoding="utf-8")
    print(f"💾 レポート保存(MD): {md_path}")

    # 月次ダッシュボード（タブ切替式の単一HTMLページ）を自動更新
    generate_monthly_dashboard(LAST_MONDAY.month, LAST_MONDAY.year)

    return md_path


def generate_monthly_dashboard(month, year):
    """月次ダッシュボード — 全週をタブ切替で表示する単一HTMLページ"""
    from datetime import date

    # 当月の週次MDファイルを全て探す
    weeks = []  # [(monday_date, md_text), ...]
    first_day = date(year, month, 1)
    wd = first_day.weekday()
    mon = first_day - timedelta(days=wd)
    if mon.month != month and mon.day > 7:
        mon += timedelta(days=7)

    while True:
        sun = mon + timedelta(days=6)
        if mon.month > month:
            break
        start_str = f"{mon.month}月{mon.day}日"
        end_str = f"{sun.month}月{sun.day}日"
        md_path = REPORTS_DIR / f"会員分析（日報）{start_str}-{end_str}.md"
        if md_path.exists():
            md = md_path.read_text(encoding="utf-8")
            weeks.append((mon, md))
        mon += timedelta(days=7)
        if mon.day > 28 and mon.month > month:
            break

    if not weeks:
        return

    # 各週のbodyコンテンツを生成
    week_bodies = []
    for i, (monday, md_text) in enumerate(weeks):
        prev_md = weeks[i - 1][1] if i > 0 else None
        body = _build_week_body(md_text, monday, prev_md)
        sun = monday + timedelta(days=6)
        label = f"W{i+1}: {monday.month}/{monday.day}–{sun.month}/{sun.day}"
        week_bodies.append((label, body))

    # 月間KPIサマリー用データ
    weeks_parsed = [_parse_md_report(md) for _, md in weeks]
    total_hours_sum = round(sum(w.get("total_hours", 0) for w in weeks_parsed), 1)
    avg_pos = round(sum(w.get("pos_ratio", 0) for w in weeks_parsed) / len(weeks_parsed))
    latest_inactive = weeks_parsed[-1].get("inactive_count", 0)

    # 月間トレンドバー
    max_hours = max((w.get("total_hours", 0) for w in weeks_parsed), default=1) or 1
    max_inactive = max((w.get("inactive_count", 0) for w in weeks_parsed), default=1) or 1
    trend_hours = ""
    trend_pos = ""
    trend_inactive = ""
    for i, (monday, _) in enumerate(weeks):
        sun = monday + timedelta(days=6)
        lbl = f"{monday.month}/{monday.day}"
        w = weeks_parsed[i]
        h = w.get("total_hours", 0)
        p = w.get("pos_ratio", 0)
        ic = w.get("inactive_count", 0)
        trend_hours += f'<div class="bar-row"><span class="bar-label">{lbl}</span><div class="bar-track"><div class="bar-fill blue" style="width:{int(h/max_hours*100)}%;">{h}h</div></div></div>\n'
        trend_pos += f'<div class="bar-row"><span class="bar-label">{lbl}</span><div class="bar-track"><div class="bar-fill green" style="width:{p}%;">{p}%</div></div></div>\n'
        trend_inactive += f'<div class="bar-row"><span class="bar-label">{lbl}</span><div class="bar-track"><div class="bar-fill red" style="width:{int(ic/max_inactive*100)}%;">{ic}名</div></div></div>\n'

    # 週タブHTML
    week_tabs_html = ""
    for i, (label, _) in enumerate(week_bodies):
        active = " active" if i == len(week_bodies) - 1 else ""
        week_tabs_html += f'<div class="week-tab{active}" onclick="switchWeek({i})">{label}</div>\n'

    # 週コンテンツHTML
    week_contents_html = ""
    for i, (_, body) in enumerate(week_bodies):
        display = "block" if i == len(week_bodies) - 1 else "none"
        week_contents_html += f'<div class="week-content" id="week-{i}" style="display:{display};">\n{body}\n</div>\n'

    # 月間サマリータブ（常に先頭に表示）
    summary_body = f"""<div class="kpi-grid">
  <div class="kpi-card"><div class="kpi-label">レポート週数</div><div class="kpi-value">{len(weeks)}週</div></div>
  <div class="kpi-card"><div class="kpi-label">月間総稼働</div><div class="kpi-value">{total_hours_sum}h</div></div>
  <div class="kpi-card"><div class="kpi-label">平均ポジティブ率</div><div class="kpi-value">{avg_pos}%</div></div>
  <div class="kpi-card"><div class="kpi-label">最新非アクティブ</div><div class="kpi-value">{latest_inactive}名</div></div>
</div>
<div class="section"><div class="section-header"><span class="section-title">総稼働時間トレンド</span></div>
<div class="section-body"><div class="bar-chart">{trend_hours}</div></div></div>
<div class="section"><div class="section-header"><span class="section-title">ポジティブ率トレンド</span></div>
<div class="section-body"><div class="bar-chart">{trend_pos}</div></div></div>
<div class="section"><div class="section-header"><span class="section-title">非アクティブ数トレンド</span></div>
<div class="section-body"><div class="bar-chart">{trend_inactive}</div></div></div>"""

    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PRO PREMIUM TEAM {month}月 ダッシュボード</title>
<style>
:root {{
  --primary: #1e3a5f; --accent: #2b6cb0; --accent-light: #ebf4ff;
  --success: #16a34a; --warning: #d97706; --danger: #dc2626;
  --gray-50: #f8fafc; --gray-100: #f1f5f9; --gray-200: #e2e8f0;
  --gray-300: #cbd5e1; --gray-400: #94a3b8; --gray-600: #475569; --gray-800: #1e293b;
}}
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Hiragino Kaku Gothic ProN','Yu Gothic UI','Meiryo',sans-serif; background:var(--gray-100); color:var(--gray-800); line-height:1.7; }}
.header {{ background:linear-gradient(135deg,var(--primary),var(--accent)); color:#fff; padding:24px 40px 0; }}
.header h1 {{ font-size:24px; font-weight:700; }}
.header .subtitle {{ font-size:13px; opacity:0.85; margin-top:2px; }}
.tab-bar {{ display:flex; gap:0; margin-top:16px; overflow-x:auto; }}
.tab-bar .tab {{ padding:10px 20px; cursor:pointer; font-size:13px; font-weight:600; color:rgba(255,255,255,0.7); border-bottom:3px solid transparent; white-space:nowrap; transition:all 0.2s; }}
.tab-bar .tab:hover {{ color:#fff; background:rgba(255,255,255,0.1); }}
.tab-bar .tab.active {{ color:#fff; border-bottom-color:#fff; background:rgba(255,255,255,0.1); }}
.week-nav {{ background:#fff; border-bottom:1px solid var(--gray-200); padding:0 40px; display:flex; gap:0; overflow-x:auto; }}
.week-tab {{ padding:12px 20px; cursor:pointer; font-size:13px; font-weight:600; color:var(--gray-400); border-bottom:3px solid transparent; white-space:nowrap; transition:all 0.2s; }}
.week-tab:hover {{ color:var(--gray-800); background:var(--gray-50); }}
.week-tab.active {{ color:var(--accent); border-bottom-color:var(--accent); }}
.week-tab.disabled {{ color:var(--gray-300); cursor:default; }}
.main {{ max-width:1100px; margin:0 auto; padding:28px 20px 60px; }}
.kpi-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:16px; margin-bottom:28px; }}
.kpi-card {{ background:#fff; border-radius:10px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.06); text-align:center; }}
.kpi-label {{ font-size:11px; color:var(--gray-400); font-weight:600; text-transform:uppercase; }}
.kpi-value {{ font-size:28px; font-weight:800; color:var(--primary); margin:4px 0; }}
.kpi-change {{ font-size:11px; font-weight:600; }}
.kpi-change.up {{ color:var(--success); }}
.kpi-change.down {{ color:var(--danger); }}
.kpi-sub {{ font-size:11px; color:var(--gray-400); }}
.section {{ background:#fff; border-radius:10px; box-shadow:0 1px 3px rgba(0,0,0,0.06); margin-bottom:20px; overflow:hidden; }}
.section.collapsed .section-body {{ display:none; }}
.section-header {{ padding:14px 24px; border-bottom:1px solid var(--gray-200); cursor:pointer; display:flex; align-items:center; gap:10px; }}
.section-header:hover {{ background:var(--gray-50); }}
.section-num {{ background:var(--accent); color:#fff; width:24px; height:24px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:12px; font-weight:700; flex-shrink:0; }}
.section-title {{ font-size:15px; font-weight:700; color:var(--gray-800); flex:1; }}
.section-toggle {{ font-size:12px; color:var(--gray-400); }}
.section-body {{ padding:20px 24px; }}
.bar-row {{ display:flex; align-items:center; margin:5px 0; }}
.bar-label {{ width:90px; font-size:12px; color:var(--gray-600); flex-shrink:0; }}
.bar-track {{ flex:1; height:22px; background:var(--gray-100); border-radius:4px; overflow:hidden; }}
.bar-fill {{ height:100%; border-radius:4px; display:flex; align-items:center; padding-left:8px; font-size:11px; font-weight:600; color:#fff; min-width:fit-content; }}
.bar-fill.blue {{ background:linear-gradient(90deg,#3b82f6,#2563eb); }}
.bar-fill.green {{ background:linear-gradient(90deg,#22c55e,#16a34a); }}
.bar-fill.red {{ background:linear-gradient(90deg,#ef4444,#dc2626); }}
.bar-fill.orange {{ background:linear-gradient(90deg,#f59e0b,#d97706); }}
.bar-fill.purple {{ background:linear-gradient(90deg,#8b5cf6,#7c3aed); }}
.bar-value {{ width:50px; text-align:right; font-size:12px; font-weight:600; color:var(--gray-600); margin-left:8px; }}
.sentiment-gauge {{ display:flex; height:32px; border-radius:6px; overflow:hidden; margin:8px 0; }}
.sentiment-pos {{ background:linear-gradient(90deg,#22c55e,#16a34a); display:flex; align-items:center; justify-content:center; color:#fff; font-weight:700; font-size:13px; }}
.sentiment-neg {{ background:linear-gradient(90deg,#ef4444,#dc2626); display:flex; align-items:center; justify-content:center; color:#fff; font-weight:700; font-size:13px; }}
.voice-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; }}
.voice-section-label {{ font-size:13px; font-weight:700; margin-bottom:8px; }}
.voice-section-label.pos {{ color:var(--success); }}
.voice-section-label.neg {{ color:var(--danger); }}
.voice-card {{ padding:12px; border-radius:8px; margin-bottom:8px; font-size:13px; line-height:1.6; }}
.voice-card.pos {{ background:#f0fdf4; border-left:3px solid var(--success); }}
.voice-card.neg {{ background:#fef2f2; border-left:3px solid var(--danger); }}
.voice-name {{ font-weight:700; font-size:12px; color:var(--gray-600); margin-bottom:2px; }}
.phase-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:12px; }}
.phase-card {{ padding:16px; border-radius:8px; }}
.phase-card.p1 {{ background:#f0fdf4; border-left:4px solid var(--success); }}
.phase-card.p2 {{ background:#ebf4ff; border-left:4px solid var(--accent); }}
.phase-card.p3 {{ background:#fffbeb; border-left:4px solid var(--warning); }}
.phase-card.p4 {{ background:#fef2f2; border-left:4px solid var(--danger); }}
.monetize-podium {{ display:flex; justify-content:center; gap:16px; margin:16px 0; align-items:flex-end; }}
.podium-item {{ text-align:center; padding:12px; border-radius:8px; background:var(--gray-50); min-width:120px; }}
.podium-item.gold {{ background:#fffbeb; border:2px solid #f59e0b; transform:scale(1.05); }}
.podium-item.silver {{ background:#f8fafc; border:2px solid #94a3b8; }}
.podium-item.bronze {{ background:#fff7ed; border:2px solid #d97706; }}
.podium-medal {{ font-size:24px; }}
.podium-name {{ font-weight:700; font-size:14px; margin:4px 0; }}
.podium-amount {{ font-size:13px; color:var(--gray-600); }}
.top3 {{ display:flex; gap:12px; }}
.top3-item {{ display:flex; align-items:center; gap:8px; padding:8px 12px; background:var(--gray-50); border-radius:8px; flex:1; }}
.rank {{ font-size:20px; }}
.top3-name {{ font-weight:700; font-size:13px; }}
.top3-detail {{ font-size:12px; color:var(--gray-600); }}
.day-chart {{ display:flex; align-items:flex-end; gap:8px; height:120px; padding:10px 0; }}
.day-col {{ flex:1; display:flex; flex-direction:column; align-items:center; }}
.day-bar {{ width:100%; background:linear-gradient(180deg,#3b82f6,#2563eb); border-radius:4px 4px 0 0; display:flex; align-items:flex-start; justify-content:center; color:#fff; font-size:11px; font-weight:700; padding-top:4px; min-height:2px; }}
.day-label {{ font-size:12px; color:var(--gray-600); margin-top:4px; }}
.proposal-list {{ margin-bottom:8px; }}
.proposal-item {{ display:flex; align-items:flex-start; gap:8px; padding:8px 0; border-bottom:1px solid var(--gray-100); font-size:13px; }}
.proposal-icon {{ font-size:16px; flex-shrink:0; }}
.alert {{ display:flex; align-items:flex-start; gap:8px; padding:12px 16px; border-radius:8px; margin-top:8px; font-size:13px; }}
.alert.danger {{ background:#fef2f2; color:#991b1b; }}
.alert.warning {{ background:#fffbeb; color:#92400e; }}
.alert-icon {{ font-size:16px; flex-shrink:0; }}
table {{ width:100%; border-collapse:collapse; font-size:13px; }}
th {{ background:var(--gray-50); padding:8px 12px; text-align:left; font-weight:600; border-bottom:2px solid var(--gray-200); }}
td {{ padding:8px 12px; border-bottom:1px solid var(--gray-100); }}
td.zero {{ color:var(--danger); font-weight:700; }}
td.low {{ color:var(--warning); font-weight:600; }}
.float-top {{ position:fixed; bottom:24px; right:24px; background:var(--accent); color:#fff; border:none; width:44px; height:44px; border-radius:50%; cursor:pointer; font-size:20px; box-shadow:0 2px 8px rgba(0,0,0,0.15); display:flex; align-items:center; justify-content:center; z-index:100; }}
.footer {{ text-align:center; padding:24px; color:var(--gray-400); font-size:12px; }}
@media print {{ .week-nav,.float-top,.tab-bar {{ display:none; }} .main {{ padding:0; }} }}
@media (max-width:768px) {{
  .kpi-grid {{ grid-template-columns:repeat(2,1fr); }}
  .voice-grid {{ grid-template-columns:1fr; }}
  .header {{ padding:16px 20px 0; }}
  .week-nav {{ padding:0 12px; }}
  .main {{ padding:16px 12px 40px; }}
}}
</style>
</head>
<body>
<div class="header">
  <h1>PRO PREMIUM TEAM ダッシュボード</h1>
  <div class="subtitle">{year}年{month}月 — 週次分析レポート</div>
  <div class="tab-bar">
    <div class="tab active" onclick="switchView('summary')">月間サマリー</div>
    <div class="tab" onclick="switchView('weekly')">週次詳細</div>
  </div>
</div>
<div class="week-nav" id="weekNav" style="display:none;">
  {week_tabs_html}
</div>
<div class="main">
  <div id="view-summary">
    {summary_body}
  </div>
  <div id="view-weekly" style="display:none;">
    {week_contents_html}
  </div>
</div>
<button class="float-top" onclick="scrollTo({{top:0,behavior:'smooth'}})">↑</button>
<div class="footer">Generated by SAEPIN — PRO PREMIUM TEAM AI Agent System</div>
<script>
function switchView(view) {{
  document.getElementById('view-summary').style.display = view === 'summary' ? 'block' : 'none';
  document.getElementById('view-weekly').style.display = view === 'weekly' ? 'block' : 'none';
  document.getElementById('weekNav').style.display = view === 'weekly' ? 'flex' : 'none';
  document.querySelectorAll('.tab-bar .tab').forEach(function(t) {{ t.classList.remove('active'); }});
  document.querySelectorAll('.tab-bar .tab').forEach(function(t) {{
    if ((view === 'summary' && t.textContent.includes('月間')) || (view === 'weekly' && t.textContent.includes('週次'))) t.classList.add('active');
  }});
  scrollTo({{top:0,behavior:'smooth'}});
}}
function switchWeek(idx) {{
  document.querySelectorAll('.week-content').forEach(function(el) {{ el.style.display = 'none'; }});
  document.getElementById('week-' + idx).style.display = 'block';
  document.querySelectorAll('.week-tab').forEach(function(t) {{ t.classList.remove('active'); }});
  document.querySelectorAll('.week-tab')[idx].classList.add('active');
  scrollTo({{top:0,behavior:'smooth'}});
}}
</script>
</body></html>"""

    dashboard_path = REPORTS_DIR / f"dashboard_{month}月.html"
    dashboard_path.write_text(html, encoding="utf-8")
    print(f"💾 月次ダッシュボード保存: {dashboard_path}")


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

    # Google Sheets接続（report/dry-run/awardはFB会データも必要）
    print("🔗 Google Sheets接続中...")
    if mode in ("award", "report", "dry-run", "remind"):
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

    # FB会データ取得（report/dry-runで使用）
    fb_data = None
    name_map = None
    if spreadsheet and mode in ("report", "dry-run"):
        print("📥 実践FB会データ取得中...")
        fb_data = fetch_fb_data(spreadsheet)
        print(f"  → {len(fb_data)}件のFB会参加を取得")
        name_map = build_discord_name_map(weekly_data, fb_data, spreadsheet)

    # モード別処理
    if mode == "report" or mode == "dry-run":
        report = format_weekly_report(submissions, sentiment, weekly_data, fb_data, name_map, spreadsheet)
        save_report(report, "weekly")
        print("--- 週次レポート ---")
        print(report)
        print("--- ここまで ---\n")

        if mode == "report":
            send_to_discord(report)

    elif mode == "remind":
        print("📥 DM対象者分析中（2週連続80%以下）...")
        dm_data = analyze_dm_targets(spreadsheet)
        print(f"  → 該当者: {len(dm_data['targets'])}名 / PRO全体: {dm_data['total_pro']}名")
        print(f"  → うちPRO+PREMIUM: {dm_data['premium_count']}名")
        reminder = format_reminder_list(dm_data)
        if reminder:
            print("--- リマインドDMリスト ---")
            print(reminder)
            print("--- ここまで ---\n")
            send_to_discord(reminder)
        else:
            print("✨ 2週連続80%以下の会員なし — リマインド不要です")

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
