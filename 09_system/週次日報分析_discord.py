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

            if name and name != "テスト":
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
            lines.append(f"週報提出者: {monetize_ranking['total_members']}名 / 収益報告者: {monetize_ranking['earners']}名")
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


def format_html_report(md_report):
    """MDレポートをHTML形式に変換"""
    period_start = f"{LAST_MONDAY.month}月{LAST_MONDAY.day}日"
    period_end = f"{LAST_SUNDAY.month}月{LAST_SUNDAY.day}日"

    lines = md_report.split("\n")
    body_parts = []
    in_table = False
    table_rows = []

    for line in lines:
        stripped = line.strip()

        # テーブル行の処理
        if stripped.startswith("|"):
            if not in_table:
                in_table = True
                table_rows = []
            table_rows.append(stripped)
            continue
        elif in_table:
            # テーブル終了 → HTMLテーブルに変換
            body_parts.append(_render_html_table(table_rows))
            in_table = False
            table_rows = []

        # タイトル行
        if stripped.rstrip() in ("週次報告書", "週次報告書　"):
            body_parts.append('<h1 class="report-title">週次報告書</h1>')
            continue
        if stripped == "成功習慣行動率100％":
            body_parts.append(f'<p class="subtitle">成功習慣行動率100％</p>')
            continue

        # セクション見出し（① 〜 ⑧）
        if stripped and stripped[0] in "①②③④⑤⑥⑦⑧⑨":
            body_parts.append(f'<h2 class="section-title">{_esc(stripped)}</h2>')
            continue

        # 警告・急降下マーク
        if stripped.startswith("⚠️") or stripped.startswith("📉"):
            cls = "alert-danger" if "⚠️" in stripped else "alert-warning"
            body_parts.append(f'<div class="{cls}">{_esc(stripped)}</div>')
            continue

        # インデント付き行（2スペース以上）
        if line.startswith("  ") and stripped:
            body_parts.append(f'<p class="detail">{_esc(stripped)}</p>')
            continue

        # 空行
        if not stripped:
            continue

        # 通常行
        body_parts.append(f'<p>{_esc(stripped)}</p>')

    # テーブルが末尾にある場合
    if in_table and table_rows:
        body_parts.append(_render_html_table(table_rows))

    body_html = "\n".join(body_parts)

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>週次報告書 {period_start}〜{period_end}</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{
    font-family: 'Hiragino Kaku Gothic ProN', 'Yu Gothic', 'Meiryo', sans-serif;
    background: #f5f7fa;
    color: #333;
    line-height: 1.8;
    padding: 20px;
}}
.container {{
    max-width: 900px;
    margin: 0 auto;
    background: #fff;
    border-radius: 12px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.08);
    padding: 40px;
}}
.report-title {{
    font-size: 28px;
    color: #1a365d;
    border-bottom: 3px solid #2b6cb0;
    padding-bottom: 12px;
    margin-bottom: 4px;
}}
.subtitle {{
    font-size: 16px;
    color: #2b6cb0;
    font-weight: bold;
    margin-bottom: 30px;
}}
.section-title {{
    font-size: 18px;
    color: #fff;
    background: linear-gradient(135deg, #2b6cb0, #3182ce);
    padding: 10px 16px;
    border-radius: 6px;
    margin: 28px 0 12px 0;
}}
p {{
    margin: 6px 0;
    font-size: 14px;
}}
.detail {{
    margin: 4px 0 4px 20px;
    font-size: 14px;
    color: #4a5568;
}}
table {{
    width: 100%;
    border-collapse: collapse;
    margin: 12px 0;
    font-size: 13px;
}}
th {{
    background: #2b6cb0;
    color: #fff;
    padding: 8px 10px;
    text-align: left;
    font-weight: 600;
    white-space: nowrap;
}}
td {{
    padding: 6px 10px;
    border-bottom: 1px solid #e2e8f0;
}}
tr:nth-child(even) td {{
    background: #f7fafc;
}}
tr:hover td {{
    background: #ebf8ff;
}}
.alert-danger {{
    background: #fff5f5;
    border-left: 4px solid #e53e3e;
    padding: 10px 16px;
    margin: 8px 0;
    border-radius: 0 6px 6px 0;
    font-size: 14px;
    color: #c53030;
}}
.alert-warning {{
    background: #fffaf0;
    border-left: 4px solid #dd6b20;
    padding: 10px 16px;
    margin: 8px 0;
    border-radius: 0 6px 6px 0;
    font-size: 14px;
    color: #c05621;
}}
.footer {{
    text-align: center;
    margin-top: 40px;
    padding-top: 20px;
    border-top: 1px solid #e2e8f0;
    color: #a0aec0;
    font-size: 12px;
}}
</style>
</head>
<body>
<div class="container">
{body_html}
<div class="footer">Generated by SAEPIN — AI Agent System</div>
</div>
</body>
</html>"""


def _esc(text):
    """HTMLエスケープ"""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _render_html_table(rows):
    """パイプ区切りテーブル行をHTMLテーブルに変換"""
    html = ['<table>']
    for i, row in enumerate(rows):
        cells = [c.strip() for c in row.strip("|").split("|")]
        if i == 0:
            html.append("<thead><tr>")
            for c in cells:
                html.append(f"<th>{_esc(c)}</th>")
            html.append("</tr></thead><tbody>")
        else:
            html.append("<tr>")
            for c in cells:
                html.append(f"<td>{_esc(c)}</td>")
            html.append("</tr>")
    html.append("</tbody></table>")
    return "\n".join(html)


def save_report(report, report_type="weekly"):
    """レポートをMD＋HTMLの両方で保存"""
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    start_str = f"{LAST_MONDAY.month}月{LAST_MONDAY.day}日"
    end_str = f"{LAST_SUNDAY.month}月{LAST_SUNDAY.day}日"
    filename = f"会員分析（日報）{start_str}-{end_str}"

    # MD保存
    md_path = REPORTS_DIR / f"{filename}.md"
    md_path.write_text(report, encoding="utf-8")
    print(f"💾 レポート保存(MD): {md_path}")

    # HTML保存
    html_path = REPORTS_DIR / f"{filename}.html"
    html_content = format_html_report(report)
    html_path.write_text(html_content, encoding="utf-8")
    print(f"💾 レポート保存(HTML): {html_path}")

    return md_path


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
