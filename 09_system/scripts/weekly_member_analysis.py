"""
週次会員分析レポート自動生成・Slack投稿スクリプト

毎週月曜7:00にWindowsタスクスケジューラで実行。
会員カルテ（Google Sheets）からデータを取得し、分析レポートを生成。
Slackチャンネルに投稿 + HTMLファイルをローカル保存。

使い方:
  python weekly_member_analysis.py                  # 先週分を分析・Slack投稿
  python weekly_member_analysis.py --dry-run        # Slack投稿なし（テスト用）
  python weekly_member_analysis.py --weeks-ago 2    # 2週前を分析
  python weekly_member_analysis.py --backfill       # 2/8〜直近全週のHTMLを一括生成
"""

import sys
import io
import json
import argparse
import re
import random
from pathlib import Path
from datetime import datetime, timedelta
from collections import defaultdict

# Windows環境のUTF-8対応
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# ==================== 設定 ====================
BASE_DIR = Path(__file__).parent.parent.parent
CONFIG_DIR = BASE_DIR / "09_system" / "config"
REPORT_DIR = BASE_DIR / "03_clients" / "SIFTAI" / "プロプレミアムTEAM" / "会員管理" / "週次レポート"
SA_FILE = CONFIG_DIR / "google_service_account.json"
OAUTH_CLIENT_FILE = CONFIG_DIR / "google_oauth_client.json"
OAUTH_TOKEN_FILE = CONFIG_DIR / "google_oauth_token.pickle"
SLACK_TOKEN_FILE = CONFIG_DIR / "slack_user_token.txt"  # 冴香名義で投稿

SPREADSHEET_ID = "1EHKgmE1d7T5N9GbT_QJV72Q8Dh91iNysHHkz6LEWJy4"
SLACK_CHANNEL = "C0AC8404FPE"

# データ開始日
DATA_START = datetime(2026, 2, 2)  # 2/8を含む週の月曜日

# 感情分析キーワード
POSITIVE_WORDS = ['できた', '成功', '嬉しい', '発見', '達成', '納品', '受注', '学び', '進め', '完了', '理解', '実現', '構築', '獲得', '成果']
NEGATIVE_WORDS = ['難しい', '不安', '分からな', 'わからな', 'つまず', '悩', '課題', '苦戦', '止まっ', 'できない', 'できな', '迷', '全然']


# ==================== Google Sheets アクセス ====================
def get_access_token():
    """Google Sheets APIのアクセストークンを取得（SA優先、なければOAuth）"""
    from google.auth.transport.requests import Request

    # サービスアカウントがあればそちらを使う（Windows環境）
    if SA_FILE.exists():
        from google.oauth2.service_account import Credentials
        creds = Credentials.from_service_account_file(
            str(SA_FILE),
            scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
        )
        creds.refresh(Request())
        return creds.token

    # OAuthトークン（Mac環境）
    if OAUTH_TOKEN_FILE.exists():
        import pickle
        with open(OAUTH_TOKEN_FILE, 'rb') as f:
            creds = pickle.load(f)
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
            with open(OAUTH_TOKEN_FILE, 'wb') as f:
                pickle.dump(creds, f)
        return creds.token

    # どちらもない場合はOAuthクライアントから新規認証
    if OAUTH_CLIENT_FILE.exists():
        from google_auth_oauthlib.flow import InstalledAppFlow
        flow = InstalledAppFlow.from_client_secrets_file(
            str(OAUTH_CLIENT_FILE),
            scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
        )
        creds = flow.run_local_server(port=0)
        OAUTH_TOKEN_FILE.parent.mkdir(parents=True, exist_ok=True)
        import pickle
        with open(OAUTH_TOKEN_FILE, 'wb') as f:
            pickle.dump(creds, f)
        return creds.token

    raise FileNotFoundError("Google認証ファイルが見つかりません。SA keyまたはOAuthクライアントを配置してください。")


def fetch_sheet_data(token, sheet_name, range_suffix="A:Z"):
    """指定シートのデータをrequestsで取得"""
    import requests as req
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values/{sheet_name}!{range_suffix}"
    resp = req.get(url, headers={"Authorization": f"Bearer {token}"})
    resp.raise_for_status()
    rows = resp.json().get('values', [])
    if not rows:
        return [], []
    return rows[0], rows[1:]


def detect_concierge(all_text):
    """全面談テキストからコンシェルジュ名を検出"""
    if not all_text:
        return '未割当'
    for keyword, name in [('安部', '安部友博'), ('山本', '山本'), ('安永', '安永')]:
        if keyword in all_text:
            return name
    return '未割当'


# 手動マッピング（カルテに記載がない会員用）
CONCIERGE_OVERRIDE = {
    'P00026': '山本',      # 田村耕太郎 — 冴香確認済み
    'P00027': '安永',      # 石渡一雄 — 冴香確認済み
    # 'P00018': '',         # 河野徹 — 冴香に確認待ち
    # 'P00021': '',         # 天野雅喜 — 冴香に確認待ち
    # 'P00025': '',         # 鍋島佑太 — 冴香に確認待ち
    # 'P00030': '',         # 廣瀬亘 — 冴香に確認待ち
    # 'P00058': '',         # 南山絵里 — 冴香に確認待ち
}


def fetch_all_data(token):
    """全シートのデータを一括取得（全面談列含む）"""
    print("  データ取得中...")
    _, daily_rows = fetch_sheet_data(token, 'DB_Daily-Report', 'A:F')
    _, weekly_rows = fetch_sheet_data(token, 'DB_Weekly-Report', 'A:J')
    _, fb_rows = fetch_sheet_data(token, 'DB_FB-Report', 'A:L')
    # 全面談列（BR列まで）を取得してコンシェルジュ情報をスキャン
    _, pro_rows = fetch_sheet_data(token, 'DB_Pro-Members', 'A:BR')
    _, prem_rows = fetch_sheet_data(token, 'DB_Premium-Members', 'A:BR')
    print(f"  日報: {len(daily_rows)}件, 週報: {len(weekly_rows)}件, FB: {len(fb_rows)}件")
    print(f"  PRO会員: {len(pro_rows)}名, PREMIUM会員: {len(prem_rows)}名")

    # (id, name, concierge) のタプルで返す
    # 列K以降（インデックス10〜）= 面談関連データすべてをスキャン
    def build_members(rows):
        members = []
        for r in rows:
            if len(r) > 1 and r[1]:
                member_id = r[0]
                name = r[1]
                # 手動マッピングが優先
                if member_id in CONCIERGE_OVERRIDE:
                    concierge = CONCIERGE_OVERRIDE[member_id]
                else:
                    # 面談関連列（列K以降）を全結合してスキャン
                    all_interview_text = ' '.join(str(c) for c in r[10:])
                    concierge = detect_concierge(all_interview_text)
                members.append((member_id, name, concierge))
        return members

    # 氏名列の全角・半角スペースを除去（表記揺れ対策）
    def clean_name_col(rows, name_col):
        for r in rows:
            if len(r) > name_col and r[name_col]:
                r[name_col] = r[name_col].strip().replace('\u3000', '')
        return rows

    daily_rows = clean_name_col(daily_rows, 1)
    weekly_rows = clean_name_col(weekly_rows, 1)
    fb_rows = clean_name_col(fb_rows, 2)  # FB会はcol2が氏名

    return {
        'daily': daily_rows,
        'weekly': weekly_rows,
        'fb': fb_rows,
        'pro_members': build_members(pro_rows),
        'premium_members': build_members(prem_rows),
    }


# ==================== ユーティリティ ====================
def parse_date(text):
    """日付テキストをdatetimeに変換"""
    if not text:
        return None
    for fmt in ('%Y/%m/%d', '%Y-%m-%d', '%Y/%m/%d %H:%M:%S'):
        try:
            return datetime.strptime(text.strip(), fmt)
        except ValueError:
            continue
    return None


def filter_by_date(rows, date_col_idx, start_date, end_date):
    """日付列でフィルタ"""
    filtered = []
    for row in rows:
        if len(row) > date_col_idx and row[date_col_idx]:
            d = parse_date(row[date_col_idx])
            if d and start_date <= d <= end_date:
                filtered.append(row)
    return filtered


def parse_hours(text):
    """稼働時間テキストをfloatに変換"""
    if not text:
        return 0.0
    match = re.match(r'(\d+):(\d+)', text)
    if match:
        return int(match.group(1)) + int(match.group(2)) / 60
    match = re.match(r'([\d.]+)', text.replace('約', '').replace('時間', '').replace('h', '').strip())
    if match:
        return float(match.group(1))
    return 0.0


def analyze_sentiment(text):
    """テキストの感情傾向を判定"""
    if not text:
        return 'neutral'
    positive = any(w in text for w in POSITIVE_WORDS)
    negative = any(w in text for w in NEGATIVE_WORDS)
    if positive and not negative:
        return 'positive'
    elif negative and not positive:
        return 'negative'
    elif positive and negative:
        return 'mixed'
    return 'neutral'


def extract_monetize_amount(text):
    """マネタイズテキストから金額を抽出"""
    if not text:
        return 0
    if re.match(r'^(0円?|なし|特になし|マネタイズなし|収益なし|直接収益なし|ー|−|-)$', text.strip()):
        return 0
    match = re.search(r'[¥￥]([\d,]+)', text)
    if match:
        return int(match.group(1).replace(',', ''))
    match = re.search(r'([\d,]+)\s*円', text)
    if match:
        return int(match.group(1).replace(',', ''))
    return 0


def extract_positive_achievements(texts, max_count=5):
    """日報テキストからポジティブな成果を最大max_count件抽出"""
    achievements = []
    for text in texts:
        if not text or not text.strip():
            continue
        s = analyze_sentiment(text)
        if s in ('positive', 'mixed'):
            # 短くカットして成果として追加
            short = text.strip()[:120]
            if short:
                achievements.append(short)
    return achievements[:max_count]


def jp_date(d):
    """datetimeを '3月16日' 形式に変換"""
    return f"{d.month}月{d.day}日"


def get_weeks(start_date, end_date):
    """指定期間の全週（月曜〜日曜）リストを返す"""
    weeks = []
    # start_dateを含む週の月曜日から開始
    monday = start_date - timedelta(days=start_date.weekday())
    while monday <= end_date:
        sunday = monday + timedelta(days=6)
        weeks.append((monday, sunday))
        monday += timedelta(weeks=1)
    return weeks


# ==================== データ分析 ====================
def analyze_week(all_data, start_date, end_date, member_list, plan_type):
    """1週間分のデータを分析して構造化データを返す

    member_list: [(id, name, concierge), ...] のリスト
    """
    member_names = [m[1] for m in member_list]
    member_set = set(member_names)

    # コンシェルジュマップ: {名前: コンシェルジュ名}
    concierge_map = {}
    member_id_map = {}  # {名前: ID}
    for mid, name, con in member_list:
        concierge_map[name] = con
        member_id_map[name] = mid

    # 日報フィルタ（日付 + メンバー）
    daily_filtered = filter_by_date(all_data['daily'], 2, start_date, end_date)
    daily = [r for r in daily_filtered if len(r) > 1 and r[1] in member_set]

    # 週報フィルタ
    weekly_filtered = filter_by_date(all_data['weekly'], 2, start_date, end_date)
    weekly = [r for r in weekly_filtered if len(r) > 1 and r[1] in member_set]

    # FB会フィルタ
    fb_filtered = filter_by_date(all_data['fb'], 1, start_date, end_date)
    fb = [r for r in fb_filtered if len(r) > 2 and r[2] in member_set]

    # ① 基本統計
    member_daily = defaultdict(list)
    for row in daily:
        name = row[1] if len(row) > 1 else 'unknown'
        hours = parse_hours(row[3] if len(row) > 3 else '')
        text = row[4] if len(row) > 4 else ''
        member_daily[name].append({'hours': hours, 'text': text})

    total_hours = sum(h['hours'] for posts in member_daily.values() for h in posts)
    submitters = len(member_daily)
    avg_hours = total_hours / submitters if submitters else 0
    total_submissions = len(daily)
    two_plus = sum(1 for posts in member_daily.values() if len(posts) >= 2)

    hours_ranking = sorted(
        [(name, sum(p['hours'] for p in posts), len(posts)) for name, posts in member_daily.items()],
        key=lambda x: -x[1]
    )
    count_ranking = sorted(
        [(name, len(posts)) for name, posts in member_daily.items()],
        key=lambda x: -x[1]
    )

    # ② FB会
    fb_members = defaultdict(int)
    fb_by_day = defaultdict(int)
    for f in fb:
        name = f[2] if len(f) > 2 else '?'
        fb_members[name] += 1
        d = parse_date(f[1] if len(f) > 1 else '')
        if d:
            fb_by_day[d.weekday()] += 1
    fb_ranking = sorted(fb_members.items(), key=lambda x: -x[1])

    # ③ 感情分析
    sentiments = {'positive': set(), 'negative': set(), 'mixed': set(), 'neutral': set()}
    positive_voices = []
    negative_voices = []
    for row in daily:
        if len(row) > 4:
            name = row[1]
            text = row[4]
            s = analyze_sentiment(text)
            sentiments[s].add(name)
            if s in ('positive', 'mixed') and text.strip():
                positive_voices.append((name, text))
            if s in ('negative', 'mixed') and text.strip():
                negative_voices.append((name, text))

    active_count = len(sentiments['positive'] | sentiments['mixed'])
    stagnant_count = len(sentiments['negative'])
    total_sentiment = active_count + stagnant_count
    active_rate = int(active_count / total_sentiment * 100) if total_sentiment else 0
    stagnant_rate = 100 - active_rate if total_sentiment else 0

    # ⑤ マネタイズフェーズ
    phases = defaultdict(list)
    member_phase = {}  # メンバーごとのフェーズ
    for w in weekly:
        name = w[1] if len(w) > 1 else '?'
        phase = w[5] if len(w) > 5 else '未報告'
        if name not in phases[phase]:
            phases[phase].append(name)
        member_phase[name] = phase

    # ⑥ マネタイズ成果
    monetize = []
    member_monetize = {}  # メンバーごとのマネタイズ額
    for w in weekly:
        name = w[1] if len(w) > 1 else '?'
        money_text = w[9] if len(w) > 9 else '0'
        amount = extract_monetize_amount(money_text)
        member_monetize[name] = amount
        if amount > 0:
            monetize.append((name, amount, money_text))
    monetize.sort(key=lambda x: -x[1])
    total_money = sum(m[1] for m in monetize)
    avg_money = total_money // len(monetize) if monetize else 0

    # ⑨ 非アクティブ
    daily_submitters = set(member_daily.keys())
    fb_submitter_set = set(fb_members.keys())
    active_set = daily_submitters | fb_submitter_set
    inactive = [m for m in member_names if m not in active_set]

    # 声サンプリング
    random.seed(int(start_date.timestamp()))
    sampled_pos = random.sample(positive_voices, min(5, len(positive_voices))) if positive_voices else []
    sampled_neg = random.sample(negative_voices, min(3, len(negative_voices))) if negative_voices else []

    # 個別メンバーKPI
    member_kpi = {}
    for name in member_names:
        daily_posts = member_daily.get(name, [])
        hours = sum(p['hours'] for p in daily_posts)
        daily_count = len(daily_posts)
        fb_count = fb_members.get(name, 0)
        mon = member_monetize.get(name, 0)
        phase = member_phase.get(name, '未報告')
        # 日報テキストからポジティブな成果を抽出
        texts = [p['text'] for p in daily_posts]
        achievements = extract_positive_achievements(texts, max_count=5)
        member_kpi[name] = {
            'hours': hours,
            'daily_count': daily_count,
            'fb_count': fb_count,
            'monetize': mon,
            'phase': phase,
            'achievements': achievements,
            'id': member_id_map.get(name, ''),
        }

    return {
        'plan_type': plan_type,
        'start_date': start_date,
        'end_date': end_date,
        'total_members': len(member_names),
        'total_hours': total_hours,
        'avg_hours': avg_hours,
        'submitters': submitters,
        'total_submissions': total_submissions,
        'two_plus': two_plus,
        'submission_rate': int(two_plus / len(member_names) * 100) if member_names else 0,
        'hours_ranking': hours_ranking,
        'count_ranking': count_ranking,
        'fb_total': len(fb),
        'fb_participants': len(fb_members),
        'fb_ranking': fb_ranking,
        'fb_by_day': fb_by_day,
        'active_rate': active_rate,
        'stagnant_rate': stagnant_rate,
        'phases': dict(phases),
        'monetize': monetize,
        'total_money': total_money,
        'avg_money': avg_money,
        'weekly_submitters': len(weekly),
        'weekly_rate': int(len(weekly) / len(member_names) * 100) if member_names else 0,
        'inactive': inactive,
        'inactive_count': len(inactive),
        'active_set': active_set,
        'positive_voices': sampled_pos,
        'negative_voices': sampled_neg,
        'member_daily': dict(member_daily),
        'fb_members': dict(fb_members),
        'member_kpi': member_kpi,
        'concierge_map': concierge_map,
    }


def merge_results(pro, prem):
    """PRO結果とPREMIUM結果を単純合算してALL結果を作る"""
    # ランキング系: 両方のリストを結合して再ソート（併用メンバーは重複排除、PREMIUM側を優先）
    hours_map = {}
    for name, hours, count in pro['hours_ranking']:
        hours_map[name] = (hours, count)
    for name, hours, count in prem['hours_ranking']:
        hours_map[name] = (hours, count)  # PREMIUM側で上書き（同じデータなので重複排除）
    hours_ranking = sorted([(n, h, c) for n, (h, c) in hours_map.items()], key=lambda x: -x[1])

    count_map = {}
    for name, count in pro['count_ranking']:
        count_map[name] = count
    for name, count in prem['count_ranking']:
        count_map[name] = count  # 重複排除
    count_ranking = sorted(count_map.items(), key=lambda x: -x[1])

    # FB: 合算（重複排除）
    fb_map = {}
    for name, count in pro['fb_ranking']:
        fb_map[name] = count
    for name, count in prem['fb_ranking']:
        fb_map[name] = count  # 重複排除
    fb_ranking = sorted(fb_map.items(), key=lambda x: -x[1])

    fb_by_day = defaultdict(int)
    for d_idx in range(7):
        fb_by_day[d_idx] = pro['fb_by_day'].get(d_idx, 0) + prem['fb_by_day'].get(d_idx, 0)

    # 感情: 加重平均
    total_sub = pro['submitters'] + prem['submitters']
    active_rate = int((pro['active_rate'] * pro['submitters'] + prem['active_rate'] * prem['submitters']) / total_sub) if total_sub else 0
    stagnant_rate = 100 - active_rate

    # フェーズ: マージ
    phases = {}
    for phase, members in pro['phases'].items():
        phases.setdefault(phase, []).extend(members)
    for phase, members in prem['phases'].items():
        phases.setdefault(phase, []).extend(members)

    # マネタイズ: 結合して再ソート（併用メンバーは重複排除）
    money_map = {}
    text_map = {}
    for name, amount, text in pro['monetize']:
        money_map[name] = amount
        text_map[name] = text
    for name, amount, text in prem['monetize']:
        money_map[name] = amount  # 重複排除
        text_map[name] = text
    monetize = sorted([(n, a, text_map[n]) for n, a in money_map.items()], key=lambda x: -x[1])
    total_money = sum(a for _, a, _ in monetize)
    avg_money = total_money // len(monetize) if monetize else 0

    # 非アクティブ: 結合（重複除外）
    inactive = list(set(pro['inactive'] + prem['inactive']))
    active_set = pro['active_set'] | prem['active_set']

    # 声: 結合（重複排除）
    seen_pos = set()
    positive_voices = []
    for name, text in pro['positive_voices'] + prem['positive_voices']:
        if (name, text) not in seen_pos:
            seen_pos.add((name, text))
            positive_voices.append((name, text))
    seen_neg = set()
    negative_voices = []
    for name, text in pro['negative_voices'] + prem['negative_voices']:
        if (name, text) not in seen_neg:
            seen_neg.add((name, text))
            negative_voices.append((name, text))

    # member_daily / fb_members: マージ（併用メンバーは重複排除、PREMIUM優先）
    member_daily = dict(pro['member_daily'])
    for name, posts in prem['member_daily'].items():
        member_daily[name] = posts  # 重複排除（同じデータなので上書きでOK）
    fb_members = dict(pro['fb_members'])
    for name, count in prem['fb_members'].items():
        fb_members[name] = count  # 重複排除

    # member_kpi: マージ（併用メンバーは重複排除、PREMIUM優先）
    member_kpi = dict(pro['member_kpi'])
    for name, kpi in prem['member_kpi'].items():
        member_kpi[name] = kpi  # 重複排除（同じ日報データなので上書きでOK）

    # concierge_map: マージ
    concierge_map = dict(pro['concierge_map'])
    concierge_map.update(prem['concierge_map'])

    # 併用メンバーの重複を除いた集計
    # member_dailyから正しい値を再計算
    total_members = len(set(list(pro['concierge_map'].keys()) + list(prem['concierge_map'].keys())))
    total_hours = sum(sum(p['hours'] for p in posts) for posts in member_daily.values())
    submitters = len(member_daily)
    total_submissions = sum(len(posts) for posts in member_daily.values())
    two_plus = sum(1 for posts in member_daily.values() if len(posts) >= 2)
    # weekly_submittersも重複排除
    pro_weekly_names = set(pro.get('weekly_submitter_names', []))
    prem_weekly_names = set(prem.get('weekly_submitter_names', []))
    weekly_submitters = len(pro_weekly_names | prem_weekly_names) if pro_weekly_names or prem_weekly_names else pro['weekly_submitters'] + prem['weekly_submitters']

    return {
        'plan_type': 'ALL',
        'start_date': pro['start_date'],
        'end_date': pro['end_date'],
        'total_members': total_members,
        'total_hours': total_hours,
        'avg_hours': total_hours / submitters if submitters else 0,
        'submitters': submitters,
        'total_submissions': total_submissions,
        'two_plus': two_plus,
        'submission_rate': int(two_plus / total_members * 100) if total_members else 0,
        'hours_ranking': hours_ranking,
        'count_ranking': count_ranking,
        'fb_total': sum(fb_members.values()),
        'fb_participants': len(fb_members),
        'fb_ranking': fb_ranking,
        'fb_by_day': dict(fb_by_day),
        'active_rate': active_rate,
        'stagnant_rate': stagnant_rate,
        'phases': phases,
        'monetize': monetize,
        'total_money': total_money,
        'avg_money': avg_money,
        'weekly_submitters': weekly_submitters,
        'weekly_rate': int(weekly_submitters / total_members * 100) if total_members else 0,
        'inactive': inactive,
        'inactive_count': len(inactive),
        'active_set': active_set,
        'positive_voices': positive_voices,
        'negative_voices': negative_voices,
        'member_daily': member_daily,
        'fb_members': fb_members,
        'member_kpi': member_kpi,
        'concierge_map': concierge_map,
    }


# ==================== HTML生成 ====================
CSS = """
:root {
  --primary: #1e3a5f; --accent: #2b6cb0; --accent-light: #ebf4ff;
  --success: #16a34a; --warning: #d97706; --danger: #dc2626;
  --gray-50: #f8fafc; --gray-100: #f1f5f9; --gray-200: #e2e8f0;
  --gray-300: #cbd5e1; --gray-400: #94a3b8; --gray-600: #475569; --gray-800: #1e293b;
  --pro-color: #2b6cb0; --premium-color: #7c3aed;
}
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'Hiragino Kaku Gothic ProN','Yu Gothic UI','Meiryo',sans-serif; background:var(--gray-100); color:var(--gray-800); line-height:1.7; }
.header { color:#fff; padding:32px 40px 24px; }
.header.pro { background:linear-gradient(135deg,#1e3a5f,#2b6cb0); }
.header.premium { background:linear-gradient(135deg,#4c1d95,#7c3aed); }
.header.all { background:linear-gradient(135deg,#1e3a5f,#7c3aed); }
.header-inner { max-width:1100px; margin:0 auto; }
.header h1 { font-size:26px; font-weight:700; }
.header .subtitle { font-size:14px; opacity:0.85; margin-top:4px; }
.nav-bar { background:#fff; border-bottom:1px solid var(--gray-200); position:sticky; top:0; z-index:10; }
.nav-inner { max-width:1100px; margin:0 auto; display:flex; flex-wrap:wrap; }
.plan-tabs { display:flex; border-right:1px solid var(--gray-200); }
.plan-tab { padding:12px 20px; font-size:13px; font-weight:600; color:var(--gray-400); cursor:pointer; border-bottom:3px solid transparent; text-decoration:none; }
.plan-tab:hover { color:var(--accent); }
.plan-tab.active-pro { color:var(--pro-color); border-bottom-color:var(--pro-color); }
.plan-tab.active-premium { color:var(--premium-color); border-bottom-color:var(--premium-color); }
.plan-tab.active-all { color:#4f46e5; border-bottom-color:#4f46e5; }
.view-tabs { display:flex; border-right:1px solid var(--gray-200); }
.view-tab { padding:12px 20px; font-size:13px; font-weight:600; color:var(--gray-400); cursor:pointer; border:none; background:none; border-bottom:3px solid transparent; }
.view-tab.active { color:#059669; border-bottom-color:#059669; }
.month-tabs { display:flex; border-right:1px solid var(--gray-200); }
.month-tab { padding:12px 20px; font-size:13px; font-weight:600; color:var(--gray-400); cursor:pointer; border-bottom:3px solid transparent; }
.month-tab:hover { color:var(--accent); }
.month-tab.active { color:var(--accent); border-bottom-color:var(--accent); }
.week-tabs { display:flex; flex:1; overflow-x:auto; }
.week-tab { padding:12px 18px; font-size:13px; color:var(--gray-400); cursor:pointer; border-bottom:3px solid transparent; white-space:nowrap; text-decoration:none; }
.week-tab:hover { color:var(--accent); }
.week-tab.active { color:var(--accent); font-weight:600; border-bottom-color:var(--accent); }
.week-tab.disabled { color:#cbd5e1; cursor:default; }
.nav-top-btn { padding:12px 16px; font-size:12px; font-weight:700; color:var(--accent); cursor:pointer; border-left:1px solid var(--gray-200); white-space:nowrap; display:flex; align-items:center; }
.nav-top-btn:hover { background:var(--accent-light); }
.main { max-width:1100px; margin:0 auto; padding:28px 20px 60px; }
.kpi-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:16px; margin-bottom:32px; }
.kpi-card { background:#fff; border-radius:10px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.06); }
.kpi-label { font-size:12px; color:var(--gray-400); font-weight:600; text-transform:uppercase; letter-spacing:0.5px; }
.kpi-value { font-size:32px; font-weight:700; color:var(--primary); margin:4px 0; }
.kpi-change { font-size:13px; font-weight:600; }
.kpi-change.up { color:var(--success); }
.kpi-change.down { color:var(--danger); }
.kpi-sub { font-size:12px; color:var(--gray-400); margin-top:2px; }
.section { background:#fff; border-radius:10px; box-shadow:0 1px 3px rgba(0,0,0,0.06); margin-bottom:24px; overflow:hidden; }
.section-header { display:flex; align-items:center; gap:10px; padding:16px 24px; border-bottom:1px solid var(--gray-200); cursor:pointer; user-select:none; }
.section-header:hover { background:var(--gray-50); }
.section-num { display:inline-flex; align-items:center; justify-content:center; width:28px; height:28px; background:var(--accent); color:#fff; border-radius:50%; font-size:13px; font-weight:700; flex-shrink:0; }
.section-num.premium { background:var(--premium-color); }
.section-num.all { background:#4f46e5; }
.section-title { font-size:16px; font-weight:700; color:var(--gray-800); flex:1; }
.section-toggle { font-size:18px; color:var(--gray-400); transition:transform 0.2s; }
.section.collapsed .section-toggle { transform:rotate(-90deg); }
.section.collapsed .section-body { display:none; }
.section-body { padding:20px 24px; }
.top3 { display:flex; gap:12px; flex-wrap:wrap; margin:12px 0; }
.top3-item { display:flex; align-items:center; gap:8px; background:var(--gray-50); border-radius:8px; padding:10px 16px; flex:1; min-width:200px; }
.rank { font-size:20px; width:32px; text-align:center; }
.top3-name { font-weight:700; font-size:14px; }
.top3-detail { font-size:12px; color:var(--gray-600); }
.sentiment-gauge { display:flex; height:36px; border-radius:8px; overflow:hidden; margin:16px 0 8px; }
.sentiment-pos { background:linear-gradient(90deg,#22c55e,#4ade80); display:flex; align-items:center; justify-content:center; color:#fff; font-weight:700; font-size:14px; }
.sentiment-neg { background:linear-gradient(90deg,#f87171,#ef4444); display:flex; align-items:center; justify-content:center; color:#fff; font-weight:700; font-size:14px; }
.phase-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:12px; margin:12px 0; }
.phase-card { border:1px solid var(--gray-200); border-radius:8px; padding:14px 16px; position:relative; }
.phase-card::before { content:''; position:absolute; left:0; top:0; bottom:0; width:4px; border-radius:8px 0 0 8px; }
.phase-card.p1::before { background:#3b82f6; } .phase-card.p2::before { background:#f59e0b; }
.phase-card.p3::before { background:#8b5cf6; } .phase-card.p4::before { background:#22c55e; }
.phase-card.p5::before { background:#ef4444; }
.monetize-podium { display:flex; gap:12px; justify-content:center; align-items:flex-end; margin:20px 0; }
.podium-item { text-align:center; padding:16px 20px; border-radius:10px; min-width:140px; }
.podium-item.gold { background:linear-gradient(135deg,#fef3c7,#fde68a); order:1; }
.podium-item.silver { background:linear-gradient(135deg,#f1f5f9,#e2e8f0); order:0; }
.podium-item.bronze { background:linear-gradient(135deg,#fed7aa,#fdba74); order:2; }
.podium-medal { font-size:28px; } .podium-name { font-weight:700; font-size:15px; margin:4px 0; }
.podium-amount { font-size:20px; font-weight:800; color:var(--primary); }
.voice-grid { display:grid; grid-template-columns:1fr 1fr; gap:16px; margin:12px 0; }
@media (max-width:768px) { .voice-grid { grid-template-columns:1fr; } }
.voice-section-label { font-size:13px; font-weight:700; margin-bottom:8px; padding:4px 10px; border-radius:4px; display:inline-block; }
.voice-section-label.pos { background:#dcfce7; color:#166534; }
.voice-section-label.neg { background:#fef2f2; color:#991b1b; }
.voice-card { border-left:3px solid var(--gray-300); padding:10px 14px; margin:8px 0; background:var(--gray-50); border-radius:0 6px 6px 0; font-size:13px; line-height:1.6; }
.voice-card.pos { border-left-color:var(--success); } .voice-card.neg { border-left-color:var(--danger); }
.voice-name { font-weight:700; font-size:13px; margin-bottom:4px; }
.proposal-list { margin:12px 0; }
.proposal-item { display:flex; align-items:flex-start; gap:10px; padding:10px 0; border-bottom:1px solid var(--gray-100); font-size:14px; }
.proposal-icon { flex-shrink:0; width:24px; height:24px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:12px; margin-top:2px; }
.proposal-icon.follow { background:#dbeafe; color:#2563eb; } .proposal-icon.improve { background:#fef3c7; color:#d97706; }
table { width:100%; border-collapse:collapse; font-size:13px; margin:12px 0; }
th { background:var(--primary); color:#fff; padding:10px 12px; text-align:left; font-weight:600; font-size:12px; white-space:nowrap; }
th.premium-th { background:var(--premium-color); }
th.all-th { background:#4f46e5; }
td { padding:8px 12px; border-bottom:1px solid var(--gray-200); }
tr:nth-child(even) td { background:var(--gray-50); } tr:hover td { background:var(--accent-light); }
.zero { color:var(--danger); font-weight:600; } .low { color:var(--warning); }
.alert { display:flex; align-items:flex-start; gap:10px; padding:12px 16px; border-radius:8px; margin:8px 0; font-size:13px; line-height:1.6; }
.alert-icon { font-size:18px; flex-shrink:0; margin-top:1px; }
.alert.danger { background:#fef2f2; border:1px solid #fecaca; color:#991b1b; }
.alert.warning { background:#fffbeb; border:1px solid #fde68a; color:#92400e; }
.day-chart { display:flex; align-items:flex-end; gap:8px; height:140px; margin:16px 0 8px; padding:0 20px; }
.day-col { flex:1; display:flex; flex-direction:column; align-items:center; gap:4px; }
.day-bar { width:100%; max-width:50px; border-radius:4px 4px 0 0; display:flex; align-items:flex-start; justify-content:center; padding-top:4px; font-size:12px; font-weight:700; color:#fff; }
.day-bar.pro-bar { background:linear-gradient(180deg,#3b82f6,#2563eb); }
.day-bar.premium-bar { background:linear-gradient(180deg,#a78bfa,#7c3aed); }
.day-bar.all-bar { background:linear-gradient(180deg,#6366f1,#4f46e5); }
.day-label { font-size:12px; color:var(--gray-600); font-weight:600; }
.footer { text-align:center; padding:24px; color:var(--gray-400); font-size:12px; }
.float-top { position:fixed; bottom:24px; right:24px; width:48px; height:48px; background:var(--accent); color:#fff; border:none; border-radius:50%; font-size:20px; cursor:pointer; box-shadow:0 4px 12px rgba(0,0,0,0.2); display:none; align-items:center; justify-content:center; z-index:100; }
.float-top:hover { transform:scale(1.1); background:var(--primary); }
.no-data { text-align:center; padding:40px; color:var(--gray-400); font-size:15px; }
/* 成果報告ビュー用CSS */
.con-tab { padding:10px 16px; font-size:13px; font-weight:600; color:var(--gray-400); cursor:pointer; border:1px solid var(--gray-200); background:#fff; border-radius:8px 8px 0 0; }
.con-tab.active { color:var(--accent); background:var(--accent-light); border-bottom-color:transparent; }
.concierge-tabs { display:flex; gap:4px; margin-bottom:-1px; padding:0 24px; }
.member-card { border:1px solid var(--gray-200); border-radius:12px; padding:24px; margin-bottom:20px; }
.member-header h3 { font-size:18px; font-weight:700; }
.member-meta { font-size:12px; color:var(--gray-400); margin-top:4px; }
.member-kpi-row { display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin:16px 0; }
.member-kpi { text-align:center; padding:12px; background:var(--gray-50); border-radius:8px; }
.member-kpi .value { font-size:1.4rem; font-weight:800; color:var(--primary); }
.member-kpi .label { font-size:11px; color:var(--gray-400); margin-top:2px; }
.member-achievements { margin-top:12px; }
.member-achievements h4 { font-size:14px; font-weight:700; margin-bottom:8px; }
.member-achievements ul { list-style:none; padding:0; }
.member-achievements li { padding:6px 0; border-bottom:1px solid var(--gray-100); font-size:13px; line-height:1.6; }
.member-achievements li:before { content:'\\2022'; color:var(--accent); margin-right:8px; }
.bar-chart-h { margin:12px 0; }
.bar-row-h { display:flex; align-items:center; gap:10px; margin:6px 0; }
.bar-label-h { width:100px; font-size:13px; text-align:right; flex-shrink:0; }
.bar-track-h { flex:1; height:24px; background:var(--gray-100); border-radius:4px; overflow:hidden; }
.bar-fill-h { height:100%; border-radius:4px; display:flex; align-items:center; padding-left:8px; font-size:12px; font-weight:600; color:#fff; }
.bar-fill-pro { background:linear-gradient(90deg,#3b82f6,#2563eb); }
.bar-fill-premium { background:linear-gradient(90deg,#a78bfa,#7c3aed); }
.bar-fill-all { background:linear-gradient(90deg,#6366f1,#4f46e5); }
@media print { .nav-bar { position:static; } .section.collapsed .section-body { display:block; } body { background:#fff; } }
"""


def build_nav_html(plan_type, weeks, current_idx):
    """ナビゲーションバーのHTMLを生成（ビュー切り替えボタン付き）"""

    # プラン切り替えタブ（PRO / PREMIUM / ALL）
    plan_tabs = '<div class="plan-tabs">'
    for p in ['PRO', 'PREMIUM', 'ALL']:
        label = p if p != 'ALL' else '合算'
        if p == plan_type:
            cls = f'active-{p.lower()}'
            plan_tabs += f'<div class="plan-tab {cls}">{label}</div>'
        else:
            fname = get_filename(p, weeks[current_idx][0], weeks[current_idx][1])
            plan_tabs += f'<a class="plan-tab" href="{fname}">{label}</a>'
    plan_tabs += '</div>'

    # ビュー切り替えタブ（日報分析 / 成果報告）
    view_tabs = '''<div class="view-tabs">
  <button class="view-tab active" onclick="switchView('daily')">日報分析</button>
  <button class="view-tab" onclick="switchView('report')">成果報告</button>
</div>'''

    # 月タブ
    months = sorted(set(w[0].month for w, _ in [(w, None) for w in weeks]))
    month_names = {2: '2月', 3: '3月', 4: '4月'}
    current_month = weeks[current_idx][0].month
    month_tabs = '<div class="month-tabs">'
    for m in months:
        cls = 'active' if m == current_month else ''
        # 月の最初の週へリンク
        first_week_idx = next(i for i, w in enumerate(weeks) if w[0].month == m)
        if m == current_month:
            month_tabs += f'<div class="month-tab {cls}">{month_names.get(m, f"{m}月")}</div>'
        else:
            fname = get_filename(plan_type, weeks[first_week_idx][0], weeks[first_week_idx][1])
            month_tabs += f'<a class="month-tab" href="{fname}" style="text-decoration:none;">{month_names.get(m, f"{m}月")}</a>'
    month_tabs += '</div>'

    # 週タブ
    week_tabs = '<div class="week-tabs">'
    month_week_counter = defaultdict(int)
    for i, (start, end) in enumerate(weeks):
        m = start.month
        month_week_counter[m] += 1
        wn = month_week_counter[m]
        label = f"W{wn}: {start.month}/{start.day}\u2013{end.month}/{end.day}"
        if start.month == current_month:
            if i == current_idx:
                week_tabs += f'<div class="week-tab active">{label}</div>'
            else:
                fname = get_filename(plan_type, start, end)
                week_tabs += f'<a class="week-tab" href="{fname}" style="text-decoration:none;">{label}</a>'
        # 他の月の週は表示しない（月タブで切り替え）
    week_tabs += '</div>'

    return f'''<div class="nav-bar"><div class="nav-inner">
  {plan_tabs}{view_tabs}{month_tabs}{week_tabs}
  <div class="nav-top-btn" onclick="scrollTo({{top:0,behavior:'smooth'}})" title="ページトップへ">&#9650; TOP</div>
</div></div>'''


def build_kpi_html(data, prev_data):
    """KPIグリッドのHTMLを生成"""
    def change(cur, prev, fmt='pct', suffix=''):
        if prev is None or prev == 0:
            return ''
        if fmt == 'pct':
            diff = int((cur - prev) / prev * 100) if prev else 0
            cls = 'up' if diff >= 0 else 'down'
            sign = '+' if diff >= 0 else ''
            return f'<div class="kpi-change {cls}">{sign}{diff}% (前週 {int(prev)}{suffix})</div>'
        elif fmt == 'pt':
            diff = cur - prev
            cls = 'up' if diff >= 0 else 'down'
            sign = '+' if diff >= 0 else ''
            return f'<div class="kpi-change {cls}">{sign}{diff}pt (前週 {prev}%)</div>'
        elif fmt == 'count':
            if cur > prev:
                return f'<div class="kpi-change down">悪化 (前週 {prev}名)</div>'
            elif cur < prev:
                return f'<div class="kpi-change up">改善 (前週 {prev}名)</div>'
            else:
                return f'<div class="kpi-change">変動なし</div>'
        return ''

    p = prev_data
    d = data
    money_display = f"{int(d['total_money']//10000)}万円" if d['total_money'] >= 10000 else f"{int(d['total_money']):,}円"
    p_money = f"{int(p['total_money']//10000)}万円" if p and p['total_money'] >= 10000 else (f"{int(p['total_money']):,}円" if p else '')

    cards = f'''<div class="kpi-grid">
  <div class="kpi-card"><div class="kpi-label">総稼働時間</div><div class="kpi-value">{int(d['total_hours'])}h</div>{change(d['total_hours'], p['total_hours'], suffix='h') if p else ''}</div>
  <div class="kpi-card"><div class="kpi-label">平均稼働時間</div><div class="kpi-value">{int(d['avg_hours'])}h</div>{change(d['avg_hours'], p['avg_hours'], suffix='h') if p else ''}</div>
  <div class="kpi-card"><div class="kpi-label">日報提出者</div><div class="kpi-value">{d['submitters']}名</div><div class="kpi-sub">提出{d['total_submissions']}件</div></div>
  <div class="kpi-card"><div class="kpi-label">アクティブ傾向率</div><div class="kpi-value">{d['active_rate']}%</div>{change(d['active_rate'], p['active_rate'], fmt='pt') if p else ''}</div>
  <div class="kpi-card"><div class="kpi-label">マネタイズ合計</div><div class="kpi-value">{money_display}</div><div class="kpi-sub">収益報告 {len(d['monetize'])}名 / 週報 {d['weekly_submitters']}名</div></div>
  <div class="kpi-card"><div class="kpi-label">非アクティブ</div><div class="kpi-value">{d['inactive_count']}名</div>{change(d['inactive_count'], p['inactive_count'], fmt='count') if p else ''}<div class="kpi-sub">{d['plan_type']}全体 {d['total_members']}名</div></div>
</div>'''
    return cards


def build_section_html(num, title, body, plan_type='PRO'):
    """折りたたみセクションのラッパー"""
    num_cls = 'premium' if plan_type == 'PREMIUM' else 'all' if plan_type == 'ALL' else ''
    return f'''<div class="section"><div class="section-header" onclick="this.parentElement.classList.toggle('collapsed')">
  <span class="section-num {num_cls}">{num}</span><span class="section-title">{title}</span><span class="section-toggle">&#9660;</span>
</div><div class="section-body">{body}</div></div>'''


def build_top3_html(items, show_detail=True):
    """TOP3表示のHTML（同率は同じメダル、順位は1→2→3で進む）"""
    medals = ['\U0001f947', '\U0001f948', '\U0001f949']
    html = '<div class="top3">'

    # 比較値を取得
    def get_sort_val(item):
        if isinstance(item, tuple) and len(item) >= 2:
            return item[1]
        return 0

    # 同率は同じメダル、メダル種類が3つ使い切ったら終了
    medal_idx = 0  # 現在のメダル種類（0=金, 1=銀, 2=銅）
    prev_val = None
    for i, item in enumerate(items):
        val = get_sort_val(item)
        if prev_val is not None and val != prev_val:
            medal_idx += 1  # 値が変わったら次のメダルへ
        if medal_idx >= 3:
            break
        prev_val = val
        medal = medals[medal_idx]
        if isinstance(item, tuple) and len(item) >= 3:
            name, v, count = item[0], item[1], item[2]
            detail = f'{int(v)}h（平均 {int(v//count)}h/日）' if show_detail else f'{count}回'
        elif isinstance(item, tuple) and len(item) == 2:
            name, v = item
            detail = f'{v}回'
        else:
            continue
        html += f'<div class="top3-item"><span class="rank">{medal}</span><div><div class="top3-name">{name}</div><div class="top3-detail">{detail}</div></div></div>'

    html += '</div>'
    return html


def build_report_view_html(data):
    """成果報告ビューのHTMLを生成"""
    plan = data['plan_type']
    plan_lower = plan.lower()
    concierge_map = data.get('concierge_map', {})
    member_kpi = data.get('member_kpi', {})

    # ========== a. 成果サマリーKPI ==========
    money_display = f"{int(data['total_money']//10000)}万円" if data['total_money'] >= 10000 else f"{int(data['total_money']):,}円"
    reporters = len(data['monetize'])
    avg_money_display = f"{data['avg_money']:,}円" if data['avg_money'] else '0円'

    report_kpi = f'''<div class="kpi-grid">
  <div class="kpi-card"><div class="kpi-label">マネタイズ合計</div><div class="kpi-value">{money_display}</div></div>
  <div class="kpi-card"><div class="kpi-label">収益報告者数</div><div class="kpi-value">{reporters}名</div><div class="kpi-sub">全{data["total_members"]}名中</div></div>
  <div class="kpi-card"><div class="kpi-label">週報提出率</div><div class="kpi-value">{data["weekly_rate"]}%</div><div class="kpi-sub">{data["weekly_submitters"]}名提出</div></div>
  <div class="kpi-card"><div class="kpi-label">平均収益額</div><div class="kpi-value">{avg_money_display}</div><div class="kpi-sub">報告者平均</div></div>
</div>'''

    # ========== b. マネタイズ成果ランキング ==========
    ranking_html = ''

    # 表彰台（同率は同じメダル、メダル種類3つまで）
    if len(data['monetize']) > 0:
        medals_cls = ['gold', 'silver', 'bronze']
        medal_icons = ['\U0001f947', '\U0001f948', '\U0001f949']
        # 同率判定してメダルを割り当て
        podium_items = []  # (name, amount, medal_idx)
        medal_idx = 0
        prev_amount = None
        for name, amount, _ in data['monetize']:
            if prev_amount is not None and amount != prev_amount:
                medal_idx += 1
            if medal_idx >= 3:
                break
            podium_items.append((name, amount, medal_idx))
            prev_amount = amount
        # 表彰台レイアウト（金を中央に配置）
        if len(podium_items) >= 3:
            # 最初の銀・金・銅を表示（同率で4人以上の場合は全員表示）
            gold = [(n, a, mi) for n, a, mi in podium_items if mi == 0]
            silver = [(n, a, mi) for n, a, mi in podium_items if mi == 1]
            bronze = [(n, a, mi) for n, a, mi in podium_items if mi == 2]
            ranking_html += '<div class="monetize-podium">'
            for n, a, _ in silver:
                ranking_html += f'<div class="podium-item silver"><div class="podium-medal">{medal_icons[1]}</div><div class="podium-name">{n}</div><div class="podium-amount">{a:,}円</div></div>'
            for n, a, _ in gold:
                ranking_html += f'<div class="podium-item gold"><div class="podium-medal">{medal_icons[0]}</div><div class="podium-name">{n}</div><div class="podium-amount">{a:,}円</div></div>'
            for n, a, _ in bronze:
                ranking_html += f'<div class="podium-item bronze"><div class="podium-medal">{medal_icons[2]}</div><div class="podium-name">{n}</div><div class="podium-amount">{a:,}円</div></div>'
            ranking_html += '</div>'
        elif len(podium_items) > 0:
            ranking_html += '<div class="monetize-podium">'
            for n, a, mi in podium_items:
                ranking_html += f'<div class="podium-item {medals_cls[mi]}"><div class="podium-medal">{medal_icons[mi]}</div><div class="podium-name">{n}</div><div class="podium-amount">{a:,}円</div></div>'
            ranking_html += '</div>'

    # バーチャート（全員）
    if data['monetize']:
        max_amount = data['monetize'][0][1] if data['monetize'] else 1
        bar_cls = 'bar-fill-premium' if plan == 'PREMIUM' else 'bar-fill-all' if plan == 'ALL' else 'bar-fill-pro'
        ranking_html += '<div class="bar-chart-h">'
        for name, amount, _ in data['monetize']:
            width_pct = int(amount / max_amount * 100) if max_amount else 0
            ranking_html += f'''<div class="bar-row-h">
  <div class="bar-label-h">{name}</div>
  <div class="bar-track-h"><div class="bar-fill-h {bar_cls}" style="width:{width_pct}%;min-width:40px;">{amount:,}円</div></div>
</div>'''
        ranking_html += '</div>'

        # ランキングテーブル
        th_cls = 'premium-th' if plan == 'PREMIUM' else 'all-th' if plan == 'ALL' else ''
        ranking_html += f'<table><thead><tr><th class="{th_cls}">順位</th><th class="{th_cls}">会員名</th><th class="{th_cls}">金額</th><th class="{th_cls}">フェーズ</th></tr></thead><tbody>'
        for i, (name, amount, _) in enumerate(data['monetize'], start=1):
            phase = member_kpi.get(name, {}).get('phase', '未報告')
            ranking_html += f'<tr><td>{i}位</td><td>{name}</td><td>{amount:,}円</td><td>{phase}</td></tr>'
        ranking_html += '</tbody></table>'
    else:
        ranking_html += '<div class="no-data">今週の収益報告はありません</div>'

    ranking_section = build_section_html('R1', 'マネタイズ成果ランキング', ranking_html, plan)

    # ========== c. フェーズ別分布（パイプライン） ==========
    phase_colors = {0: 'p1', 1: 'p2', 2: 'p3', 3: 'p4', 4: 'p5'}
    phase_html = '<div class="phase-grid">'
    for i, (phase, members) in enumerate(sorted(data['phases'].items())):
        p_cls = phase_colors.get(i % 5, 'p1')
        phase_html += f'''<div class="phase-card {p_cls}">
  <div style="font-size:11px;font-weight:700;color:var(--gray-400);">PHASE {i+1}</div>
  <div style="font-size:14px;font-weight:700;margin:2px 0;">{phase}</div>
  <div style="font-size:24px;font-weight:800;color:var(--primary);">{len(members)}名</div>
  <div style="font-size:12px;color:var(--gray-600);margin-top:4px;">{", ".join(members)}</div>
</div>'''
    phase_html += '</div>'
    phase_html += f'<div style="font-size:13px;color:var(--gray-600);margin-top:8px;">週報提出者: {data["weekly_submitters"]}名 / {plan}基準{data["total_members"]}名（提出率 {data["weekly_rate"]}%）</div>'
    phase_section = build_section_html('R2', 'フェーズ別分布（パイプライン）', phase_html, plan)

    # ========== d. コンシェルジュ別タブ + e. メンバーカード ==========
    # コンシェルジュごとにメンバーをグループ化
    concierge_groups = defaultdict(list)
    for name, con in concierge_map.items():
        concierge_groups[con].append(name)

    # コンシェルジュタブのID・ラベルマッピング
    con_tab_info = []
    for con_name in sorted(concierge_groups.keys()):
        # IDを生成（安全な文字列に変換）
        if '安部' in con_name:
            con_id = 'abe'
            label = '安部チーム'
        elif '山本' in con_name:
            con_id = 'yamamoto'
            label = '山本チーム'
        elif '安永' in con_name:
            con_id = 'yasunaga'
            label = '安永チーム'
        else:
            con_id = 'unassigned'
            label = '未割当'
        con_tab_info.append((con_id, label, con_name))

    # タブボタン
    con_tabs_html = '<div class="concierge-tabs">'
    for idx, (con_id, label, _) in enumerate(con_tab_info):
        active = ' active' if idx == 0 else ''
        con_tabs_html += f'<button class="con-tab{active}" onclick="switchConcierge(\'{con_id}\')">{label}</button>'
    con_tabs_html += '</div>'

    # 各コンシェルジュパネル
    con_panels_html = ''
    for idx, (con_id, label, con_name) in enumerate(con_tab_info):
        display = 'block' if idx == 0 else 'none'
        members = concierge_groups[con_name]
        con_panels_html += f'<div class="concierge-panel" id="con-{con_id}" style="display:{display};padding:24px;">'
        con_panels_html += f'<h3 style="font-size:18px;font-weight:700;margin-bottom:16px;">{label}（{len(members)}名）</h3>'

        for name in sorted(members):
            kpi = member_kpi.get(name, {})
            mid = kpi.get('id', '')
            hours = kpi.get('hours', 0)
            daily_count = kpi.get('daily_count', 0)
            fb_count = kpi.get('fb_count', 0)
            mon = kpi.get('monetize', 0)
            achievements = kpi.get('achievements', [])
            mon_display = f"\u00a5{mon:,}" if mon > 0 else '\u00a50'

            con_panels_html += f'''<div class="member-card">
  <div class="member-header">
    <h3>{name}</h3>
    <div class="member-meta">ID: {mid} | コンシェルジュ: {con_name}</div>
  </div>
  <div class="member-kpi-row">
    <div class="member-kpi"><div class="value">{mon_display}</div><div class="label">収益</div></div>
    <div class="member-kpi"><div class="value">{int(hours)}h</div><div class="label">稼働時間</div></div>
    <div class="member-kpi"><div class="value">{daily_count}件</div><div class="label">日報提出</div></div>
    <div class="member-kpi"><div class="value">{fb_count}回</div><div class="label">FB参加</div></div>
  </div>'''

            if achievements:
                con_panels_html += '<div class="member-achievements"><h4>主な成果</h4><ul>'
                for ach in achievements:
                    con_panels_html += f'<li>{ach}</li>'
                con_panels_html += '</ul></div>'
            else:
                con_panels_html += '<div class="member-achievements"><h4>主な成果</h4><div style="color:var(--gray-400);font-size:13px;padding:8px 0;">今週の成果報告なし</div></div>'

            con_panels_html += '</div>'
        con_panels_html += '</div>'

    concierge_section = f'''<div class="section">
<div class="section-header" onclick="this.parentElement.classList.toggle('collapsed')">
  <span class="section-num {"premium" if plan == "PREMIUM" else ""}">R3</span>
  <span class="section-title">コンシェルジュ別メンバー詳細</span>
  <span class="section-toggle">&#9660;</span>
</div>
<div class="section-body" style="padding:0;">
  {con_tabs_html}
  {con_panels_html}
</div></div>'''

    return f'''<div id="view-report" style="display:none;">
  {report_kpi}
  {ranking_section}
  {phase_section}
  {concierge_section}
</div>'''


def build_html_report(data, prev_data, weeks, current_idx):
    """完全なHTMLレポートを生成（日報分析 + 成果報告ビュー切り替え対応）"""
    plan = data['plan_type']
    plan_lower = plan.lower()
    start = data['start_date']
    end = data['end_date']

    # ヘッダー
    header_title = 'PRO + PREMIUM 合算 週次レポート' if plan == 'ALL' else f'{plan} TEAM 週次レポート'
    header = f'''<div class="header {plan_lower}"><div class="header-inner">
  <h1>{header_title}</h1>
  <div class="subtitle">成功習慣行動率100% — 会員分析ダッシュボード</div>
</div></div>'''

    # ナビ
    nav = build_nav_html(plan, weeks, current_idx)

    # 日付表示
    date_display = f'''<div style="font-size:13px;color:var(--gray-600);margin-bottom:20px;">
  <strong style="font-size:18px;color:var(--gray-800);">{start.strftime('%Y年%m月%d日')}（{['月','火','水','木','金','土','日'][start.weekday()]}）〜 {end.strftime('%m月%d日')}（{['月','火','水','木','金','土','日'][end.weekday()]}）</strong><br>
  集計日: {datetime.now().strftime('%Y-%m-%d')} / 対象: {plan}会員 {data["total_members"]}名
</div>'''

    # KPI（日報分析ビュー用）
    kpi = build_kpi_html(data, prev_data)

    # ① 基本統計
    s1_body = '<h4 style="font-size:14px;color:var(--gray-600);margin-bottom:8px;">稼働時間 TOP3</h4>'
    s1_body += build_top3_html(data['hours_ranking'][:3], show_detail=True)
    s1_body += '<h4 style="font-size:14px;color:var(--gray-600);margin:16px 0 8px;">日報提出数 TOP3</h4>'
    count_items = [(n, c, c) for n, c in data['count_ranking'][:3]]
    s1_body += build_top3_html(count_items, show_detail=False)
    section1 = build_section_html(1, '基本統計データ（稼働状況）', s1_body, plan)

    # ② FB会
    day_names = ['月', '火', '水', '木', '金', '土', '日']
    bar_cls = 'premium-bar' if plan == 'PREMIUM' else 'all-bar' if plan == 'ALL' else 'pro-bar'
    max_fb = max(data['fb_by_day'].values()) if data['fb_by_day'] else 1
    s2_body = '<div class="day-chart">'
    for d_idx in range(7):
        count = data['fb_by_day'].get(d_idx, 0)
        h = max(int(count / max_fb * 100), 5) if max_fb > 0 and count > 0 else 0
        s2_body += f'<div class="day-col"><div class="day-bar {bar_cls}" style="height:{h}px;">{count}</div><div class="day-label">{day_names[d_idx]}</div></div>'
    s2_body += f'</div><div style="font-size:12px;color:var(--gray-400);text-align:center;">合計 {data["fb_total"]}件 / 参加者 {data["fb_participants"]}名</div>'
    s2_body += '<h4 style="font-size:14px;color:var(--gray-600);margin:16px 0 8px;">参加回数 TOP3</h4>'
    fb_items = [(n, c, c) for n, c in data['fb_ranking'][:3]]
    s2_body += build_top3_html(fb_items, show_detail=False)
    section2 = build_section_html(2, '実践フィードバック会（FB会）分析', s2_body, plan)

    # ③ 感情分析
    s3_body = f'''<div class="sentiment-gauge">
  <div class="sentiment-pos" style="width:{data['active_rate']}%;">{data['active_rate']}%</div>
  <div class="sentiment-neg" style="width:{data['stagnant_rate']}%;">{data['stagnant_rate']}%</div>
</div>
<div style="display:flex;justify-content:space-between;font-size:12px;color:var(--gray-600);">
  <span>前進・自信・成功体験・学び</span><span>不安・停滞・環境課題・操作の迷い</span>
</div>'''
    if prev_data:
        diff = data['active_rate'] - prev_data['active_rate']
        diff_cls = 'color:#16a34a' if diff >= 0 else 'color:#dc2626'
        diff_sign = '+' if diff >= 0 else ''
        s3_body += f'''<div style="display:flex;gap:16px;margin-top:16px;">
  <div style="flex:1;text-align:center;padding:12px;background:#f8fafc;border-radius:8px;">
    <div style="font-size:11px;color:#94a3b8;font-weight:600;">前週</div>
    <div style="display:flex;gap:4px;justify-content:center;margin-top:4px;">
      <span style="color:#16a34a;font-weight:700;font-size:18px;">{prev_data['active_rate']}%</span>
      <span style="color:#94a3b8;font-size:14px;">/</span>
      <span style="color:#dc2626;font-weight:700;font-size:18px;">{prev_data['stagnant_rate']}%</span>
    </div>
  </div>
  <div style="flex:1;text-align:center;padding:12px;background:#f8fafc;border-radius:8px;">
    <div style="font-size:11px;color:#94a3b8;font-weight:600;">今週</div>
    <div style="display:flex;gap:4px;justify-content:center;margin-top:4px;">
      <span style="color:#16a34a;font-weight:700;font-size:18px;">{data['active_rate']}%</span>
      <span style="color:#94a3b8;font-size:14px;">/</span>
      <span style="color:#dc2626;font-weight:700;font-size:18px;">{data['stagnant_rate']}%</span>
    </div>
  </div>
  <div style="flex:1;text-align:center;padding:12px;background:#f8fafc;border-radius:8px;">
    <div style="font-size:11px;color:#94a3b8;font-weight:600;">変化</div>
    <div style="{diff_cls};font-weight:700;font-size:18px;margin-top:4px;">{diff_sign}{diff}pt</div>
  </div>
</div>'''
    section3 = build_section_html(3, '感情・傾向占有率', s3_body, plan)

    # ⑤ マネタイズフェーズ
    phase_colors = {0: 'p1', 1: 'p2', 2: 'p3', 3: 'p4', 4: 'p5'}
    s5_body = '<div class="phase-grid">'
    for i, (phase, members) in enumerate(sorted(data['phases'].items())):
        p_cls = phase_colors.get(i % 5, 'p1')
        s5_body += f'''<div class="phase-card {p_cls}">
  <div style="font-size:11px;font-weight:700;color:var(--gray-400);">PHASE {i+1}</div>
  <div style="font-size:14px;font-weight:700;margin:2px 0;">{phase}</div>
  <div style="font-size:24px;font-weight:800;color:var(--primary);">{len(members)}名</div>
  <div style="font-size:12px;color:var(--gray-600);margin-top:4px;">{", ".join(members[:5])}{"..." if len(members) > 5 else ""}</div>
</div>'''
    s5_body += '</div>'
    s5_body += f'<div style="font-size:13px;color:var(--gray-600);margin-top:8px;">週報提出者: {data["weekly_submitters"]}名 / {plan}基準{data["total_members"]}名（提出率 {data["weekly_rate"]}%）</div>'
    section5 = build_section_html(5, 'マネタイズフェーズ分析（週報ベース）', s5_body, plan)

    # ⑥ マネタイズ成果
    money_display = f"{int(data['total_money']//10000)}万円" if data['total_money'] >= 10000 else f"{int(data['total_money']):,}円"
    s6_body = f'''<div style="display:flex;gap:16px;margin-bottom:16px;">
  <div style="flex:1;text-align:center;padding:12px;background:var(--gray-50);border-radius:8px;">
    <div style="font-size:11px;color:var(--gray-400);font-weight:600;">合計額</div>
    <div style="font-size:22px;font-weight:800;color:var(--primary);">{money_display}</div>
  </div>
  <div style="flex:1;text-align:center;padding:12px;background:var(--gray-50);border-radius:8px;">
    <div style="font-size:11px;color:var(--gray-400);font-weight:600;">平均額</div>
    <div style="font-size:22px;font-weight:800;color:var(--primary);">{data["avg_money"]:,}円</div>
  </div>
</div>'''
    if len(data['monetize']) >= 3:
        m = data['monetize']
        s6_body += f'''<div class="monetize-podium">
  <div class="podium-item silver"><div class="podium-medal">\U0001f948</div><div class="podium-name">{m[1][0]}</div><div class="podium-amount">{m[1][1]:,}円</div></div>
  <div class="podium-item gold"><div class="podium-medal">\U0001f947</div><div class="podium-name">{m[0][0]}</div><div class="podium-amount">{m[0][1]:,}円</div></div>
  <div class="podium-item bronze"><div class="podium-medal">\U0001f949</div><div class="podium-name">{m[2][0]}</div><div class="podium-amount">{m[2][1]:,}円</div></div>
</div>'''
    if len(data['monetize']) > 3:
        th_cls = 'premium-th' if plan == 'PREMIUM' else 'all-th' if plan == 'ALL' else ''
        s6_body += f'<table><thead><tr><th class="{th_cls}">順位</th><th class="{th_cls}">会員名</th><th class="{th_cls}">金額</th></tr></thead><tbody>'
        for i, (name, amount, _) in enumerate(data['monetize'][3:10], start=4):
            s6_body += f'<tr><td>{i}位</td><td>{name}</td><td>{amount:,}円</td></tr>'
        s6_body += '</tbody></table>'
    elif not data['monetize']:
        s6_body += '<div class="no-data">今週の収益報告はありません</div>'
    section6 = build_section_html(6, 'マネタイズ成果ランキング', s6_body, plan)

    # ⑦ 生の声
    s7_body = '<div class="voice-grid"><div>'
    s7_body += '<div class="voice-section-label pos">前進を支える声</div>'
    for name, text in data['positive_voices']:
        short = text[:150] + '...' if len(text) > 150 else text
        s7_body += f'<div class="voice-card pos"><div class="voice-name">{name}</div>{short}</div>'
    if not data['positive_voices']:
        s7_body += '<div style="color:var(--gray-400);font-size:13px;padding:8px;">該当なし</div>'
    s7_body += '</div><div>'
    s7_body += '<div class="voice-section-label neg">停滞を招く声</div>'
    for name, text in data['negative_voices']:
        short = text[:150] + '...' if len(text) > 150 else text
        s7_body += f'<div class="voice-card neg"><div class="voice-name">{name}</div>{short}</div>'
    if not data['negative_voices']:
        s7_body += '<div style="color:var(--gray-400);font-size:13px;padding:8px;">該当なし</div>'
    s7_body += '</div></div>'
    section7 = build_section_html(7, '会員様の生の声', s7_body, plan)

    # ⑨ 非アクティブ
    th_cls = 'premium-th' if plan == 'PREMIUM' else 'all-th' if plan == 'ALL' else ''
    s9_body = f'''<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px;">
  <div style="background:var(--accent-light);border-radius:10px;padding:20px;">
    <div style="font-size:13px;font-weight:700;color:var(--accent);margin-bottom:12px;">今週（{jp_date(start)}\u2013{jp_date(end)}）</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;">
      <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">非アクティブ</div><div style="font-size:28px;font-weight:800;color:var(--danger);">{data['inactive_count']}名</div></div>
      <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">アクティブ率</div><div style="font-size:28px;font-weight:800;color:var(--success);">{100 - int(data['inactive_count']/data['total_members']*100) if data['total_members'] else 0}%</div></div>
      <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">日報提出率</div><div style="font-size:28px;font-weight:800;color:var(--primary);">{data['submission_rate']}%</div></div>
      <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">週報提出率</div><div style="font-size:28px;font-weight:800;color:var(--accent);">{data['weekly_rate']}%</div><div style="font-size:11px;color:var(--gray-400);">目標 80%</div></div>
    </div>
  </div>'''
    if prev_data:
        prev_active_rate = 100 - int(prev_data['inactive_count']/prev_data['total_members']*100) if prev_data['total_members'] else 0
        s9_body += f'''<div style="background:var(--gray-50);border-radius:10px;padding:20px;">
    <div style="font-size:13px;font-weight:700;color:var(--gray-400);margin-bottom:12px;">前週</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;">
      <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">非アクティブ</div><div style="font-size:28px;font-weight:800;color:var(--gray-400);">{prev_data['inactive_count']}名</div></div>
      <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">アクティブ率</div><div style="font-size:28px;font-weight:800;color:var(--gray-400);">{prev_active_rate}%</div></div>
      <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">日報提出率</div><div style="font-size:28px;font-weight:800;color:var(--gray-400);">{prev_data['submission_rate']}%</div></div>
      <div><div style="font-size:11px;color:var(--gray-400);font-weight:600;">週報提出率</div><div style="font-size:28px;font-weight:800;color:var(--gray-400);">{prev_data['weekly_rate']}%</div></div>
    </div>
  </div>'''
    else:
        s9_body += '<div style="background:var(--gray-50);border-radius:10px;padding:20px;"><div style="font-size:13px;font-weight:700;color:var(--gray-400);">前週データなし</div></div>'
    s9_body += '</div>'

    # 非アクティブ会員テーブル
    if data['inactive']:
        s9_body += f'<div style="overflow-x:auto;"><table><thead><tr><th class="{th_cls}">会員名</th><th class="{th_cls}">今週日報</th><th class="{th_cls}">今週FB</th></tr></thead><tbody>'
        for name in data['inactive']:
            daily_count = len(data['member_daily'].get(name, []))
            fb_count = data['fb_members'].get(name, 0)
            d_cls = 'zero' if daily_count == 0 else ''
            f_cls = 'zero' if fb_count == 0 else ''
            s9_body += f'<tr><td>{name}</td><td class="{d_cls}">{daily_count}/7</td><td class="{f_cls}">{fb_count}回</td></tr>'
        s9_body += '</tbody></table></div>'

    section9 = build_section_html(9, f'非アクティブ{plan}会員', s9_body, plan)

    # 成果報告ビュー
    report_view = build_report_view_html(data)

    # 組み立て
    html = f'''<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{plan} 週次レポート — {jp_date(start)}〜{jp_date(end)}</title>
<style>{CSS}</style>
</head>
<body>
{header}
{nav}
<div class="main">
  {date_display}
  <div id="view-daily">
    {kpi}
    {section1}
    {section2}
    {section3}
    {section5}
    {section6}
    {section7}
    {section9}
  </div>
  {report_view}
</div>
<div class="footer">Generated by SAEPIN — {plan} TEAM AI Agent System<br>集計日: {datetime.now().strftime('%Y-%m-%d')}</div>
<button class="float-top" id="floatTop" onclick="scrollTo({{top:0,behavior:'smooth'}}}})">&#9650;</button>
<script>
const btn=document.getElementById('floatTop');
window.addEventListener('scroll',()=>{{btn.style.display=window.scrollY>400?'flex':'none';}});

function switchView(view) {{
  document.getElementById('view-daily').style.display = view === 'daily' ? 'block' : 'none';
  document.getElementById('view-report').style.display = view === 'report' ? 'block' : 'none';
  document.querySelectorAll('.view-tab').forEach(t => t.classList.remove('active'));
  event.target.classList.add('active');
}}

function switchConcierge(con) {{
  document.querySelectorAll('.concierge-panel').forEach(p => p.style.display = 'none');
  document.getElementById('con-' + con).style.display = 'block';
  document.querySelectorAll('.con-tab').forEach(t => t.classList.remove('active'));
  event.target.classList.add('active');
}}
</script>
</body></html>'''
    return html


def get_filename(plan_type, start_date, end_date):
    """ファイル名を生成"""
    prefix = '合算' if plan_type == 'ALL' else plan_type
    return f"{prefix}_{start_date.month}月{start_date.day}日-{end_date.month}月{end_date.day}日.html"


# ==================== Slack投稿 ====================
def get_slack_client():
    """Slack APIクライアントを取得"""
    from slack_sdk import WebClient
    token = SLACK_TOKEN_FILE.read_text(encoding="utf-8").strip()
    return WebClient(token=token)


GITHUB_PAGES_BASE = "https://saepin0519.github.io/ai-agent"


def get_github_pages_url(filepath):
    """ローカルファイルパスからGitHub PagesのURLを生成"""
    from urllib.parse import quote
    rel_path = filepath.relative_to(BASE_DIR)
    # パスの各部分をURLエンコード
    encoded_parts = [quote(str(part)) for part in rel_path.parts]
    return f"{GITHUB_PAGES_BASE}/{'/'.join(encoded_parts)}"


def post_weekly_slack(client, html_files, start_date, end_date, channel=SLACK_CHANNEL):
    """GitHub PagesのURLリンクをSlackに投稿"""
    date_range = f"{start_date.month}/{start_date.day}〜{end_date.month}/{end_date.day}"

    # GitHub Pages URLを生成
    links = {}
    for plan_label, filepath in html_files.items():
        url = get_github_pages_url(filepath)
        links[plan_label] = url
        print(f"  URL生成: {plan_label} → {url}")

    # メッセージ本文を組み立て
    # 週次レポートとチャットログ分析を分けて表示
    weekly_links = []
    chatlog_link = ""
    for label, permalink in links.items():
        if not permalink:
            continue
        if 'チャットログ' in label:
            chatlog_link = f"<{permalink}|{label}>"
        else:
            weekly_links.append(f"<{permalink}|{label}>")

    lines = [
        f"お疲れ様です。{date_range}の週次報告内容になります。",
        "",
        ":bar_chart: *週次レポート*",
        "  ".join(weekly_links),
    ]

    if chatlog_link:
        lines += [
            "",
            ":speech_balloon: *PREMIUM専用ルーム チャットログ分析*",
            chatlog_link,
            "会員別の相談内容・コンシェルジュ別の稼働時間・応答速度・チャット内容分類を掲載しています。",
        ]

    lines += [
        "",
        "皆さんからのTODOのコメント宜しくお願い致します！",
        "また、私よりclaudecodeで生成自動配信しておりますが、まだまだ改良の余地があるかと存じますので、お気軽にご意見いただけますと幸いです:pray:",
        "",
        "「日報分析」「成果報告」ボタンでビュー切り替えできます。",
    ]

    text = "\n".join(lines)
    client.chat_postMessage(channel=channel, text=text, unfurl_links=False)
    return len(html_files)


# ==================== メイン ====================
def main():
    parser = argparse.ArgumentParser(description="週次会員分析レポート自動生成")
    parser.add_argument("--dry-run", action="store_true", help="Slack投稿なし（テスト用）")
    parser.add_argument("--weeks-ago", type=int, default=1, help="何週間前を分析するか")
    parser.add_argument("--backfill", action="store_true", help="2/8〜直近全週のHTMLを一括生成")
    parser.add_argument("--combined", action="store_true", help="PRO+PREMIUM合算ページも生成")
    parser.add_argument("--generate-only", action="store_true", help="レポート生成+git push+LINE通知のみ（Slack投稿なし）")
    parser.add_argument("--slack-only", action="store_true", help="Slack投稿のみ（レポートは生成済み前提）")
    args = parser.parse_args()

    # 対象週の計算
    today = datetime.now()
    this_monday = today - timedelta(days=today.weekday())

    # slack-onlyモード: Google Sheets接続不要、既存ファイルからSlack投稿のみ
    if args.slack_only:
        target_monday = this_monday - timedelta(weeks=args.weeks_ago)
        target_sunday = target_monday + timedelta(days=6)
        start_date = target_monday.replace(hour=0, minute=0, second=0, microsecond=0)
        end_date = target_sunday.replace(hour=23, minute=59, second=59, microsecond=0)

        html_files = {}
        for plan_type in ['PRO', 'PREMIUM']:
            filename = get_filename(plan_type, start_date, end_date)
            filepath = REPORT_DIR / filename
            if filepath.exists():
                html_files[plan_type] = filepath
        if args.combined:
            all_filename = get_filename('ALL', start_date, end_date)
            all_filepath = REPORT_DIR / all_filename
            if all_filepath.exists():
                html_files['合算'] = all_filepath
        cl_filepath = REPORT_DIR / "チャンネルログ分析.html"
        if cl_filepath.exists():
            html_files['チャットログ分析（全期間）'] = cl_filepath

        print(f"=== Slack投稿のみ（{len(html_files)}ファイル）===")
        try:
            client = get_slack_client()
            post_weekly_slack(client, html_files, start_date, end_date)
            print(f"\nSlack投稿完了")
        except Exception as e:
            print(f"\nSlack投稿エラー: {e}")
        print("\n完了")
        return

    print("Google Sheets接続中...")
    token = get_access_token()
    all_data = fetch_all_data(token)

    # member_listは (id, name, concierge) のタプルリスト
    pro_list = all_data['pro_members']
    premium_list = all_data['premium_members']
    pro_names = [name for _, name, _ in pro_list]
    premium_names = [name for _, name, _ in premium_list]
    print(f"PRO: {len(pro_names)}名, PREMIUM: {len(premium_names)}名")

    if args.backfill:
        # 全週バックフィル
        end_sunday = this_monday - timedelta(days=1)  # 先週の日曜
        weeks = get_weeks(DATA_START, end_sunday)
        print(f"\nバックフィル: {len(weeks)}週間分を生成")

        for plan_type, member_list in [('PRO', pro_list), ('PREMIUM', premium_list)]:
            print(f"\n=== {plan_type} レポート生成 ===")
            prev_result = None
            for i, (start, end) in enumerate(weeks):
                print(f"  {jp_date(start)}〜{jp_date(end)} ...", end='')
                result = analyze_week(all_data, start, end, member_list, plan_type)
                html = build_html_report(result, prev_result, weeks, i)

                filename = get_filename(plan_type, start, end)
                filepath = REPORT_DIR / filename
                filepath.write_text(html, encoding='utf-8')
                print(f" 保存: {filename}")
                prev_result = result

        print(f"\n全HTMLファイルを {REPORT_DIR} に保存完了")
    else:
        # 単週実行（毎週月曜の自動実行用）
        target_monday = this_monday - timedelta(weeks=args.weeks_ago)
        target_sunday = target_monday + timedelta(days=6)
        start_date = target_monday.replace(hour=0, minute=0, second=0, microsecond=0)
        end_date = target_sunday.replace(hour=23, minute=59, second=59, microsecond=0)

        # 前週データも取得（比較用）
        prev_monday = target_monday - timedelta(weeks=1)
        prev_sunday = prev_monday + timedelta(days=6)

        # 直近の全週リスト（ナビ用 — 対象週も必ず含める）
        recent_end = max(this_monday - timedelta(days=1), target_sunday)
        weeks = get_weeks(DATA_START, recent_end)

        # 現在の週のインデックスを特定
        current_idx = next((i for i, (s, e) in enumerate(weeks) if s == target_monday), len(weeks) - 1)

        html_files = {}  # Slack投稿用のファイルパス
        results = {}   # PRO/PREMIUM結果を保持（合算用）
        prev_results = {}
        for plan_type, member_list in [('PRO', pro_list), ('PREMIUM', premium_list)]:
            print(f"\n=== {plan_type} レポート ===")
            result = analyze_week(all_data, start_date, end_date, member_list, plan_type)
            prev_result = analyze_week(all_data, prev_monday, prev_sunday, member_list, plan_type)
            results[plan_type] = result
            prev_results[plan_type] = prev_result

            # HTML保存
            html = build_html_report(result, prev_result, weeks, current_idx)
            filename = get_filename(plan_type, start_date, end_date)
            filepath = REPORT_DIR / filename
            filepath.write_text(html, encoding='utf-8')
            print(f"  HTML保存: {filepath}")
            html_files[plan_type] = filepath

        # 合算ページ生成（PRO + PREMIUM を単純合算）
        if args.combined:
            print(f"\n=== 合算（PRO+PREMIUM）レポート ===")
            all_result = merge_results(results['PRO'], results['PREMIUM'])
            all_prev = merge_results(prev_results['PRO'], prev_results['PREMIUM'])
            print(f"  合算会員数: {all_result['total_members']}名（PRO {results['PRO']['total_members']} + PREMIUM {results['PREMIUM']['total_members']}）")
            all_html = build_html_report(all_result, all_prev, weeks, current_idx)
            all_filename = get_filename('ALL', start_date, end_date)
            all_filepath = REPORT_DIR / all_filename
            all_filepath.write_text(all_html, encoding='utf-8')
            print(f"  HTML保存: {all_filepath}")
            html_files['合算'] = all_filepath

        # チャンネルログ分析（PREMIUM専用ルーム）を実行して結果を追加
        try:
            # 同ディレクトリのスクリプトをインポート
            sys.path.insert(0, str(Path(__file__).parent))
            from channel_log_analysis import get_access_token as cl_get_token, fetch_channel_log, analyze_channel_log, generate_html as cl_generate_html
            print(f"\n=== PREMIUM チャンネルログ分析 ===")
            cl_token = cl_get_token()
            cl_data = fetch_channel_log(cl_token)
            print(f"  取得完了: {len(cl_data)}行")
            cl_analysis = analyze_channel_log(cl_data)
            cl_html = cl_generate_html(cl_analysis)
            cl_filepath = REPORT_DIR / "チャンネルログ分析.html"
            cl_filepath.write_text(cl_html, encoding='utf-8')
            print(f"  HTML保存: {cl_filepath}")
            html_files['チャットログ分析（全期間）'] = cl_filepath
        except Exception as e:
            print(f"\nチャンネルログ分析エラー: {e}")
            print("週次レポートのみ投稿します。")

        # generate-only: git push + LINE通知のみ
        if args.generate_only:
            print("\n=== レポート生成完了 → git push + LINE通知 ===")
            try:
                import subprocess
                # git add & commit & push
                subprocess.run(['git', 'add', str(REPORT_DIR)], cwd=str(BASE_DIR), check=True)
                subprocess.run(['git', 'commit', '-m', '週次レポート自動生成（日曜定時）'], cwd=str(BASE_DIR), check=False)
                subprocess.run(['git', 'push', 'origin', 'main'], cwd=str(BASE_DIR), check=True)
                print("  git push完了")
            except Exception as e:
                print(f"  git pushエラー: {e}")

            # LINE通知
            try:
                sys.path.insert(0, str(BASE_DIR / '09_system'))
                from line_notify import send_to_line
                date_range = f"{start_date.month}/{start_date.day}〜{end_date.month}/{end_date.day}"
                urls = []
                for label, filepath in html_files.items():
                    urls.append(f"・{label}: {get_github_pages_url(filepath)}")
                msg = f"\n【週次レポート生成完了】{date_range}\n\n" + "\n".join(urls) + "\n\n確認してSlackに流す場合は指示してね"
                send_to_line(msg)
                print("  LINE通知完了")
            except Exception as e:
                print(f"  LINE通知エラー: {e}")

        # 通常実行 or dry-run
        elif args.dry_run:
            print("\n=== DRY RUN（Slack投稿スキップ）===")
            for label, path in html_files.items():
                print(f"  {label}: {path}")
        else:
            try:
                client = get_slack_client()
                post_weekly_slack(client, html_files, start_date, end_date)
                print(f"\nSlack投稿完了（{len(html_files)}ファイル + メッセージ）")
            except Exception as e:
                print(f"\nSlack投稿エラー: {e}")
                print("レポートはローカルに保存済みです。")

    print("\n完了")


if __name__ == "__main__":
    main()
