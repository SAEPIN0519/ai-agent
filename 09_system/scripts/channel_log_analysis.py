"""
全チャンネルログ分析スクリプト

PREMIUM専用ルームのチャットログを分析し、
会員別のチャット活用状況 + コンシェルジュ別の対応指標をHTML出力する。

使い方:
  python channel_log_analysis.py              # 分析実行・HTML生成
  python channel_log_analysis.py --dry-run    # 画面出力のみ（HTML保存なし）
"""

import sys
import io
import json
import argparse
from pathlib import Path
from datetime import datetime, timedelta
from collections import defaultdict

# Windows環境のUTF-8対応（直接実行時のみ。他スクリプトからimport時はスキップ）
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# ==================== 設定 ====================
BASE_DIR = Path(__file__).parent.parent.parent
CONFIG_DIR = BASE_DIR / "09_system" / "config"
REPORT_DIR = BASE_DIR / "03_clients" / "SIFTAI" / "プロプレミアムTEAM" / "会員管理" / "週次レポート"
SA_FILE = CONFIG_DIR / "google_service_account.json"

SPREADSHEET_ID = "1EHKgmE1d7T5N9GbT_QJV72Q8Dh91iNysHHkz6LEWJy4"
SHEET_NAME = "全チャンネルログ"

# BOTのDiscord ID（分析から除外）
BOT_DISCORD_IDS = {"1404654033330897087"}

# スタッフ名の正規化マッピング
STAFF_NAME_MAP = {
    "SHIFT AI運営｜安部友博": "安部友博",
    "SHIFT AI 運営 | Tomoya": "安永智也",
    "SHIFT AI運営｜山本 大輔": "山本大輔",
}


# ==================== Google Sheets アクセス ====================
def get_access_token():
    """サービスアカウントのアクセストークンを取得"""
    from google.oauth2.service_account import Credentials
    from google.auth.transport.requests import Request
    creds = Credentials.from_service_account_file(
        str(SA_FILE),
        scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
    )
    creds.refresh(Request())
    return creds.token


def fetch_channel_log(token):
    """全チャンネルログシートのデータを取得"""
    import requests as req
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values/{SHEET_NAME}!A:H"
    resp = req.get(url, headers={"Authorization": f"Bearer {token}"})
    resp.raise_for_status()
    rows = resp.json().get('values', [])
    if not rows:
        return []
    headers = rows[0]
    data = []
    for row in rows[1:]:
        # 行をパディング
        padded = row + [''] * (len(headers) - len(row))
        data.append({
            'channel': padded[0],
            'name': padded[1],
            'discord_id': padded[2],
            'discord_name': padded[3],
            'role': padded[4],         # スタッフ / 会員
            'timestamp': padded[5],
            'content': padded[6],
            'message_id': padded[7],
        })
    return data


def parse_timestamp(ts_str):
    """タイムスタンプ文字列をdatetimeに変換"""
    if not ts_str:
        return None
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S', '%Y-%m-%d', '%Y/%m/%d'):
        try:
            return datetime.strptime(ts_str.strip(), fmt)
        except ValueError:
            continue
    return None


# ==================== チャット内容分類 ====================
CONTENT_CATEGORIES = {
    'AIツール活用': [
        'chatgpt', 'gpt', 'claude', 'gemini', 'copilot', 'midjourney',
        'stable diffusion', 'dall-e', 'dalle', 'perplexity', 'notebooklm',
        'cursor', 'v0', 'bolt', 'replit', 'dify', 'make', 'zapier',
        'ai', 'プロンプト', 'prompt', 'エージェント', 'agent',
        'llm', 'チャットボット', 'api',
    ],
    'ビジネス・マネタイズ': [
        'マネタイズ', '収益', '売上', '案件', 'クライアント', '営業',
        '見積', '請求書', '単価', '報酬', '副業', 'フリーランス',
        '起業', '事業', 'ビジネス', '受注', '納品', '提案',
        '集客', 'リード', '商談', '契約',
    ],
    'スキル・学習': [
        '勉強', '学習', '講座', 'セミナー', 'ウェビナー', '教材',
        '書籍', 'udemy', 'youtube', 'チュートリアル',
        '資格', '検定', 'スキル', '練習', 'インプット',
        'アーカイブ', 'フィードバック会', '出陣式',
    ],
    'コンテンツ制作': [
        'ブログ', '記事', 'note', 'ライティング', '執筆', 'seo',
        'sns', 'twitter', 'instagram', 'tiktok',
        '画像生成', '動画制作', 'サムネ', 'デザイン', 'canva',
        'lp', 'ランディング', 'ポートフォリオ', 'web',
    ],
    '技術・開発': [
        'python', 'javascript', 'html', 'css', 'コード', 'プログラミング',
        'github', 'git', 'vscode', 'スクリプト', '自動化', 'rpa',
        'スプレッドシート', 'エクセル', 'google sheets', 'notion',
        'データ分析', 'スクレイピング', 'api連携',
        'powershell', 'ターミナル', 'anti-gravity',
    ],
    '案件活動・営業': [
        'クラウドワークス', 'ランサーズ', 'ココナラ', '応募',
        '面談', 'ポートフォリオ', '実績', '営業',
        '行動計画', '宿題', 'step',
    ],
    '相談・悩み': [
        '悩み', '不安', '困', 'わからない', 'わかりません', '迷',
        'どうすれば', 'どうしたら', 'アドバイス', '相談',
        'モチベーション', 'やる気', '挫折', '続かない',
        '気を付ける', 'ポイント', '教えて',
    ],
    '成果・報告': [
        'できました', 'できた', '完成', '達成', '成功', '受注',
        '初案件', '納品', '合格', '成果', '報告', 'やった',
        '連絡ありました', '頑張り',
    ],
    '面談・日程調整': [
        '面談', '日程', '時間', '何時', '大丈夫です', '承知',
        '午前', '午後', '空い', '調整', '予約', 'お願いします',
        '○○時', 'meet.google', 'zoom', '通話',
    ],
    '事務・手続き': [
        '支払', '決済', '返金', '退会', '入会', '会員',
        '手続き', 'カード', 'アカウント', 'ログイン', 'パスワード',
        '永久会員', 'プレミアム会員', '月払い', '会費',
    ],
    '挨拶・返信': [
        'ありがとう', 'お疲れ様', 'よろしくお願い', '了解',
        'こんにちは', 'こんばんは', 'こんばんわ', 'おはよう',
        '承知', 'すみません', '申し訳', 'ごめん',
        'わかりました', '大丈夫です', '確認します',
    ],
}


def classify_content(text):
    """テキストをカテゴリに分類（複数カテゴリ可）"""
    if not text:
        return []
    text_lower = text.lower()
    matched = []
    for category, keywords in CONTENT_CATEGORIES.items():
        for kw in keywords:
            if kw in text_lower:
                matched.append(category)
                break
    return matched if matched else ['その他']


# ==================== 分析ロジック ====================
def analyze_channel_log(data):
    """全チャンネルログを分析"""

    # BOT除外
    filtered = [d for d in data if d['discord_id'] not in BOT_DISCORD_IDS]

    # 会員メッセージとスタッフメッセージを分離
    member_msgs = [d for d in filtered if d['role'] == '会員']
    staff_msgs = [d for d in filtered if d['role'] == 'スタッフ']

    # ----- 会員別分析 -----
    member_stats = defaultdict(lambda: {
        'channel': '',
        'msg_count': 0,
        'total_chars': 0,
        'messages': [],
        'first_post': None,
        'last_post': None,
    })

    for msg in member_msgs:
        name = msg['name']
        ts = parse_timestamp(msg['timestamp'])
        content = msg['content'] or ''
        char_count = len(content)

        stats = member_stats[name]
        stats['channel'] = msg['channel']
        stats['msg_count'] += 1
        stats['total_chars'] += char_count
        stats['messages'].append({
            'timestamp': ts,
            'content': content,
            'chars': char_count,
        })
        if ts:
            if stats['first_post'] is None or ts < stats['first_post']:
                stats['first_post'] = ts
            if stats['last_post'] is None or ts > stats['last_post']:
                stats['last_post'] = ts

    # 平均文字数を計算 + 会員別トピック分類
    overall_categories = defaultdict(int)  # 全体カテゴリ集計
    for name, stats in member_stats.items():
        stats['avg_chars'] = stats['total_chars'] // stats['msg_count'] if stats['msg_count'] else 0
        # 会員ごとのカテゴリ集計
        cat_counts = defaultdict(int)
        for msg_item in stats['messages']:
            cats = classify_content(msg_item['content'])
            for cat in cats:
                cat_counts[cat] += 1
                overall_categories[cat] += 1
        # 上位3カテゴリを保持
        sorted_cats = sorted(cat_counts.items(), key=lambda x: -x[1])
        stats['top_categories'] = sorted_cats[:3]
        stats['all_categories'] = dict(cat_counts)

    # ----- コンシェルジュ別分析 -----
    staff_stats = defaultdict(lambda: {
        'reply_count': 0,
        'total_chars': 0,
        'channels': set(),
        'messages': [],
    })

    for msg in staff_msgs:
        raw_name = msg['name']
        name = STAFF_NAME_MAP.get(raw_name, raw_name)
        ts = parse_timestamp(msg['timestamp'])
        content = msg['content'] or ''

        stats = staff_stats[name]
        stats['reply_count'] += 1
        stats['total_chars'] += len(content)
        stats['channels'].add(msg['channel'])
        stats['messages'].append({
            'channel': msg['channel'],
            'timestamp': ts,
            'content': content,
            'chars': len(content),
        })

    for name, stats in staff_stats.items():
        stats['channel_count'] = len(stats['channels'])
        stats['avg_chars'] = stats['total_chars'] // stats['reply_count'] if stats['reply_count'] else 0
        stats['per_member_replies'] = round(stats['reply_count'] / stats['channel_count'], 1) if stats['channel_count'] else 0
        # Discord稼働時間（1返信 = 15分）
        stats['discord_hours'] = round(stats['reply_count'] * 15 / 60, 1)

    # ----- チャンネル→担当コンシェルジュ マッピング -----
    # 各チャンネルで最も返信が多いスタッフを担当者とする
    channel_staff_counts = defaultdict(lambda: defaultdict(int))
    for msg in staff_msgs:
        raw_name = msg['name']
        name = STAFF_NAME_MAP.get(raw_name, raw_name)
        channel_staff_counts[msg['channel']][name] += 1

    channel_concierge = {}
    for ch, counts in channel_staff_counts.items():
        channel_concierge[ch] = max(counts, key=counts.get)

    # 会員にコンシェルジュ名を紐づけ
    for name, stats in member_stats.items():
        stats['concierge'] = channel_concierge.get(stats['channel'], '未割当')

    # ----- 応答時間分析（チャンネルごと） -----
    # チャンネルごとにメッセージを時系列ソート
    channel_msgs = defaultdict(list)
    for msg in filtered:
        ts = parse_timestamp(msg['timestamp'])
        if ts:
            channel_msgs[msg['channel']].append({
                'role': msg['role'],
                'name': msg['name'],
                'timestamp': ts,
                'content': msg['content'],
            })

    # 各チャンネルのメッセージを時系列ソート
    for ch in channel_msgs:
        channel_msgs[ch].sort(key=lambda x: x['timestamp'])

    # 会員投稿→最初のスタッフ返信までの時間を計測
    response_times = defaultdict(list)  # コンシェルジュ名 -> [時間(hours)のリスト]
    channel_response_times = {}  # チャンネル名 -> 平均応答時間

    for ch, msgs in channel_msgs.items():
        ch_response_times = []
        i = 0
        while i < len(msgs):
            if msgs[i]['role'] == '会員':
                member_ts = msgs[i]['timestamp']
                # 次のスタッフ返信を探す
                j = i + 1
                while j < len(msgs):
                    if msgs[j]['role'] == 'スタッフ':
                        staff_name = STAFF_NAME_MAP.get(msgs[j]['name'], msgs[j]['name'])
                        delta = (msgs[j]['timestamp'] - member_ts).total_seconds() / 3600
                        if delta >= 0:  # 負の時間は無視
                            response_times[staff_name].append(delta)
                            ch_response_times.append(delta)
                        break
                    j += 1
            i += 1

        if ch_response_times:
            channel_response_times[ch] = round(sum(ch_response_times) / len(ch_response_times), 1)

    # コンシェルジュ別の平均応答時間
    staff_avg_response = {}
    for name, times in response_times.items():
        if times:
            staff_avg_response[name] = round(sum(times) / len(times), 1)

    # チーム全体の平均応答時間
    all_times = [t for times in response_times.values() for t in times]
    team_avg_response = round(sum(all_times) / len(all_times), 1) if all_times else 0

    # ----- データ期間の週数計算 -----
    all_timestamps = [parse_timestamp(d['timestamp']) for d in filtered]
    all_timestamps = [t for t in all_timestamps if t]
    if all_timestamps:
        data_start = min(all_timestamps)
        data_end = max(all_timestamps)
        data_weeks = max((data_end - data_start).days / 7, 1)
    else:
        data_weeks = 1
        data_start = None
        data_end = None

    # コンシェルジュの週あたり・月あたりDiscord稼働時間
    for name, stats in staff_stats.items():
        stats['discord_hours_per_week'] = round(stats['discord_hours'] / data_weeks, 1)
        stats['discord_hours_per_month'] = round(stats['discord_hours_per_week'] * 4.3, 1)
        stats['per_member_monthly_hours'] = round(stats['discord_hours_per_month'] / stats['channel_count'], 1) if stats['channel_count'] else 0

    # ----- サマリー -----
    summary = {
        'total_messages': len(filtered),
        'member_messages': len(member_msgs),
        'staff_messages': len(staff_msgs),
        'channel_count': len(set(d['channel'] for d in filtered)),
        'member_count': len(member_stats),
        'staff_count': len(staff_stats),
        'team_avg_response_hours': team_avg_response,
        'data_weeks': round(data_weeks, 1),
        'data_start': data_start,
        'data_end': data_end,
    }

    return {
        'summary': summary,
        'member_stats': dict(member_stats),
        'staff_stats': {k: {**v, 'channels': list(v['channels'])} for k, v in staff_stats.items()},
        'staff_avg_response': staff_avg_response,
        'channel_response_times': channel_response_times,
        'overall_categories': dict(overall_categories),
    }


# ==================== HTML生成 ====================
def format_hours(h):
    """時間をわかりやすい文字列に変換"""
    if h < 1:
        return f"{int(h * 60)}分"
    elif h < 24:
        hours = int(h)
        minutes = int((h - hours) * 60)
        return f"{hours}時間{minutes}分" if minutes else f"{hours}時間"
    else:
        days = int(h // 24)
        hours = int(h % 24)
        return f"{days}日{hours}時間" if hours else f"{days}日"


def generate_html(analysis):
    """分析結果をHTMLに変換"""
    summary = analysis['summary']
    overall_categories = analysis.get('overall_categories', {})
    member_stats = analysis['member_stats']
    staff_stats = analysis['staff_stats']
    staff_avg_response = analysis['staff_avg_response']
    channel_response_times = analysis['channel_response_times']
    team_avg = summary['team_avg_response_hours']

    # メンバーを投稿数でソート
    sorted_members = sorted(member_stats.items(), key=lambda x: -x[1]['msg_count'])
    # スタッフを返信数でソート
    sorted_staff = sorted(staff_stats.items(), key=lambda x: -x[1]['reply_count'])

    now = datetime.now()

    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PREMIUM専用ルーム チャットログ分析</title>
<style>
:root {{
  --primary: #4c1d95; --accent: #7c3aed; --accent-light: #f5f3ff;
  --success: #16a34a; --warning: #d97706; --danger: #dc2626;
  --gray-50: #f8fafc; --gray-100: #f1f5f9; --gray-200: #e2e8f0;
  --gray-300: #cbd5e1; --gray-400: #94a3b8; --gray-600: #475569; --gray-800: #1e293b;
}}
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Hiragino Kaku Gothic ProN','Yu Gothic UI','Meiryo',sans-serif; background:var(--gray-100); color:var(--gray-800); line-height:1.7; }}
.header {{ background:linear-gradient(135deg,#4c1d95,#7c3aed); color:#fff; padding:32px 40px 24px; }}
.header-inner {{ max-width:1100px; margin:0 auto; }}
.header h1 {{ font-size:26px; font-weight:700; }}
.header .subtitle {{ font-size:14px; opacity:0.85; margin-top:4px; }}
.main {{ max-width:1100px; margin:0 auto; padding:28px 20px 60px; }}
.kpi-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:16px; margin-bottom:32px; }}
.kpi-card {{ background:#fff; border-radius:10px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.06); }}
.kpi-label {{ font-size:12px; color:var(--gray-400); font-weight:600; text-transform:uppercase; letter-spacing:0.5px; }}
.kpi-value {{ font-size:28px; font-weight:700; color:var(--primary); margin:4px 0; }}
.kpi-sub {{ font-size:12px; color:var(--gray-400); margin-top:2px; }}
.section {{ background:#fff; border-radius:10px; box-shadow:0 1px 3px rgba(0,0,0,0.06); margin-bottom:24px; overflow:hidden; }}
.section-header {{ display:flex; align-items:center; gap:10px; padding:16px 24px; border-bottom:1px solid var(--gray-200); cursor:pointer; user-select:none; }}
.section-header:hover {{ background:var(--gray-50); }}
.section-num {{ display:inline-flex; align-items:center; justify-content:center; width:28px; height:28px; background:var(--accent); color:#fff; border-radius:50%; font-size:13px; font-weight:700; flex-shrink:0; }}
.section-title {{ font-size:16px; font-weight:700; color:var(--gray-800); flex:1; }}
.section-toggle {{ font-size:18px; color:var(--gray-400); transition:transform 0.2s; }}
.section.collapsed .section-toggle {{ transform:rotate(-90deg); }}
.section.collapsed .section-body {{ display:none; }}
.section-body {{ padding:20px 24px; }}
table {{ width:100%; border-collapse:collapse; font-size:13px; margin:12px 0; }}
th {{ background:var(--primary); color:#fff; padding:10px 12px; text-align:left; font-weight:600; font-size:12px; white-space:nowrap; }}
td {{ padding:8px 12px; border-bottom:1px solid var(--gray-200); }}
tr:nth-child(even) td {{ background:var(--gray-50); }}
tr:hover td {{ background:var(--accent-light); }}
.bar-container {{ display:flex; align-items:center; gap:8px; }}
.bar {{ height:20px; border-radius:4px; min-width:2px; }}
.bar-purple {{ background:linear-gradient(90deg,#a78bfa,#7c3aed); }}
.bar-blue {{ background:linear-gradient(90deg,#60a5fa,#3b82f6); }}
.bar-green {{ background:linear-gradient(90deg,#4ade80,#22c55e); }}
.bar-orange {{ background:linear-gradient(90deg,#fbbf24,#f59e0b); }}
.bar-value {{ font-size:12px; color:var(--gray-600); white-space:nowrap; }}
.staff-profile {{ border:1px solid var(--gray-200); border-radius:12px; padding:24px; margin-bottom:20px; }}
.staff-header {{ display:flex; align-items:center; gap:16px; margin-bottom:16px; }}
.staff-avatar {{ width:48px; height:48px; border-radius:50%; background:var(--accent); color:#fff; display:flex; align-items:center; justify-content:center; font-size:20px; font-weight:700; }}
.staff-name {{ font-size:18px; font-weight:700; }}
.staff-meta {{ font-size:12px; color:var(--gray-400); }}
.staff-kpi-row {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(140px,1fr)); gap:12px; margin:16px 0; }}
.staff-kpi {{ text-align:center; padding:12px; background:var(--gray-50); border-radius:8px; }}
.staff-kpi .value {{ font-size:1.4rem; font-weight:800; color:var(--primary); }}
.staff-kpi .label {{ font-size:11px; color:var(--gray-400); margin-top:2px; }}
.team-avg {{ font-size:12px; color:var(--gray-400); margin-top:4px; }}
.team-avg span {{ font-weight:700; color:var(--accent); }}
.response-indicator {{ display:inline-block; padding:2px 8px; border-radius:12px; font-size:11px; font-weight:600; }}
.response-fast {{ background:#dcfce7; color:#166534; }}
.response-normal {{ background:#fef3c7; color:#92400e; }}
.response-slow {{ background:#fef2f2; color:#991b1b; }}
.note {{ font-size:12px; color:var(--gray-400); padding:12px 16px; background:var(--gray-50); border-radius:8px; margin-top:16px; line-height:1.6; }}
.footer {{ text-align:center; padding:24px; color:var(--gray-400); font-size:12px; }}
@media (max-width:768px) {{
  .kpi-grid {{ grid-template-columns:repeat(2,1fr); }}
  .staff-kpi-row {{ grid-template-columns:repeat(2,1fr); }}
}}
</style>
</head>
<body>

<div class="header">
  <div class="header-inner">
    <h1>PREMIUM専用ルーム チャットログ分析</h1>
    <div class="subtitle">生成日時: {now.strftime('%Y年%m月%d日 %H:%M')} | データ期間: 全期間</div>
  </div>
</div>

<div class="main">

<!-- KPIサマリー -->
<div class="kpi-grid">
  <div class="kpi-card">
    <div class="kpi-label">総メッセージ数</div>
    <div class="kpi-value">{summary['total_messages']}</div>
    <div class="kpi-sub">BOT除外後</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">会員メッセージ</div>
    <div class="kpi-value">{summary['member_messages']}</div>
    <div class="kpi-sub">{summary['member_count']}名</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">スタッフ返信</div>
    <div class="kpi-value">{summary['staff_messages']}</div>
    <div class="kpi-sub">{summary['staff_count']}名</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">チャンネル数</div>
    <div class="kpi-value">{summary['channel_count']}</div>
    <div class="kpi-sub">PREMIUM専用ルーム</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">チーム平均応答時間</div>
    <div class="kpi-value">{format_hours(team_avg)}</div>
    <div class="kpi-sub">会員投稿→初回スタッフ返信</div>
  </div>
</div>
"""

    # ----- セクション1: コンシェルジュ別プロファイル -----
    html += """
<div class="section" id="sec-staff">
  <div class="section-header" onclick="toggleSection('sec-staff')">
    <span class="section-num">1</span>
    <span class="section-title">コンシェルジュ別 対応プロファイル</span>
    <span class="section-toggle">▼</span>
  </div>
  <div class="section-body">
    <div class="note">
      各コンシェルジュの対応状況を「チーム平均」と比較して表示しています。
      チャットログ上の数値であり、面談・音声対応等は含みません。
    </div>
"""

    # チーム平均の計算
    all_reply_counts = [s['reply_count'] for s in staff_stats.values()]
    all_avg_chars = [s['avg_chars'] for s in staff_stats.values()]
    all_per_member = [s['per_member_replies'] for s in staff_stats.values()]
    team_avg_replies = round(sum(all_reply_counts) / len(all_reply_counts), 1) if all_reply_counts else 0
    team_avg_chars = round(sum(all_avg_chars) / len(all_avg_chars)) if all_avg_chars else 0
    team_avg_per_member = round(sum(all_per_member) / len(all_per_member), 1) if all_per_member else 0

    for staff_name, stats in sorted_staff:
        avg_resp = staff_avg_response.get(staff_name, None)
        avg_resp_display = format_hours(avg_resp) if avg_resp is not None else "N/A"

        # 応答速度のインジケーター
        if avg_resp is not None:
            if avg_resp <= team_avg * 0.8:
                resp_class = "response-fast"
                resp_label = "速い"
            elif avg_resp <= team_avg * 1.2:
                resp_class = "response-normal"
                resp_label = "平均的"
            else:
                resp_class = "response-slow"
                resp_label = "ゆっくり"
        else:
            resp_class = "response-normal"
            resp_label = ""

        initial = staff_name[0] if staff_name else "?"

        html += f"""
    <div class="staff-profile">
      <div class="staff-header">
        <div class="staff-avatar">{initial}</div>
        <div>
          <div class="staff-name">{staff_name}</div>
          <div class="staff-meta">担当: {stats['channel_count']}チャンネル</div>
        </div>
      </div>
      <div class="staff-kpi-row">
        <div class="staff-kpi">
          <div class="value">{stats['reply_count']}</div>
          <div class="label">総返信数</div>
          <div class="team-avg">チーム平均: <span>{team_avg_replies}</span></div>
        </div>
        <div class="staff-kpi">
          <div class="value">{stats['avg_chars']}</div>
          <div class="label">平均文字数/返信</div>
          <div class="team-avg">チーム平均: <span>{team_avg_chars}</span></div>
        </div>
        <div class="staff-kpi">
          <div class="value">{stats['per_member_replies']}</div>
          <div class="label">会員1人あたり返信</div>
          <div class="team-avg">チーム平均: <span>{team_avg_per_member}</span></div>
        </div>
        <div class="staff-kpi">
          <div class="value">{avg_resp_display}</div>
          <div class="label">平均応答時間 <span class="{resp_class} response-indicator">{resp_label}</span></div>
          <div class="team-avg">チーム平均: <span>{format_hours(team_avg)}</span></div>
        </div>
        <div class="staff-kpi">
          <div class="value">{stats['discord_hours']}h</div>
          <div class="label">Discord稼働（累計）</div>
          <div class="team-avg">週あたり: <span>{stats['discord_hours_per_week']}h</span>（月{stats['discord_hours_per_month']}h）</div>
        </div>
      </div>
      <div style="margin-top:12px;">
        <strong style="font-size:13px;">担当チャンネル:</strong>
        <span style="font-size:13px; color:var(--gray-600);">{', '.join(ch.replace('💎-', '').replace('様専用ルーム', '') for ch in stats['channels'])}</span>
      </div>
    </div>
"""

    html += """
  </div>
</div>
"""

    # ----- セクション2: 会員別チャット活用状況 -----
    html += """
<div class="section" id="sec-member">
  <div class="section-header" onclick="toggleSection('sec-member')">
    <span class="section-num">2</span>
    <span class="section-title">会員別 チャット活用状況</span>
    <span class="section-toggle">▼</span>
  </div>
  <div class="section-body">
    <table>
      <thead>
        <tr>
          <th>会員名</th>
          <th>担当</th>
          <th>投稿数</th>
          <th>総文字数</th>
          <th>平均文字数</th>
          <th>利用期間</th>
          <th>主な相談内容</th>
          <th>投稿数（ビジュアル）</th>
        </tr>
      </thead>
      <tbody>
"""

    max_msg_count = max((s['msg_count'] for s in member_stats.values()), default=1)

    for name, stats in sorted_members:
        # 利用期間の計算
        if stats['first_post'] and stats['last_post']:
            days = (stats['last_post'] - stats['first_post']).days
            period = f"{days}日間" if days > 0 else "1日"
        else:
            period = "-"

        bar_width = int(stats['msg_count'] / max_msg_count * 200) if max_msg_count else 0

        concierge = stats.get('concierge', '未割当')
        top_cats = stats.get('top_categories', [])
        cats_display = '、'.join(f'{cat}({cnt})' for cat, cnt in top_cats) if top_cats else '-'

        html += f"""
        <tr>
          <td><strong>{name}</strong></td>
          <td style="font-size:12px; color:var(--gray-600);">{concierge}</td>
          <td>{stats['msg_count']}</td>
          <td>{stats['total_chars']:,}</td>
          <td>{stats['avg_chars']}</td>
          <td>{period}</td>
          <td style="font-size:11px; max-width:200px;">{cats_display}</td>
          <td>
            <div class="bar-container">
              <div class="bar bar-purple" style="width:{bar_width}px;"></div>
              <span class="bar-value">{stats['msg_count']}件</span>
            </div>
          </td>
        </tr>
"""

    html += """
      </tbody>
    </table>
  </div>
</div>
"""

    # ----- セクション3: チャンネル別応答時間 -----
    html += """
<div class="section" id="sec-response">
  <div class="section-header" onclick="toggleSection('sec-response')">
    <span class="section-num">3</span>
    <span class="section-title">チャンネル別 平均応答時間</span>
    <span class="section-toggle">▼</span>
  </div>
  <div class="section-body">
    <div class="note">
      会員様の投稿から最初のスタッフ返信までの平均時間です。短いほど迅速な対応を示します。
    </div>
    <table>
      <thead>
        <tr>
          <th>チャンネル</th>
          <th>平均応答時間</th>
          <th>速度</th>
          <th>応答時間（ビジュアル）</th>
        </tr>
      </thead>
      <tbody>
"""

    sorted_channels = sorted(channel_response_times.items(), key=lambda x: x[1])
    max_resp_time = max(channel_response_times.values(), default=1)

    for ch, avg_time in sorted_channels:
        display_name = ch.replace('💎-', '').replace('様専用ルーム', '')

        if avg_time <= team_avg * 0.8:
            speed_class = "response-fast"
            speed_label = "速い"
        elif avg_time <= team_avg * 1.2:
            speed_class = "response-normal"
            speed_label = "平均的"
        else:
            speed_class = "response-slow"
            speed_label = "ゆっくり"

        bar_width = int(avg_time / max_resp_time * 200) if max_resp_time else 0
        bar_color = "bar-green" if avg_time <= team_avg else "bar-orange"

        html += f"""
        <tr>
          <td><strong>{display_name}</strong></td>
          <td>{format_hours(avg_time)}</td>
          <td><span class="{speed_class} response-indicator">{speed_label}</span></td>
          <td>
            <div class="bar-container">
              <div class="bar {bar_color}" style="width:{bar_width}px;"></div>
              <span class="bar-value">{format_hours(avg_time)}</span>
            </div>
          </td>
        </tr>
"""

    html += """
      </tbody>
    </table>
  </div>
</div>
"""

    # ----- セクション4: コンシェルジュ別 Discord稼働時間 -----
    data_weeks = summary.get('data_weeks', 1)
    data_start = summary.get('data_start')
    data_end = summary.get('data_end')
    period_label = ""
    if data_start and data_end:
        period_label = f"{data_start.strftime('%m/%d')}〜{data_end.strftime('%m/%d')}（約{data_weeks}週間）"

    html += f"""
<div class="section" id="sec-discord-hours">
  <div class="section-header" onclick="toggleSection('sec-discord-hours')">
    <span class="section-num">4</span>
    <span class="section-title">コンシェルジュ別 Discord稼働時間</span>
    <span class="section-toggle">▼</span>
  </div>
  <div class="section-body">
    <div class="note">
      1返信 = 15分として算出した推定稼働時間です。面談・音声対応等は含みません。<br>
      データ期間: {period_label}
    </div>
    <table>
      <thead>
        <tr>
          <th>コンシェルジュ</th>
          <th>総返信数</th>
          <th>受け持ち会員数</th>
          <th>累計稼働時間</th>
          <th>週あたり稼働時間</th>
          <th>月換算</th>
          <th>1人あたり月稼働</th>
          <th>稼働時間（ビジュアル）</th>
        </tr>
      </thead>
      <tbody>
"""

    max_discord_hours = max((s['discord_hours'] for s in staff_stats.values()), default=1)

    for staff_name, stats in sorted_staff:
        bar_width = int(stats['discord_hours'] / max_discord_hours * 200) if max_discord_hours else 0

        html += f"""
        <tr>
          <td><strong>{staff_name}</strong></td>
          <td>{stats['reply_count']}件</td>
          <td>{stats['channel_count']}名</td>
          <td><strong>{stats['discord_hours']}時間</strong></td>
          <td>{stats['discord_hours_per_week']}時間/週</td>
          <td>（月{stats['discord_hours_per_month']}時間）</td>
          <td>{stats['per_member_monthly_hours']}時間</td>
          <td>
            <div class="bar-container">
              <div class="bar bar-blue" style="width:{bar_width}px;"></div>
              <span class="bar-value">{stats['discord_hours']}h</span>
            </div>
          </td>
        </tr>
"""

    total_discord_hours = sum(s['discord_hours'] for s in staff_stats.values())
    total_weekly = round(total_discord_hours / data_weeks, 1) if data_weeks else 0

    html += f"""
      </tbody>
      <tfoot>
        <tr style="background:var(--gray-100); font-weight:700;">
          <td>チーム合計</td>
          <td>{sum(s['reply_count'] for s in staff_stats.values())}件</td>
          <td>{sum(s['channel_count'] for s in staff_stats.values())}名</td>
          <td>{total_discord_hours}時間</td>
          <td>{total_weekly}時間/週</td>
          <td>（月{round(total_weekly * 4.3, 1)}時間）</td>
          <td>—</td>
          <td></td>
        </tr>
      </tfoot>
    </table>
  </div>
</div>
"""

    # ----- セクション5: チャット内容分類（全体） -----
    sorted_overall_cats = sorted(overall_categories.items(), key=lambda x: -x[1])
    total_cat_count = sum(c for _, c in sorted_overall_cats) if sorted_overall_cats else 1
    max_cat_count = sorted_overall_cats[0][1] if sorted_overall_cats else 1

    # カテゴリ色マップ
    cat_colors = {
        'AIツール活用': 'bar-purple', 'ビジネス・マネタイズ': 'bar-orange',
        'スキル・学習': 'bar-blue', 'コンテンツ制作': 'bar-green',
        '技術・開発': '#6366f1', '相談・悩み': '#ef4444',
        '成果・報告': '#10b981', 'その他': '#94a3b8',
    }

    html += """
<div class="section" id="sec-categories">
  <div class="section-header" onclick="toggleSection('sec-categories')">
    <span class="section-num">5</span>
    <span class="section-title">チャット内容分類（全体）</span>
    <span class="section-toggle">▼</span>
  </div>
  <div class="section-body">
    <div class="note">
      会員様の投稿をキーワードベースで自動分類しています（1投稿が複数カテゴリに該当する場合あり）。
    </div>
    <table>
      <thead>
        <tr>
          <th>カテゴリ</th>
          <th>該当数</th>
          <th>割合</th>
          <th>分布（ビジュアル）</th>
        </tr>
      </thead>
      <tbody>
"""

    for cat_name, cat_count in sorted_overall_cats:
        pct = round(cat_count / total_cat_count * 100, 1)
        bar_width = int(cat_count / max_cat_count * 250) if max_cat_count else 0
        bar_class = cat_colors.get(cat_name, 'bar-purple')
        # CSSクラスかカスタム色か判定
        if bar_class.startswith('#'):
            bar_style = f'background:{bar_class}; width:{bar_width}px;'
            bar_cls = 'bar'
        else:
            bar_style = f'width:{bar_width}px;'
            bar_cls = f'bar {bar_class}'

        html += f"""
        <tr>
          <td><strong>{cat_name}</strong></td>
          <td>{cat_count}件</td>
          <td>{pct}%</td>
          <td>
            <div class="bar-container">
              <div class="{bar_cls}" style="{bar_style} height:20px; border-radius:4px; min-width:2px;"></div>
              <span class="bar-value">{pct}%</span>
            </div>
          </td>
        </tr>
"""

    html += """
      </tbody>
    </table>
  </div>
</div>
"""

    # ----- フッター -----
    html += f"""
<div class="footer">
  PREMIUM専用ルーム チャットログ分析 | 自動生成: {now.strftime('%Y-%m-%d %H:%M')} | BOTメッセージ除外済み
</div>

</div><!-- /main -->

<script>
function toggleSection(id) {{
  document.getElementById(id).classList.toggle('collapsed');
}}
</script>
</body>
</html>"""

    return html


# ==================== テキストサマリー ====================
def print_summary(analysis):
    """分析結果をコンソールに出力"""
    s = analysis['summary']
    print(f"\n{'='*60}")
    print(f"PREMIUM専用ルーム チャットログ分析")
    print(f"{'='*60}")
    print(f"総メッセージ: {s['total_messages']}件（BOT除外後）")
    print(f"  会員: {s['member_messages']}件 / スタッフ: {s['staff_messages']}件")
    print(f"チャンネル数: {s['channel_count']} / 会員数: {s['member_count']}名")
    print(f"チーム平均応答時間: {format_hours(s['team_avg_response_hours'])}")

    print(f"\n--- コンシェルジュ別 ---")
    for name, stats in sorted(analysis['staff_stats'].items(), key=lambda x: -x[1]['reply_count']):
        avg_resp = analysis['staff_avg_response'].get(name, None)
        resp_str = format_hours(avg_resp) if avg_resp else "N/A"
        print(f"  {name}: 返信{stats['reply_count']}件 / 担当{stats['channel_count']}ch / "
              f"平均{stats['avg_chars']}字 / 1人あたり{stats['per_member_replies']}件 / "
              f"応答{resp_str}")

    print(f"\n--- 会員別（投稿数順） ---")
    for name, stats in sorted(analysis['member_stats'].items(), key=lambda x: -x[1]['msg_count']):
        print(f"  {name}: {stats['msg_count']}件 / {stats['total_chars']:,}字 / 平均{stats['avg_chars']}字")

    print()


# ==================== メイン ====================
def main():
    parser = argparse.ArgumentParser(description='全チャンネルログ分析')
    parser.add_argument('--dry-run', action='store_true', help='HTML保存なし（画面出力のみ）')
    args = parser.parse_args()

    print("全チャンネルログ分析を開始...")

    # データ取得
    print("  Google Sheetsからデータ取得中...")
    token = get_access_token()
    data = fetch_channel_log(token)
    print(f"  取得完了: {len(data)}行")

    # 分析
    print("  分析中...")
    analysis = analyze_channel_log(data)

    # コンソール出力
    print_summary(analysis)

    # HTML生成・保存
    if not args.dry_run:
        html = generate_html(analysis)
        output_path = REPORT_DIR / "チャンネルログ分析.html"
        output_path.write_text(html, encoding='utf-8')
        print(f"  HTML保存: {output_path}")
    else:
        print("  --dry-run: HTML保存スキップ")

    print("完了!")


if __name__ == '__main__':
    main()
