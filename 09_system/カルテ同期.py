"""
Google Sheetsカルテ → member_roster.md / kpi_progress.md 自動同期スクリプト
- 毎朝7:30にタスクスケジューラで自動実行（タスク配信の前に同期）
- カルテの最新データを取得し、ローカルのMDファイルを更新する
- SAEPINが常に最新の会員情報を把握できる状態を維持

使い方:
  python 09_system/カルテ同期.py          # 通常実行（ファイル更新）
  python 09_system/カルテ同期.py dry-run   # テスト実行（ファイル更新なし）
"""

import pickle
import sys
import os
import io
from datetime import datetime
from pathlib import Path

# Windows環境でのUTF-8出力対応
os.environ["PYTHONIOENCODING"] = "utf-8"
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

from googleapiclient.discovery import build
from google.auth.transport.requests import Request

# === 設定 ===
PROJECT_ROOT = Path(__file__).resolve().parent.parent
TOKEN_FILE = PROJECT_ROOT / "09_system" / "config" / "google_oauth_token.pickle"
CLIENT_FILE = PROJECT_ROOT / "09_system" / "config" / "google_oauth_client.json"
SPREADSHEET_ID = "1EHKgmE1d7T5N9GbT_QJV72Q8Dh91iNysHHkz6LEWJy4"

ROSTER_PATH = PROJECT_ROOT / "03_clients" / "SIFTAI" / "プロプレミアムTEAM" / "会員管理" / "member_roster.md"
KPI_PATH = PROJECT_ROOT / "03_clients" / "SIFTAI" / "プロプレミアムTEAM" / "コーチング" / "kpi_progress.md"

TODAY = datetime.now()
TODAY_STR = TODAY.strftime("%Y-%m-%d")

DRY_RUN = len(sys.argv) > 1 and sys.argv[1] == "dry-run"


# === Google Sheets認証 ===

def get_sheets_service():
    """OAuth認証でSheets APIサービスを取得"""
    with open(TOKEN_FILE, 'rb') as f:
        creds = pickle.load(f)

    if creds.expired:
        creds.refresh(Request())
        with open(TOKEN_FILE, 'wb') as f:
            pickle.dump(creds, f)
        print("トークンを更新しました")

    return build('sheets', 'v4', credentials=creds)


def fetch_sheet(service, sheet_name, range_cols="A1:Z500"):
    """指定シートのデータを取得"""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!{range_cols}"
    ).execute()
    rows = result.get('values', [])
    if not rows:
        return [], []
    return rows[0], rows[1:]


# === データ取得 ===

def fetch_all_data(service):
    """カルテから全データを取得"""
    print("データ取得中...")

    pro_header, pro_rows = fetch_sheet(service, "DB_Pro-Members", "A1:J100")
    print(f"  PRO会員: {len(pro_rows)}名")

    pre_header, pre_rows = fetch_sheet(service, "DB_Premium-Members", "A1:J100")
    print(f"  PREMIUM会員: {len(pre_rows)}名")

    wr_header, wr_rows = fetch_sheet(service, "DB_Weekly-Report", "A1:K500")
    print(f"  週報: {len(wr_rows)}件")

    return {
        'pro': {'header': pro_header, 'rows': pro_rows},
        'premium': {'header': pre_header, 'rows': pre_rows},
        'weekly': {'header': wr_header, 'rows': wr_rows},
    }


# === 会員データ構造化 ===

def get_col(row, idx, default=''):
    """安全にカラム値を取得"""
    return row[idx] if len(row) > idx else default


def parse_pro_members(data):
    """PRO会員をパース"""
    members = []
    for row in data['rows']:
        members.append({
            'id': get_col(row, 0),
            'name': get_col(row, 1),
            'start': get_col(row, 4),
            'end': get_col(row, 5),
            'kisu': get_col(row, 6),
            'nippo': int(get_col(row, 7, '0') or '0'),
            'shuho': int(get_col(row, 8, '0') or '0'),
            'fb': int(get_col(row, 9, '0') or '0'),
            'plan': 'PRO',
        })
    return members


def parse_premium_members(data):
    """PREMIUM会員をパース"""
    members = []
    for row in data['rows']:
        members.append({
            'id': get_col(row, 0),
            'name': get_col(row, 1),
            'start': get_col(row, 5),
            'end': get_col(row, 6),
            'nippo': int(get_col(row, 7, '0') or '0'),
            'plan': 'PREMIUM',
        })
    return members


def find_dual_members(pro_list, pre_list):
    """PRO+PREMIUM併用者を特定"""
    pro_ids = {m['id'] for m in pro_list}
    pre_ids = {m['id'] for m in pre_list}
    return pro_ids & pre_ids


def classify_status(member):
    """アクティブ/非アクティブを判定"""
    nippo = member['nippo']
    fb = member.get('fb', 0)
    start = member.get('start', '')

    # 未開始判定
    if start and start > TODAY.strftime("%Y/%m/%d"):
        return '未開始'

    # 要緊急フォロー: 累計日報2以下
    if nippo <= 2:
        return '要緊急フォロー'

    # 非アクティブ: FB会0（PRO会員のみ）
    if member['plan'] == 'PRO' and fb == 0:
        return '非アクティブ'

    # 要注意: 日報少なめ
    if nippo < 10:
        return '要注意'

    return 'アクティブ'


def parse_monetize(data):
    """週報からマネタイズデータを抽出"""
    results = []
    for row in data['rows']:
        name = get_col(row, 1)
        date = get_col(row, 2)
        hours = get_col(row, 3)
        phase = get_col(row, 5)
        monetize_text = get_col(row, 9)

        if not monetize_text.strip():
            continue
        if name == 'テスト':
            continue

        # 金額抽出（数字+円 or カンマ区切り数字）
        import re
        amount = 0
        # 「16,000円」「15000円」「300円」「209円」「40円」パターン
        match = re.search(r'([\d,]+)\s*円', monetize_text)
        if match:
            amount = int(match.group(1).replace(',', ''))

        results.append({
            'name': name,
            'date': date,
            'hours': hours,
            'phase': phase[:30] if phase else '',
            'monetize_text': monetize_text[:200],
            'amount': amount,
        })

    return results


# === MD生成 ===

def generate_roster_md(pro_list, pre_list, dual_ids):
    """member_roster.md を生成"""
    lines = []
    lines.append("# 会員ロースター")
    lines.append("")
    lines.append(f"最終更新：{TODAY_STR}（Google Sheets カルテより自動同期）")
    lines.append("")
    lines.append(f"データソース: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/")
    lines.append("")
    lines.append("---")
    lines.append("")

    # PRO会員
    lines.append(f"## PRO会員（{len(pro_list)}名）")
    lines.append("")
    lines.append("| # | ID | 名前 | 期数 | 開始日 | 終了日 | 日報 | 週報 | FB会 | PRE併用 | 状態 |")
    lines.append("|---|---|---|---|---|---|---:|---:|---:|:---:|---|")

    for i, m in enumerate(pro_list, 1):
        m['status'] = classify_status(m)
        dual = '○' if m['id'] in dual_ids else '—'
        lines.append(f"| {i} | {m['id']} | {m['name']} | {m['kisu']} | {m['start']} | {m['end']} | {m['nippo']} | {m['shuho']} | {m['fb']} | {dual} | {m['status']} |")

    lines.append("")
    lines.append("---")
    lines.append("")

    # PREMIUM会員
    lines.append(f"## PREMIUM会員（{len(pre_list)}名）")
    lines.append("")
    lines.append("| # | ID | 名前 | 開始日 | 終了日 | 日報 | PRO併用 | 状態 |")
    lines.append("|---|---|---|---|---|---:|:---:|---|")

    for i, m in enumerate(pre_list, 1):
        m['status'] = classify_status(m)
        dual = '○' if m['id'] in dual_ids else '—'
        lines.append(f"| {i} | {m['id']} | {m['name']} | {m['start']} | {m['end']} | {m['nippo']} | {dual} | {m['status']} |")

    lines.append("")
    lines.append("---")
    lines.append("")

    # PRO/PREMIUM併用
    dual_names = [m['name'] for m in pro_list if m['id'] in dual_ids]
    lines.append("## PRO/PREMIUM 併用メンバー")
    lines.append("")
    if dual_names:
        for name in dual_names:
            lines.append(f"- {name}")
    else:
        lines.append("（なし）")

    lines.append("")
    lines.append("---")
    lines.append("")

    # サマリー
    pro_active = sum(1 for m in pro_list if m.get('status') == 'アクティブ')
    pro_warn = sum(1 for m in pro_list if m.get('status') == '要注意')
    pro_urgent = sum(1 for m in pro_list if m.get('status') == '要緊急フォロー')
    pro_new = sum(1 for m in pro_list if m.get('status') == '未開始')
    pre_active = sum(1 for m in pre_list if m.get('status') == 'アクティブ')
    pre_warn = sum(1 for m in pre_list if m.get('status') == '要注意')
    pre_urgent = sum(1 for m in pre_list if m.get('status') == '要緊急フォロー')

    lines.append("## サマリー")
    lines.append("")
    lines.append("| 指標 | PRO | PREMIUM |")
    lines.append("|---|---|---|")
    lines.append(f"| 総在籍数 | {len(pro_list)}名 | {len(pre_list)}名 |")
    lines.append(f"| アクティブ | {pro_active}名 | {pre_active}名 |")
    lines.append(f"| 要注意 | {pro_warn}名 | {pre_warn}名 |")
    lines.append(f"| 要緊急フォロー | {pro_urgent}名 | {pre_urgent}名 |")
    if pro_new:
        lines.append(f"| 未開始 | {pro_new}名 | — |")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## 判定基準")
    lines.append("")
    lines.append("- **アクティブ**: 累計日報10回以上 かつ FB会参加あり")
    lines.append("- **要注意**: 累計日報3〜9回 または FB会ゼロ")
    lines.append("- **要緊急フォロー**: 累計日報2回以下")
    lines.append("- **未開始**: 開始日が未到来")

    return '\n'.join(lines) + '\n'


def generate_kpi_md(pro_list, pre_list, dual_ids, monetize_data):
    """kpi_progress.md を生成"""
    lines = []
    lines.append("# 会員KPI進捗ダッシュボード")
    lines.append("")
    lines.append(f"最終更新：{TODAY_STR}（Google Sheets カルテより自動同期）")
    lines.append("")
    lines.append(f"データソース: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/")
    lines.append("")
    lines.append("---")
    lines.append("")

    # 名前→プラン マッピング
    name_plan = {}
    for m in pro_list:
        name_plan[m['name']] = 'PRO+PRE' if m['id'] in dual_ids else 'PRO'
    for m in pre_list:
        if m['name'] not in name_plan:
            name_plan[m['name']] = 'PREMIUM'

    # マネタイズランキング（金額あり）
    ranked = sorted([r for r in monetize_data if r['amount'] > 0], key=lambda x: -x['amount'])

    # 名前ごとに最新のデータを集約
    name_totals = {}
    for r in ranked:
        if r['name'] not in name_totals:
            name_totals[r['name']] = {'amount': 0, 'texts': [], 'phase': r['phase']}
        name_totals[r['name']]['amount'] += r['amount']
        name_totals[r['name']]['texts'].append(r['monetize_text'])

    ranked_unique = sorted(name_totals.items(), key=lambda x: -x[1]['amount'])

    lines.append("## マネタイズ成果ランキング")
    lines.append("")
    lines.append("| 順位 | 名前 | プラン | フェーズ | 累計額 | 内容 |")
    lines.append("|:---:|---|---|---|---:|---|")

    for i, (name, info) in enumerate(ranked_unique, 1):
        plan = name_plan.get(name, '不明')
        text = info['texts'][0][:50] if info['texts'] else ''
        lines.append(f"| {i} | {name} | {plan} | {info['phase']} | {info['amount']:,}円 | {text} |")

    total_amount = sum(v['amount'] for v in name_totals.values())
    lines.append("")
    lines.append(f"合計: **{total_amount:,}円**（{len(ranked_unique)}名）")
    lines.append("")
    lines.append("---")
    lines.append("")

    # 要緊急フォロー
    lines.append("## 要緊急フォロー対象")
    lines.append("")
    lines.append("### PRO")
    lines.append("")
    lines.append("| 名前 | 累計日報 | 累計FB | 状態 |")
    lines.append("|---|---:|---:|---|")
    for m in pro_list:
        if m.get('status') == '要緊急フォロー':
            lines.append(f"| {m['name']} | {m['nippo']} | {m['fb']} | 要緊急フォロー |")

    lines.append("")
    lines.append("### PREMIUM")
    lines.append("")
    lines.append("| 名前 | 累計日報 | 状態 |")
    lines.append("|---|---:|---|")
    for m in pre_list:
        if m.get('status') == '要緊急フォロー':
            lines.append(f"| {m['name']} | {m['nippo']} | 要緊急フォロー |")

    lines.append("")
    lines.append("---")
    lines.append("")

    # 成果事例
    lines.append("## 成果事例（表彰・共有候補）")
    lines.append("")
    lines.append("| 名前 | プラン | 成果内容 |")
    lines.append("|---|---|---|")
    for name, info in ranked_unique[:5]:
        plan = name_plan.get(name, '不明')
        text = info['texts'][0][:60] if info['texts'] else ''
        lines.append(f"| {name} | {plan} | {info['amount']:,}円 — {text} |")

    lines.append("")
    lines.append("---")
    lines.append("")

    # Lv分布
    lines.append("## Lv.分布")
    lines.append("")
    lines.append("| Lv | 人数 |")
    lines.append("|---|---:|")
    lines.append(f"| Lv.1未満 | {len(pro_list) + len(pre_list) - len(dual_ids)}名 |")
    lines.append("| Lv.1〜5 | 0名 |")

    return '\n'.join(lines) + '\n'


# === メイン ===

def main():
    print(f"=== カルテ同期 {TODAY_STR} ===")
    if DRY_RUN:
        print("（dry-runモード：ファイル更新なし）")

    service = get_sheets_service()
    data = fetch_all_data(service)

    pro_list = parse_pro_members(data['pro'])
    pre_list = parse_premium_members(data['premium'])
    dual_ids = find_dual_members(pro_list, pre_list)
    monetize_data = parse_monetize(data['weekly'])

    print(f"\n併用者: {len(dual_ids)}名")
    print(f"マネタイズ報告: {len(monetize_data)}件")

    # マネタイズ金額あり
    with_amount = [r for r in monetize_data if r['amount'] > 0]
    total = sum(r['amount'] for r in with_amount)
    print(f"収益報告: {len(with_amount)}件 / 合計: {total:,}円")

    # MD生成
    roster_md = generate_roster_md(pro_list, pre_list, dual_ids)
    kpi_md = generate_kpi_md(pro_list, pre_list, dual_ids, monetize_data)

    if DRY_RUN:
        print(f"\n--- member_roster.md（{len(roster_md)}文字）---")
        print(roster_md[:500] + "...")
        print(f"\n--- kpi_progress.md（{len(kpi_md)}文字）---")
        print(kpi_md[:500] + "...")
    else:
        ROSTER_PATH.write_text(roster_md, encoding='utf-8')
        print(f"\n✓ {ROSTER_PATH} を更新しました")

        KPI_PATH.write_text(kpi_md, encoding='utf-8')
        print(f"✓ {KPI_PATH} を更新しました")

    print("\n同期完了")


if __name__ == "__main__":
    main()
