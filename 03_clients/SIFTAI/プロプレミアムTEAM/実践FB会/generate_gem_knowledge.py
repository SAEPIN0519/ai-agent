# 実践FB会 Gem用ナレッジベース生成スクリプト（NotebookLM連携版）
# Google Sheets + NotebookLM CLIからデータを取得し、Gemにアップロードするテキストファイルを生成する
#
# 使い方:
#   python generate_gem_knowledge.py
#
# 前提:
#   - notebooklm CLI がインストール済み（notebooklm login 済み）
#   - Google Sheets サービスアカウント認証済み
#
# 出力:
#   実践FB会_ナレッジベース.txt — Gemにアップロードするファイル

import sys
import os
import json
import subprocess

sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, '..', '..', '..', '..'))
CONFIG_DIR = os.path.join(PROJECT_ROOT, '09_system', 'config')

import gspread
from google.oauth2.service_account import Credentials

SPREADSHEET_ID = '12aAhIo5WMjd1erPMYHrSBoQF2s0heuJtdpy5kUMR2Vs'
SA_JSON = os.path.join(CONFIG_DIR, 'google_service_account.json')
OUTPUT_FILE = os.path.join(SCRIPT_DIR, '実践FB会_ナレッジベース.txt')

# NotebookLMノートブックID
NOTEBOOK_ID = '7620e75a-3a5c-4943-9be1-103348528422'


# ==================== Google Sheets ====================

def get_sheets_data():
    """Google Sheetsからデータを取得"""
    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds = Credentials.from_service_account_file(SA_JSON, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    return sh


def build_session_knowledge(sh):
    """（案）検索サイトシートからセッション情報を構造化"""
    ws = sh.worksheet('（案）検索サイト')
    rows = ws.get_all_values()
    sessions = []
    for row in rows[1:]:
        if len(row) < 6 or not row[0]:
            continue
        sessions.append({
            '日付': row[0],
            'タイトル': row[1].replace('\n', ' / '),
            'カテゴリ': row[2],
            'タグ': row[3],
            '概要': row[4],
            'URL': row[5],
        })
    return sessions


def build_task_schedule(sh):
    """タスク管理シートからスケジュール情報を取得"""
    ws = sh.worksheet('タスク管理')
    rows = ws.get_all_values()
    schedule = []
    for row in rows[2:]:
        if len(row) < 7 or not row[1]:
            continue
        schedule.append({
            '日付': row[1],
            '開始': row[2],
            '終了': row[3],
            'カテゴリ': row[4],
            'タイトル': row[5].replace('\n', ' / '),
            '講師': row[6],
        })
    return schedule


# ==================== NotebookLM ====================

def get_notebooklm_sources():
    """NotebookLM CLIでソース一覧を取得"""
    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUTF8'] = '1'

    # ノートブックを選択
    subprocess.run(
        ['notebooklm', 'use', NOTEBOOK_ID],
        capture_output=True, text=True, encoding='utf-8', env=env, timeout=30
    )

    # ソース一覧をJSON取得
    result = subprocess.run(
        ['notebooklm', 'source', 'list', '--json'],
        capture_output=True, text=True, encoding='utf-8', env=env, timeout=30
    )
    if result.returncode != 0:
        print(f'  警告: ソース一覧取得失敗 - {result.stderr[:100]}')
        return []

    data = json.loads(result.stdout)
    return data.get('sources', [])


def get_source_fulltext(source_id, env):
    """ソースのフルテキストを取得"""
    try:
        result = subprocess.run(
            ['notebooklm', 'source', 'fulltext', source_id],
            capture_output=True, text=True, encoding='utf-8', env=env, timeout=60
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout
    except Exception as e:
        print(f'  エラー: {e}')
    return None


def extract_content(text, max_chars=5000):
    """まとめ・詳細セクションを抽出"""
    lines = text.split('\n')
    summary_lines = []
    detail_lines = []
    in_summary = False
    in_detail = False

    for line in lines:
        stripped = line.strip()
        if stripped == 'まとめ':
            in_summary = True
            in_detail = False
            continue
        if stripped == '詳細':
            in_summary = False
            in_detail = True
            continue
        if in_summary:
            if stripped and stripped != 'Roboto':
                summary_lines.append(stripped)
        elif in_detail:
            if stripped and stripped != 'Roboto':
                detail_lines.append(stripped)

    summary = '\n'.join(summary_lines)

    if len(summary) > 50:
        detail_excerpt = '\n'.join(detail_lines[:50])
        combined = summary + '\n\n【詳細抜粋】\n' + detail_excerpt
        return combined[:max_chars]
    else:
        detail = '\n'.join(detail_lines[:80])
        return detail[:max_chars]


def fetch_all_source_contents(sources):
    """全ソースのフルテキストを取得してコンテンツを抽出"""
    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUTF8'] = '1'

    contents = []
    for i, src in enumerate(sources):
        sid = src['id']
        title = src['title']
        created = src.get('created_at', '')
        print(f'  [{i+1}/{len(sources)}] {title[:50]}...', flush=True)

        fulltext = get_source_fulltext(sid, env)
        if fulltext:
            extracted = extract_content(fulltext)
            if extracted:
                contents.append({
                    'title': title,
                    'created': created,
                    'content': extracted,
                })

    return contents


# ==================== ナレッジベース生成 ====================

def generate_knowledge_text(sessions, schedule, nb_contents):
    """Gem用のナレッジベーステキストを生成"""
    lines = []

    lines.append('=' * 60)
    lines.append('SHIFT AI PRO・PREMIUM 実践フィードバック会')
    lines.append('完全ナレッジベース（NotebookLM連携版）')
    lines.append('=' * 60)
    lines.append('')

    # セッション一覧
    lines.append(f'■ セッション一覧（全{len(sessions)}回）')
    lines.append('-' * 40)
    for s in sessions:
        lines.append(f'【{s["日付"]}】{s["タイトル"]}')
        lines.append(f'  カテゴリ: {s["カテゴリ"]}')
        if s['タグ']:
            lines.append(f'  タグ: {s["タグ"]}')
        lines.append(f'  概要: {s["概要"]}')
        lines.append(f'  アーカイブURL: {s["URL"]}')
        lines.append('')

    # カテゴリ別まとめ
    lines.append('')
    lines.append('■ カテゴリ別セッション一覧')
    lines.append('-' * 40)
    categories = {}
    for s in sessions:
        cat = s['カテゴリ']
        if cat not in categories:
            categories[cat] = []
        categories[cat].append(s)
    for cat, items in categories.items():
        lines.append(f'\n▼ {cat}（全{len(items)}回）')
        for s in items:
            lines.append(f'  - {s["日付"]} | {s["タイトル"]}')

    # NotebookLMソースからの内容
    if nb_contents:
        lines.append('')
        lines.append('')
        lines.append('=' * 60)
        lines.append(f'■ セッション内容詳細（NotebookLMソースより・{len(nb_contents)}件）')
        lines.append('=' * 60)

        # 作成日でソート
        nb_contents.sort(key=lambda x: x.get('created', ''))
        for item in nb_contents:
            lines.append(f'\n--- {item["title"]} ---')
            lines.append(f'日時: {item["created"][:10] if item["created"] else "不明"}')
            lines.append(item['content'])
            lines.append('')

    # スケジュール
    lines.append('')
    lines.append('■ 実施スケジュール（講師情報付き）')
    lines.append('-' * 40)
    for s in schedule:
        lines.append(f'{s["日付"]} {s["開始"]}-{s["終了"]} | {s["カテゴリ"]} | {s["タイトル"]} | 講師: {s["講師"]}')

    return '\n'.join(lines)


def main():
    # Google Sheets
    print('Google Sheetsからデータを取得中...')
    sh = get_sheets_data()

    print('セッション情報を取得中...')
    sessions = build_session_knowledge(sh)
    print(f'  → {len(sessions)}件のセッション')

    print('スケジュール情報を取得中...')
    schedule = build_task_schedule(sh)
    print(f'  → {len(schedule)}件のスケジュール')

    # NotebookLM
    print('NotebookLMからソースを取得中...')
    sources = get_notebooklm_sources()
    print(f'  → {len(sources)}件のソース')

    nb_contents = []
    if sources:
        print('各ソースのフルテキストを取得中...')
        nb_contents = fetch_all_source_contents(sources)
        print(f'  → {len(nb_contents)}件のコンテンツ取得完了')
    else:
        print('  警告: NotebookLMソースが取得できなかったため、Sheetsデータのみで生成します')

    # ナレッジベース生成
    print('ナレッジベースを生成中...')
    text = generate_knowledge_text(sessions, schedule, nb_contents)

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(text)

    print(f'\n完了！ 出力先: {OUTPUT_FILE}')
    print(f'ファイルサイズ: {len(text):,} 文字')
    print(f'  Sheetsセッション: {len(sessions)}件')
    print(f'  NotebookLMコンテンツ: {len(nb_contents)}件')
    print(f'  スケジュール: {len(schedule)}件')


if __name__ == '__main__':
    main()
