"""
毎朝8:00にDiscordへ今日のタスクを配信するスクリプト
- 全クライアントのタスクを収集
- 期限が近いもの・止まっているものを警告表示
- Discord Webhookで送信
"""

import json
import re
import os
import sys
import requests
from datetime import datetime, timedelta
from pathlib import Path

# Windows環境でのUTF-8出力対応
os.environ["PYTHONIOENCODING"] = "utf-8"
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

# === 設定 ===
WEBHOOK_URL = "https://discord.com/api/webhooks/1480602170368856275/DCDEaDwYhIIH0xGioPjanNbTocElv0FxtbnGo7HTqSDfpR_bn_IPO6Rm2PNhC7p0gUB4"
PROJECT_ROOT = Path(__file__).resolve().parent.parent
TODAY = datetime.now()
TODAY_STR = TODAY.strftime("%Y-%m-%d")

# === タスク収集 ===

def collect_sekibiz_tasks():
    """関ビズ タスク管理.htmlからタスクを抽出"""
    tasks = []
    html_path = PROJECT_ROOT / "03_clients" / "関ビズ" / "projects" / "20260423_集客ロードマップ作り講座" / "タスク管理.html"

    if not html_path.exists():
        return tasks

    content = html_path.read_text(encoding="utf-8")

    # カテゴリ（h2タグ内）とタスクアイテム（div.task-item）を解析
    # カテゴリヘッダーのh2からカテゴリ名を抽出
    category_pattern = re.compile(r'<h2>(.*?)</h2>', re.DOTALL)

    # div.category セクションごとに分割
    cat_sections = re.split(r'<div class="category"[^>]*>', content)

    for section in cat_sections[1:]:  # 最初はヘッダー部分なのでスキップ
        # カテゴリ名を取得
        cat_match = category_pattern.search(section)
        category = re.sub(r'<[^>]+>', '', cat_match.group(1)).strip() if cat_match else "その他"

        # 未完了タスク（checkedなし）のtask-itemを抽出
        # task-item checked = 完了済み、task-item = 未完了
        item_pattern = re.compile(
            r'<div class="task-item">\s*'
            r'<input type="checkbox"[^>]*>\s*'
            r'<div class="task-content">\s*'
            r'<div class="task-name">(.*?)</div>',
            re.DOTALL
        )

        for match in item_pattern.finditer(section):
            raw_name = match.group(1)
            # タスク名からHTMLタグを除去して整形
            task_name = re.sub(r'<[^>]+>', '', raw_name).strip()
            # 「未着手」「完了」などのステータスタグを除去
            task_name = re.sub(r'\s*(未着手|完了)\s*', '', task_name).strip()
            # 期限を抽出
            deadline_match = re.search(r'期限:\s*(\d{1,2}/\d{1,2})', raw_name)
            deadline = deadline_match.group(1) if deadline_match else ""
            # 期限テキストをタスク名から除去
            task_name = re.sub(r'\s*期限:\s*\d{1,2}/\d{1,2}\s*', '', task_name).strip()

            tasks.append({
                "client": "関ビズ",
                "project": "4/23 集客ロードマップ講座",
                "category": category,
                "task": task_name,
                "deadline": deadline,
                "completed": False
            })

    return tasks


def collect_siftai_tasks():
    """SIFTAIの未確認事項・タスクを収集"""
    tasks = []

    # report_analyzer.md の未確認事項
    analyzer_path = PROJECT_ROOT / "03_clients" / "SIFTAI" / "プロプレミアムTEAM" / "AIエージェント" / "report_analyzer.md"
    if analyzer_path.exists():
        content = analyzer_path.read_text(encoding="utf-8")
        # [ ] パターンで未完了タスクを抽出
        for match in re.finditer(r'- \[ \]\s*(.+)', content):
            tasks.append({
                "client": "SIFTAI",
                "project": "プロプレミアムTEAM",
                "category": "AIエージェント設計",
                "task": match.group(1).strip(),
                "deadline": "",
                "completed": False
            })

    return tasks


def collect_bcbg_tasks():
    """B.C.B.Gのタスク管理MDから未完了タスクを収集"""
    tasks = []

    # projects配下の全プロジェクトからタスク管理.mdを探す
    bcbg_dir = PROJECT_ROOT / "03_clients" / "B.C.B.G" / "projects"
    if not bcbg_dir.exists():
        return tasks

    for project_dir in bcbg_dir.iterdir():
        if not project_dir.is_dir():
            continue

        task_file = project_dir / "タスク管理.md"
        if not task_file.exists():
            continue

        content = task_file.read_text(encoding="utf-8")
        project_name = project_dir.name
        current_category = "その他"

        for line in content.split('\n'):
            line = line.strip()
            # カテゴリ（## 見出し）を取得
            if line.startswith('## '):
                current_category = line[3:].strip()
                continue
            # 未完了タスク（- [ ]）を取得
            match = re.match(r'- \[ \]\s*(.+)', line)
            if match:
                tasks.append({
                    "client": "B.C.B.G",
                    "project": project_name,
                    "category": current_category,
                    "task": match.group(1).strip(),
                    "deadline": "",
                    "completed": False
                })

    return tasks


def collect_journal_tasks():
    """最新ジャーナルの申し送り事項を収集"""
    tasks = []

    # 最新のジャーナルファイルを探す
    journal_dir = PROJECT_ROOT / "04_journal"
    latest_file = None

    for month_dir in sorted(journal_dir.glob("????-??"), reverse=True):
        for day_file in sorted(month_dir.glob("????-??-??.md"), reverse=True):
            latest_file = day_file
            break
        if latest_file:
            break

    if not latest_file:
        return tasks

    content = latest_file.read_text(encoding="utf-8")

    # 「次回への申し送り」セクションを抽出
    handover_match = re.search(r'## 次回への申し送り\n(.*?)(?:\n## |\Z)', content, re.DOTALL)
    if handover_match:
        for line in handover_match.group(1).strip().split('\n'):
            line = line.strip()
            if line.startswith('- '):
                task_text = line[2:].strip()
                # **太字**の期限を抽出
                deadline_match = re.search(r'[（(](\d{1,2}/\d{1,2})[期限）)]', task_text)
                deadline = deadline_match.group(1) if deadline_match else ""
                # マークダウン記法を除去
                clean_text = re.sub(r'\*\*(.+?)\*\*', r'\1', task_text)

                tasks.append({
                    "client": "全体",
                    "project": "申し送り",
                    "category": "ジャーナル",
                    "task": clean_text,
                    "deadline": deadline,
                    "completed": False
                })

    return tasks


def check_stalled(task):
    """タスクが止まっているかを判定"""
    deadline = task.get("deadline", "")
    if not deadline:
        return False

    # 期限をパース（4/17 形式）
    try:
        match = re.match(r'(\d{1,2})/(\d{1,2})', deadline)
        if match:
            month, day = int(match.group(1)), int(match.group(2))
            deadline_date = datetime(TODAY.year, month, day)
            # 期限が3日以内、または過ぎている場合は「止まっている」
            if deadline_date <= TODAY + timedelta(days=3):
                return True
    except (ValueError, AttributeError):
        pass

    return False


def format_discord_message(all_tasks):
    """Discord送信用のメッセージを整形"""

    # 止まっているタスクと通常タスクを分離
    stalled = [t for t in all_tasks if check_stalled(t)]
    normal = [t for t in all_tasks if not check_stalled(t)]

    lines = []
    lines.append(f"# 📋 SAEPIN日報 — {TODAY.strftime('%Y/%m/%d (%a)')}")
    lines.append("")

    # ⚠️ 止まっているタスク（警告）
    if stalled:
        lines.append("## ⚠️ 注意：期限が迫っている / 止まっているタスク")
        for t in stalled:
            deadline_str = f" 🔴 期限: {t['deadline']}" if t['deadline'] else ""
            lines.append(f"- **[{t['client']}]** {t['task']}{deadline_str}")
        lines.append("")

    # クライアント別に整理
    clients = {}
    for t in normal:
        key = t["client"]
        if key not in clients:
            clients[key] = {}
        cat = t["category"]
        if cat not in clients[key]:
            clients[key][cat] = []
        clients[key][cat].append(t)

    if clients:
        lines.append("## 📌 今日のタスク一覧")
        for client, categories in clients.items():
            lines.append(f"\n### 🏢 {client}")
            for cat, tasks in categories.items():
                lines.append(f"**{cat}**")
                for t in tasks:
                    deadline_str = f" （期限: {t['deadline']}）" if t['deadline'] else ""
                    lines.append(f"- [ ] {t['task']}{deadline_str}")
        lines.append("")

    # 未完了タスク数のサマリー
    total = len(all_tasks)
    stalled_count = len(stalled)
    lines.append(f"---")
    lines.append(f"📊 未完了タスク: **{total}件** ｜ ⚠️ 要注意: **{stalled_count}件**")
    lines.append(f"\n> 💬 COO SAEPINより：社長、今日もよろしくお願いします！")

    return "\n".join(lines)


def send_to_discord(message):
    """DiscordのWebhookにメッセージを送信"""
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
        response = requests.post(WEBHOOK_URL, json=payload)
        if response.status_code == 204:
            print(f"✅ Discord送信成功")
        else:
            print(f"❌ Discord送信失敗: {response.status_code} - {response.text}")


def main():
    print(f"📋 タスク収集開始... ({TODAY_STR})")

    # 全ソースからタスクを収集
    all_tasks = []
    all_tasks.extend(collect_sekibiz_tasks())
    all_tasks.extend(collect_siftai_tasks())
    all_tasks.extend(collect_bcbg_tasks())
    all_tasks.extend(collect_journal_tasks())

    print(f"  関ビズ: {len([t for t in all_tasks if t['client'] == '関ビズ'])}件")
    print(f"  SIFTAI: {len([t for t in all_tasks if t['client'] == 'SIFTAI'])}件")
    print(f"  B.C.B.G: {len([t for t in all_tasks if t['client'] == 'B.C.B.G'])}件")
    print(f"  申し送り: {len([t for t in all_tasks if t['client'] == '全体'])}件")
    print(f"  合計: {len(all_tasks)}件")

    if not all_tasks:
        message = f"# 📋 SAEPIN日報 — {TODAY.strftime('%Y/%m/%d (%a)')}\n\n✨ 未完了タスクはありません！お疲れ様です、社長！"
    else:
        message = format_discord_message(all_tasks)

    print("\n--- 送信内容 ---")
    print(message)
    print("--- ここまで ---\n")

    send_to_discord(message)

    # LINE にも同時配信
    from line_notify import send_to_line
    send_to_line(message)


if __name__ == "__main__":
    main()
