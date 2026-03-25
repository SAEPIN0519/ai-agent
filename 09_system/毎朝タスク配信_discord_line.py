"""
毎朝8:00にDiscord + KOKONOE管理者LINEへタスクを配信するスクリプト
- SIFTAI: Slackからメンション・タスク + ACTION_TRACKER
- SNK: 進捗管理ダッシュボードから遅延・問題を抽出
- AIsoraCode: 冴香担当タスク
- B.C.B.G: マルシェタスク
- 関ビズ: タスク管理HTMLから未完了を抽出
- ジャーナル: 申し送り事項
"""

import json
import re
import os
import sys
import urllib.request
import urllib.error
from datetime import datetime, timedelta
from pathlib import Path

# Windows環境でのUTF-8出力対応
os.environ["PYTHONIOENCODING"] = "utf-8"
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

# === 設定 ===
DISCORD_WEBHOOK = "https://discord.com/api/webhooks/1480602170368856275/DCDEaDwYhIIH0xGioPjanNbTocElv0FxtbnGo7HTqSDfpR_bn_IPO6Rm2PNhC7p0gUB4"
PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = PROJECT_ROOT / "09_system" / "config"
TODAY = datetime.now()
TODAY_STR = TODAY.strftime("%Y-%m-%d")

# Slack設定
SLACK_USER_TOKEN = (CONFIG_DIR / "slack_user_token.txt").read_text(encoding="utf-8").strip() if (CONFIG_DIR / "slack_user_token.txt").exists() else ""
SAEKA_SLACK_ID = "U0A00UNQV1V"

# LINE設定
LINE_TOKEN = (CONFIG_DIR / "line_kokonoe_access_token.txt").read_text(encoding="utf-8").strip() if (CONFIG_DIR / "line_kokonoe_access_token.txt").exists() else ""

# 監視するSlackチャンネル（SIFTAI）
SLACK_CHANNELS = [
    ("C07CPL6FHFT", "all-report_community"),
    ("C05AD7NHYU9", "community"),
    ("C077WLGN2T0", "daily-todo"),
]


# === Slack読み込み ===

def fetch_slack_mentions():
    """Slackから冴香さん宛てのメンション・タスクを取得（直近24時間）"""
    if not SLACK_USER_TOKEN:
        return []

    tasks = []
    yesterday_ts = str((TODAY - timedelta(days=1)).timestamp())

    for ch_id, ch_name in SLACK_CHANNELS:
        try:
            url = f"https://slack.com/api/conversations.history?channel={ch_id}&oldest={yesterday_ts}&limit=50"
            req = urllib.request.Request(url, headers={"Authorization": f"Bearer {SLACK_USER_TOKEN}"})
            with urllib.request.urlopen(req) as res:
                data = json.loads(res.read().decode("utf-8"))
                if not data.get("ok"):
                    continue

                for msg in data.get("messages", []):
                    text = msg.get("text", "")
                    # 冴香さんへのメンションを含むメッセージ
                    if f"<@{SAEKA_SLACK_ID}>" in text or "野村" in text or "冴香" in text:
                        # メンションタグを除去して整形
                        clean = re.sub(r"<@[A-Z0-9]+>", "", text).strip()
                        clean = re.sub(r"<[^>]+\|([^>]+)>", r"\1", clean)  # リンク整形
                        clean = re.sub(r"<[^>]+>", "", clean)
                        if clean and len(clean) > 5:
                            tasks.append({
                                "source": f"Slack#{ch_name}",
                                "text": clean[:200],
                            })
        except Exception as e:
            print(f"  Slack {ch_name}: {e}")

    return tasks


# === SIFTAI タスク収集 ===

def collect_siftai_tasks():
    """SIFTAI ACTION_TRACKERから冴香の未完了タスクを収集"""
    tasks = []

    tracker_path = PROJECT_ROOT / "03_clients" / "SIFTAI" / "プロプレミアムTEAM" / "MTG議事録" / "ACTION_TRACKER.md"
    if not tracker_path.exists():
        return tasks

    content = tracker_path.read_text(encoding="utf-8")

    # 野村冴香セクションを抽出
    saeka_match = re.search(r"### 野村冴香\n(.*?)(?=\n###|\n---|\Z)", content, re.DOTALL)
    if saeka_match:
        # テーブル行からタスクを抽出
        for line in saeka_match.group(1).strip().split("\n"):
            if "|" in line and "未着手" in line:
                cols = [c.strip() for c in line.split("|")]
                if len(cols) >= 6:
                    task_name = cols[3] if len(cols) > 3 else ""
                    deadline = cols[4] if len(cols) > 4 else ""
                    if task_name and task_name != "やること":
                        tasks.append({
                            "client": "SIFTAI",
                            "category": "ProプランMTG",
                            "task": task_name,
                            "deadline": deadline,
                            "status": "未着手",
                        })

    # report_analyzer.mdの未完了タスク
    analyzer_path = PROJECT_ROOT / "03_clients" / "SIFTAI" / "プロプレミアムTEAM" / "AIエージェント" / "report_analyzer.md"
    if analyzer_path.exists():
        content = analyzer_path.read_text(encoding="utf-8")
        for match in re.finditer(r"- \[ \]\s*(.+)", content):
            tasks.append({
                "client": "SIFTAI",
                "category": "AIエージェント設計",
                "task": match.group(1).strip(),
                "deadline": "",
                "status": "未着手",
            })

    return tasks


# === SNK タスク収集 ===

def collect_snk_issues():
    """SNK進捗管理から遅延・問題を抽出"""
    issues = []

    json_path = PROJECT_ROOT / "03_clients" / "SNK" / "更新型進捗管理" / "_dashboard_data.json"
    if not json_path.exists():
        return issues, {}

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # 統計
    stats = {
        "total": len(data),
        "delayed": 0,
        "by_tantou": {},
    }

    # 遅延案件を抽出
    delayed_items = []
    for item in data:
        okure = item.get("okure", 0)
        if okure > 0:
            stats["delayed"] += 1
            tantou = item.get("tantou", "不明")
            stats["by_tantou"][tantou] = stats["by_tantou"].get(tantou, 0) + 1

            # 直近の遅延のみ報告（30日以内）
            if okure <= 30:
                delayed_items.append({
                    "tantou": tantou,
                    "hinmei": item.get("hinmei", "").replace("\r\n", " ")[:30],
                    "status": item.get("status", ""),
                    "okure": okure,
                    "kyakusaki": item.get("kyakusaki", ""),
                })

    # 遅延日数でソート
    delayed_items.sort(key=lambda x: -x["okure"])

    return delayed_items[:10], stats


# === AIsoraCode タスク収集 ===

def collect_aisoracode_tasks():
    """ジャーナルからAIsoraCodeのタスク（冴香担当分）を収集"""
    return [
        {"task": "AIsoraCode単独 or KOKONOE統合を検討 → 税理士に相談", "deadline": "4/1"},
        {"task": "オープンチャットを準備する", "deadline": "4/1"},
        {"task": "マネタイズ後の報酬分配ルールを決める（AOSYと共同）", "deadline": "4/10"},
        {"task": "モニタリング対象者と価格を決める", "deadline": "未定"},
        {"task": "スタート時期・モニタリング開始日・ゴールを決める（AOSYと共同）", "deadline": "3/30"},
    ]


# === B.C.B.G タスク収集 ===

def collect_bcbg_tasks():
    """B.C.B.Gマルシェのタスク管理MDから未完了を収集"""
    tasks = []

    task_path = PROJECT_ROOT / "03_clients" / "B.C.B.G" / "projects" / "marche-202504" / "タスク管理.md"
    if not task_path.exists():
        return tasks

    content = task_path.read_text(encoding="utf-8")
    current_category = ""

    for line in content.split("\n"):
        line = line.strip()
        if line.startswith("## "):
            current_category = line[3:].strip()
        elif line.startswith("- [ ]"):
            task_text = line[5:].strip()
            tasks.append({
                "category": current_category,
                "task": task_text,
            })

    return tasks


# === 関ビズ タスク収集 ===

def collect_sekibiz_tasks():
    """関ビズ タスク管理.htmlからタスクを抽出"""
    tasks = []
    html_path = PROJECT_ROOT / "03_clients" / "関ビズ" / "projects" / "20260423_集客ロードマップ作り講座" / "タスク管理.html"

    if not html_path.exists():
        return tasks

    content = html_path.read_text(encoding="utf-8")

    cat_sections = re.split(r'<div class="category"[^>]*>', content)
    category_pattern = re.compile(r"<h2>(.*?)</h2>", re.DOTALL)

    for section in cat_sections[1:]:
        cat_match = category_pattern.search(section)
        category = re.sub(r"<[^>]+>", "", cat_match.group(1)).strip() if cat_match else "その他"

        item_pattern = re.compile(
            r'<div class="task-item">\s*'
            r'<input type="checkbox"[^>]*>\s*'
            r'<div class="task-content">\s*'
            r'<div class="task-name">(.*?)</div>',
            re.DOTALL,
        )

        for match in item_pattern.finditer(section):
            raw_name = match.group(1)
            task_name = re.sub(r"<[^>]+>", "", raw_name).strip()
            task_name = re.sub(r"\s*(未着手|完了)\s*", "", task_name).strip()
            deadline_match = re.search(r"期限:\s*(\d{1,2}/\d{1,2})", raw_name)
            deadline = deadline_match.group(1) if deadline_match else ""
            task_name = re.sub(r"\s*期限:\s*\d{1,2}/\d{1,2}\s*", "", task_name).strip()

            if task_name:
                tasks.append({
                    "category": category,
                    "task": task_name,
                    "deadline": deadline,
                })

    return tasks


# === ジャーナル申し送り ===

def collect_journal_handover():
    """最新ジャーナルの申し送り事項を収集"""
    items = []

    journal_dir = PROJECT_ROOT / "04_journal"
    latest_file = None

    for month_dir in sorted(journal_dir.glob("????-??"), reverse=True):
        for day_file in sorted(month_dir.glob("????-??-??.md"), reverse=True):
            latest_file = day_file
            break
        if latest_file:
            break

    if not latest_file:
        return items

    content = latest_file.read_text(encoding="utf-8")

    handover_match = re.search(r"## 次回への申し送り\n(.*?)(?:\n## |\Z)", content, re.DOTALL)
    if handover_match:
        for line in handover_match.group(1).strip().split("\n"):
            line = line.strip()
            if line.startswith("- "):
                items.append(line[2:].strip())

    return items


# === メッセージ整形 ===

def format_message():
    """全タスクを収集してメッセージを整形"""
    lines = []
    weekday_jp = ["月", "火", "水", "木", "金", "土", "日"]
    day_name = weekday_jp[TODAY.weekday()]
    lines.append(f"SAEPIN日報 {TODAY.strftime('%Y/%m/%d')}({day_name})")
    lines.append("")

    # --- SIFTAI ---
    siftai_tasks = collect_siftai_tasks()
    slack_mentions = fetch_slack_mentions()

    if siftai_tasks or slack_mentions:
        lines.append("━━ SIFTAI ━━")
        if siftai_tasks:
            for t in siftai_tasks:
                dl = f" [{t['deadline']}]" if t.get("deadline") else ""
                lines.append(f"- {t['task']}{dl}")
        if slack_mentions:
            lines.append("")
            lines.append("[Slack新着（冴香さん宛て）]")
            for m in slack_mentions[:5]:
                lines.append(f"- #{m['source'].split('#')[1]}: {m['text'][:80]}")
        lines.append("")

    # --- AIsoraCode ---
    aisora_tasks = collect_aisoracode_tasks()
    if aisora_tasks:
        lines.append("━━ AIsoraCode ━━")
        for t in sorted(aisora_tasks, key=lambda x: x.get("deadline", "z")):
            lines.append(f"- {t['task']} [{t['deadline']}]")
        lines.append("")

    # --- B.C.B.G ---
    bcbg_tasks = collect_bcbg_tasks()
    if bcbg_tasks:
        lines.append("━━ B.C.B.G マルシェ(4/19) ━━")
        current_cat = ""
        for t in bcbg_tasks:
            if t["category"] != current_cat:
                current_cat = t["category"]
                lines.append(f"[{current_cat}]")
            lines.append(f"- {t['task']}")
        lines.append("")

    # --- 関ビズ ---
    sekibiz_tasks = collect_sekibiz_tasks()
    if sekibiz_tasks:
        lines.append("━━ 関ビズ ━━")
        for t in sekibiz_tasks:
            dl = f" [{t['deadline']}]" if t.get("deadline") else ""
            lines.append(f"- {t['task']}{dl}")
        lines.append("")

    # --- 申し送り ---
    handover = collect_journal_handover()
    if handover:
        lines.append("━━ 申し送り ━━")
        for item in handover:
            lines.append(f"- {item}")
        lines.append("")

    # --- サマリー ---
    total = len(siftai_tasks) + len(aisora_tasks) + len(bcbg_tasks) + len(sekibiz_tasks)
    lines.append(f"未完了タスク合計: {total}件")
    lines.append("")
    lines.append("COO SAEPINより: 冴香、今日もよろしくね！")

    return "\n".join(lines)


# === 送信 ===

def send_to_discord(message):
    """Discord Webhookに送信"""
    # 2000文字制限で分割
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
        data = json.dumps({"content": chunk}).encode("utf-8")
        req = urllib.request.Request(
            DISCORD_WEBHOOK,
            data=data,
            headers={"Content-Type": "application/json"},
        )
        try:
            with urllib.request.urlopen(req) as res:
                if res.status == 204:
                    print("Discord送信成功")
        except urllib.error.HTTPError as e:
            print(f"Discord送信失敗: {e.code}")


def send_to_line(message):
    """KOKONOE管理者LINEに送信"""
    if not LINE_TOKEN:
        print("LINE トークン未設定")
        return

    data = json.dumps({"messages": [{"type": "text", "text": message}]}).encode("utf-8")
    req = urllib.request.Request(
        "https://api.line.me/v2/bot/message/broadcast",
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {LINE_TOKEN}",
        },
    )

    try:
        with urllib.request.urlopen(req) as res:
            print(f"LINE送信成功 (status: {res.status})")
    except urllib.error.HTTPError as e:
        print(f"LINE送信失敗: {e.code}")
        print(e.read().decode())


def main():
    print(f"タスク収集開始... ({TODAY_STR})")

    message = format_message()

    print("\n--- 送信内容 ---")
    print(message)
    print("--- ここまで ---\n")

    send_to_discord(message)
    send_to_line(message)

    print("配信完了")


if __name__ == "__main__":
    main()
