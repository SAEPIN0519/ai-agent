"""
月次活動報告 自動生成スクリプト

毎月1日に前月分のSlackチャットを集計し、
PRO/PREMIUMプランの活動報告（MD + HTML）を自動生成する。

使い方:
  python monthly_report_generator.py              # 前月分を自動生成
  python monthly_report_generator.py --year 2026 --month 3  # 指定月を生成
  python monthly_report_generator.py --dry-run    # 生成内容をプレビュー（ファイル保存なし）

Windowsタスクスケジューラで毎月1日 09:00 に実行する想定。
"""

import sys
import io
import json
import argparse
from pathlib import Path
from datetime import datetime, timedelta
from collections import defaultdict
import calendar

# Windows環境のUTF-8対応
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# パス設定
BASE_DIR = Path(__file__).parent.parent
REPORT_DIR = BASE_DIR / "03_clients" / "SIFTAI" / "プロプレミアムTEAM" / "月活動報告"
SYSTEM_DIR = Path(__file__).parent

# slack_reader.py を再利用
sys.path.insert(0, str(SYSTEM_DIR))
from slack_reader import get_client, fetch_messages, build_member_activities, summarize


def get_target_month(year=None, month=None):
    """対象月を取得。指定なしなら前月"""
    if year and month:
        return year, month
    today = datetime.now()
    first_of_this_month = today.replace(day=1)
    last_month = first_of_this_month - timedelta(days=1)
    return last_month.year, last_month.month


def get_month_days(year, month):
    """指定月の日数を取得"""
    return calendar.monthrange(year, month)[1]


def generate_md_report(year, month, activities, summary):
    """MD形式の活動報告を生成"""
    month_days = get_month_days(year, month)
    start_date = f"{year}年{month}月1日"
    end_date = f"{year}年{month}月{month_days}日"

    # 総メッセージ数
    total_posts = sum(s["post_count"] for s in summary.values())
    total_replies = sum(s["total_replies"] for s in summary.values())

    # メンバー別セクション生成
    member_sections = []
    sorted_members = sorted(summary.items(), key=lambda x: x[1]["post_count"], reverse=True)
    for name, data in sorted_members:
        topics_str = ""
        if data["topics"]:
            topics_str = "\n".join(f"  - {t}" for t in data["topics"][:10])

        member_sections.append(
            f"### {name}（投稿 {data['post_count']}件 / スレッド返信 {data['total_replies']}件）\n\n"
            f"**主なトピック:**\n{topics_str if topics_str else '  - （トピック未分類）'}"
        )

    members_text = "\n\n---\n\n".join(member_sections)

    # MD本体
    md = f"""# PRO/PREMIUMプラン {month}月活動報告

> 対象期間：{start_date}〜{end_date}
> 生成日：{datetime.now().strftime('%Y年%m月%d日')}

---

## Key Numbers

| 総投稿数 | スレッド返信数 | アクティブメンバー |
|:---:|:---:|:---:|
| **{total_posts}件** | **{total_replies}件** | **{len(summary)}名** |

---

## メンバー別活動

{members_text}

---

## 月間トピック一覧

| メンバー | 投稿数 | 返信数 | 主なトピック |
|---|---:|---:|---|
"""
    for name, data in sorted_members:
        top_topics = "、".join(data["topics"][:3]) if data["topics"] else "—"
        md += f"| {name} | {data['post_count']} | {data['total_replies']} | {top_topics} |\n"

    md += f"""
---

## 備考

- このレポートはSlackチャンネル `community_ss_pro-premium_plan` のメッセージから自動生成されました
- BOT投稿（SAEPIN, Slackbot, PRO PREMIUM Daily Progress, 成果報告通知）は集計対象外です
- 詳細な施策・成果は手動での加筆をお勧めします
"""

    return md


def generate_html_report(year, month, activities, summary):
    """HTML形式の活動報告を生成（ダークモードLP風）"""
    month_days = get_month_days(year, month)
    start_date = f"{year}年{month}月1日"
    end_date = f"{year}年{month}月{month_days}日"

    total_posts = sum(s["post_count"] for s in summary.values())
    total_replies = sum(s["total_replies"] for s in summary.values())

    sorted_members = sorted(summary.items(), key=lambda x: x[1]["post_count"], reverse=True)

    # メンバーカード生成
    member_cards = ""
    colors = ["#ff6b35", "#4ecdc4", "#45b7d1", "#96ceb4", "#ffeaa7", "#dfe6e9", "#fd79a8", "#6c5ce7", "#00b894", "#e17055"]
    for i, (name, data) in enumerate(sorted_members):
        color = colors[i % len(colors)]
        topics_html = ""
        if data["topics"]:
            for t in data["topics"][:5]:
                topics_html += f'<span class="topic-tag">{t}</span>'
        else:
            topics_html = '<span class="topic-tag">トピック未分類</span>'

        member_cards += f"""
        <div class="member-card" style="border-left: 4px solid {color};">
            <h3>{name}</h3>
            <div class="stats">
                <span class="stat">投稿 <strong>{data['post_count']}</strong>件</span>
                <span class="stat">返信 <strong>{data['total_replies']}</strong>件</span>
            </div>
            <div class="topics">{topics_html}</div>
        </div>
"""

    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PRO/PREMIUMプラン {month}月活動報告</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;700&display=swap" rel="stylesheet">
<style>
:root {{
    --bg-primary: #0a0a1a;
    --bg-card: rgba(255,255,255,0.05);
    --text-primary: #e8e8f0;
    --text-secondary: #a0a0b8;
    --accent: #667eea;
    --accent-2: #764ba2;
}}
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{
    font-family: 'Noto Sans JP', sans-serif;
    background: var(--bg-primary);
    color: var(--text-primary);
    line-height: 1.8;
    min-height: 100vh;
    background: linear-gradient(135deg, #0a0a1a 0%, #1a1a3e 50%, #0a0a1a 100%);
}}
.container {{ max-width: 900px; margin: 0 auto; padding: 40px 20px; }}
h1 {{
    font-size: 2rem;
    font-weight: 700;
    background: linear-gradient(135deg, var(--accent), var(--accent-2));
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin-bottom: 8px;
}}
.subtitle {{ color: var(--text-secondary); margin-bottom: 40px; }}
.key-numbers {{
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 20px;
    margin-bottom: 40px;
}}
.number-card {{
    background: var(--bg-card);
    backdrop-filter: blur(10px);
    border: 1px solid rgba(255,255,255,0.1);
    border-radius: 16px;
    padding: 24px;
    text-align: center;
}}
.number-card .value {{
    font-size: 2.5rem;
    font-weight: 700;
    background: linear-gradient(135deg, var(--accent), var(--accent-2));
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}}
.number-card .label {{ color: var(--text-secondary); font-size: 0.9rem; }}
.section-title {{
    font-size: 1.4rem;
    margin: 40px 0 20px;
    padding-bottom: 8px;
    border-bottom: 2px solid rgba(102,126,234,0.3);
}}
.member-card {{
    background: var(--bg-card);
    backdrop-filter: blur(10px);
    border: 1px solid rgba(255,255,255,0.1);
    border-radius: 12px;
    padding: 20px;
    margin-bottom: 16px;
}}
.member-card h3 {{ font-size: 1.1rem; margin-bottom: 8px; }}
.member-card .stats {{ margin-bottom: 10px; }}
.member-card .stat {{
    display: inline-block;
    margin-right: 16px;
    color: var(--text-secondary);
    font-size: 0.9rem;
}}
.member-card .stat strong {{ color: var(--text-primary); }}
.topic-tag {{
    display: inline-block;
    background: rgba(102,126,234,0.15);
    border: 1px solid rgba(102,126,234,0.3);
    border-radius: 20px;
    padding: 4px 12px;
    font-size: 0.8rem;
    margin: 2px 4px;
    color: var(--text-secondary);
}}
.footer {{
    margin-top: 60px;
    padding-top: 20px;
    border-top: 1px solid rgba(255,255,255,0.1);
    color: var(--text-secondary);
    font-size: 0.85rem;
}}
@media (max-width: 600px) {{
    .key-numbers {{ grid-template-columns: 1fr; }}
    h1 {{ font-size: 1.5rem; }}
}}
</style>
</head>
<body>
<div class="container">
    <h1>PRO/PREMIUMプラン {month}月活動報告</h1>
    <p class="subtitle">{start_date}〜{end_date} ｜ 自動生成：{datetime.now().strftime('%Y/%m/%d')}</p>

    <div class="key-numbers">
        <div class="number-card">
            <div class="value">{total_posts}</div>
            <div class="label">総投稿数</div>
        </div>
        <div class="number-card">
            <div class="value">{total_replies}</div>
            <div class="label">スレッド返信数</div>
        </div>
        <div class="number-card">
            <div class="value">{len(summary)}</div>
            <div class="label">アクティブメンバー</div>
        </div>
    </div>

    <h2 class="section-title">メンバー別活動</h2>
    {member_cards}

    <div class="footer">
        <p>このレポートはSlackチャンネル community_ss_pro-premium_plan のメッセージから自動生成されました。</p>
        <p>BOT投稿は集計対象外です。詳細な施策・成果は手動での加筆をお勧めします。</p>
    </div>
</div>
</body>
</html>"""

    return html


def main():
    parser = argparse.ArgumentParser(description="月次活動報告 自動生成")
    parser.add_argument("--year", type=int, help="対象年（デフォルト: 前月の年）")
    parser.add_argument("--month", type=int, help="対象月（デフォルト: 前月）")
    parser.add_argument("--dry-run", action="store_true", help="プレビューのみ（ファイル保存しない）")
    args = parser.parse_args()

    year, month = get_target_month(args.year, args.month)
    month_days = get_month_days(year, month)

    print(f"対象月: {year}年{month}月（{month_days}日間）")
    print("Slackからデータ取得中...")

    # Slack APIからデータ取得
    client = get_client()
    msgs, thread_replies, user_map = fetch_messages(client, days=month_days, channel="C0AC8404FPE")

    # 対象月のメッセージのみフィルタリング
    filtered_msgs = []
    for m in msgs:
        ts = datetime.fromtimestamp(float(m["ts"]))
        if ts.year == year and ts.month == month:
            filtered_msgs.append(m)

    # 対象月のスレッド返信のみフィルタリング
    filtered_thread_replies = {}
    for parent_ts, replies in thread_replies.items():
        parent_dt = datetime.fromtimestamp(float(parent_ts))
        if parent_dt.year == year and parent_dt.month == month:
            filtered_thread_replies[parent_ts] = replies

    print(f"対象メッセージ: {len(filtered_msgs)}件")

    # 活動データ構築
    activities = build_member_activities(filtered_msgs, filtered_thread_replies, user_map)
    summary_data = summarize(activities)

    # レポート生成
    md_report = generate_md_report(year, month, activities, summary_data)
    html_report = generate_html_report(year, month, activities, summary_data)

    if args.dry_run:
        print("\n=== MD レポート（プレビュー） ===\n")
        print(md_report)
        print("\n=== HTML レポート生成OK ===")
        return

    # 保存先ディレクトリ作成
    output_dir = REPORT_DIR / f"{year}-{month:02d}"
    output_dir.mkdir(parents=True, exist_ok=True)

    # ファイル保存
    md_path = output_dir / f"活動報告_{year}-{month:02d}.md"
    html_path = output_dir / f"活動報告_{year}-{month:02d}.html"

    md_path.write_text(md_report, encoding="utf-8")
    html_path.write_text(html_report, encoding="utf-8")

    print(f"\n保存完了:")
    print(f"  MD:   {md_path}")
    print(f"  HTML: {html_path}")


if __name__ == "__main__":
    main()
