"""
Slack チャンネル読み取りモジュール

プロプレミアムTEAMのSlackチャンネルからメッセージを取得し、
AIエージェントが活用できる形式で出力する。

使い方:
  python slack_reader.py              # 直近1週間のサマリーを表示
  python slack_reader.py --days 30    # 直近30日分
  python slack_reader.py --json       # JSON形式で出力（エージェント連携用）
  python slack_reader.py --member "野村冴香"  # 特定メンバーのみ
"""

import sys
import io
import json
import argparse
from pathlib import Path
from datetime import datetime, timedelta
from collections import defaultdict, Counter

# Windows環境のUTF-8対応
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 設定
CONFIG_DIR = Path(__file__).parent / "config"
TOKEN_FILE = CONFIG_DIR / "slack_bot_token.txt"
CHANNEL_ID = "C0AC8404FPE"  # community_ss_pro-premium_plan

# bot名（集計から除外）
BOT_NAMES = {"SAEPIN", "Slackbot", "PRO PREMIUM Daily Progress", "成果報告通知"}


def get_client():
    """Slack APIクライアントを取得"""
    from slack_sdk import WebClient
    token = TOKEN_FILE.read_text(encoding="utf-8").strip()
    return WebClient(token=token)


def resolve_users(client, user_ids):
    """ユーザーIDから表示名を解決"""
    user_map = {}
    for uid in user_ids:
        try:
            info = client.users_info(user=uid)
            user_map[uid] = info['user'].get('real_name', info['user'].get('name', uid))
        except:
            user_map[uid] = uid
    return user_map


def fetch_messages(client, days=7, channel=CHANNEL_ID):
    """チャンネルのメッセージ＋スレッド返信を取得"""
    oldest = (datetime.now() - timedelta(days=days)).timestamp()

    # メインメッセージ取得
    all_msgs = []
    cursor = None
    while True:
        kwargs = {"channel": channel, "oldest": str(oldest), "limit": 200}
        if cursor:
            kwargs["cursor"] = cursor
        result = client.conversations_history(**kwargs)
        all_msgs.extend(result["messages"])
        cursor = result.get("response_metadata", {}).get("next_cursor")
        if not cursor:
            break

    # ユーザー名解決
    user_ids = set(m.get("user") for m in all_msgs if m.get("user"))

    # スレッド返信取得
    thread_replies = {}
    for m in all_msgs:
        if m.get("reply_count", 0) > 0:
            try:
                replies = client.conversations_replies(channel=channel, ts=m["ts"], limit=100)
                thread_msgs = replies.get("messages", [])[1:]  # 親メッセージを除く
                thread_replies[m["ts"]] = thread_msgs
                # 返信者のIDも収集
                for r in thread_msgs:
                    if r.get("user"):
                        user_ids.add(r["user"])
            except:
                pass

    user_map = resolve_users(client, user_ids)
    return all_msgs, thread_replies, user_map


def build_member_activities(msgs, thread_replies, user_map):
    """メンバー別の活動データを構築"""
    activities = defaultdict(list)

    for m in msgs:
        user = user_map.get(m.get("user", ""), m.get("username", "unknown"))
        if user in BOT_NAMES:
            continue

        ts = datetime.fromtimestamp(float(m["ts"]))
        text = m.get("text", "")

        entry = {
            "date": ts.strftime("%Y-%m-%d %H:%M"),
            "text": text,
            "reply_count": m.get("reply_count", 0),
        }

        # スレッド返信を追加
        if m["ts"] in thread_replies:
            replies = []
            for r in thread_replies[m["ts"]]:
                r_user = user_map.get(r.get("user", ""), "unknown")
                replies.append({
                    "user": r_user,
                    "text": r.get("text", ""),
                    "date": datetime.fromtimestamp(float(r["ts"])).strftime("%Y-%m-%d %H:%M"),
                })
            entry["replies"] = replies

        activities[user].append(entry)

    return dict(activities)


def summarize(activities):
    """メンバー別の活動サマリーを生成"""
    summary = {}
    for name, posts in activities.items():
        # 投稿のタイトル（【】内のテキスト）を抽出
        topics = []
        for p in posts:
            text = p["text"]
            import re
            match = re.search(r"【(.+?)】", text)
            if match:
                topics.append(match.group(1))

        summary[name] = {
            "post_count": len(posts),
            "topics": topics,
            "total_replies": sum(p.get("reply_count", 0) for p in posts),
        }
    return summary


def print_summary(activities):
    """人間が読みやすい形式でサマリーを表示"""
    summary = summarize(activities)

    print("=" * 60)
    print("プロプレミアムTEAM Slack活動サマリー")
    print(f"対象チャンネル: community_ss_pro-premium_plan")
    print("=" * 60)
    print()

    # 投稿数順にソート
    sorted_members = sorted(summary.items(), key=lambda x: x[1]["post_count"], reverse=True)

    for name, data in sorted_members:
        print(f"## {name}（投稿 {data['post_count']}件 / 返信 {data['total_replies']}件）")
        if data["topics"]:
            for t in data["topics"]:
                print(f"  - {t}")
        print()


def output_json(activities):
    """JSON形式で出力（エージェント連携用）"""
    output = {
        "generated_at": datetime.now().isoformat(),
        "channel": CHANNEL_ID,
        "member_activities": activities,
        "summary": summarize(activities),
    }
    print(json.dumps(output, ensure_ascii=False, indent=2))


def main():
    parser = argparse.ArgumentParser(description="Slack チャンネル読み取り")
    parser.add_argument("--days", type=int, default=7, help="取得する日数（デフォルト: 7）")
    parser.add_argument("--json", action="store_true", help="JSON形式で出力")
    parser.add_argument("--member", type=str, help="特定メンバーのみ表示")
    parser.add_argument("--channel", type=str, default=CHANNEL_ID, help="チャンネルID")
    args = parser.parse_args()

    client = get_client()
    msgs, thread_replies, user_map = fetch_messages(client, days=args.days, channel=args.channel)
    activities = build_member_activities(msgs, thread_replies, user_map)

    # メンバーフィルタ
    if args.member:
        filtered = {k: v for k, v in activities.items() if args.member in k}
        if not filtered:
            print(f"「{args.member}」に一致するメンバーが見つかりません")
            return
        activities = filtered

    if args.json:
        output_json(activities)
    else:
        print_summary(activities)


if __name__ == "__main__":
    main()
