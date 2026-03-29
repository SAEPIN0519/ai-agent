"""
週次報告 Slack自動投稿スクリプト
毎週月曜8:30にSlackへ週次総括を投稿する。

使い方:
  python3 09_system/slack_weekly_post.py              # 即時投稿
  python3 09_system/slack_weekly_post.py --schedule    # 次の月曜8:30に予約投稿
  python3 09_system/slack_weekly_post.py --dry-run     # テスト（投稿しない）
"""

import sys
import argparse
from pathlib import Path
from datetime import datetime, timedelta

# 設定
PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = PROJECT_ROOT / "09_system" / "config"
TOKEN_FILE = CONFIG_DIR / "slack_user_token.txt"  # 冴香名義で投稿するためユーザートークンを使用
REPORT_DIR = PROJECT_ROOT / "03_clients" / "SIFTAI" / "プロプレミアムTEAM" / "週次報告"
CHANNEL_ID = "C0AC8404FPE"  # community_ss_pro-premium_plan


def get_latest_report():
    """最新の週次報告ファイルを取得"""
    reports = sorted(REPORT_DIR.glob("202*.md"), reverse=True)
    if not reports:
        print("エラー: 週次報告ファイルが見つかりません")
        sys.exit(1)

    latest = reports[0]
    print(f"投稿ファイル: {latest.name}")
    return latest.read_text(encoding="utf-8")


def get_client():
    """Slack APIクライアントを取得"""
    from slack_sdk import WebClient
    token = TOKEN_FILE.read_text(encoding="utf-8").strip()
    return WebClient(token=token)


def post_now(client, text):
    """即時投稿"""
    response = client.chat_postMessage(
        channel=CHANNEL_ID,
        text=text,
        mrkdwn=True,
    )
    print(f"✓ 投稿完了: ts={response['ts']}")


def schedule_post(client, text):
    """次の月曜9:30に予約投稿（9:00のワークフロー後に投稿）"""
    now = datetime.now()
    # 次の月曜日を計算
    days_until_monday = (7 - now.weekday()) % 7
    if days_until_monday == 0 and now.hour >= 10:
        days_until_monday = 7  # 今日が月曜で10時過ぎなら来週
    next_monday = now.replace(hour=9, minute=30, second=0, microsecond=0) + timedelta(days=days_until_monday)

    post_at = int(next_monday.timestamp())

    response = client.chat_scheduleMessage(
        channel=CHANNEL_ID,
        text=text,
        post_at=post_at,
    )
    print(f"✓ 予約投稿完了: {next_monday.strftime('%Y-%m-%d %H:%M')} (ts={response['scheduled_message_id']})")


def main():
    parser = argparse.ArgumentParser(description="週次報告 Slack投稿")
    parser.add_argument("--schedule", action="store_true", help="次の月曜8:30に予約投稿")
    parser.add_argument("--dry-run", action="store_true", help="テスト実行（投稿しない）")
    args = parser.parse_args()

    text = get_latest_report()

    if args.dry_run:
        print("=== dry-run モード ===")
        print(f"チャンネル: {CHANNEL_ID}")
        print(f"文字数: {len(text)}")
        print("--- 内容 ---")
        print(text[:500] + "..." if len(text) > 500 else text)
        return

    client = get_client()

    if args.schedule:
        schedule_post(client, text)
    else:
        post_now(client, text)


if __name__ == "__main__":
    main()
