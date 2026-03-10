"""
毎朝8:30にGoogleカレンダーの予定をDiscordへ配信するスクリプト
- Google Calendar APIで今日の予定を取得
- Discord Webhookで送信
- Windows タスクスケジューラで毎日8:30に実行

初回セットアップ:
  1. pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib
  2. Google Cloud Console で OAuth 2.0 クライアントIDを作成
  3. credentials.json を 09_system/config/ に配置
  4. python daily_calendar_discord.py を手動で1回実行（ブラウザ認証）
"""

import os
import requests
from datetime import datetime, timezone, timedelta
from pathlib import Path

# === 設定 ===
WEBHOOK_URL = "https://discord.com/api/webhooks/1480602170368856275/DCDEaDwYhIIH0xGioPjanNbTocElv0FxtbnGo7HTqSDfpR_bn_IPO6Rm2PNhC7p0gUB4"
SCRIPT_DIR = Path(__file__).resolve().parent
CREDENTIALS_PATH = SCRIPT_DIR / "config" / "credentials.json"
TOKEN_PATH = SCRIPT_DIR / "config" / "calendar_token.json"

# 日本時間
JST = timezone(timedelta(hours=9))
TODAY = datetime.now(JST)


def get_calendar_events():
    """Google Calendar APIで今日の予定を取得"""
    try:
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build
    except ImportError:
        print("❌ Google APIライブラリが未インストールです")
        print("   pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib")
        return None

    SCOPES = ["https://www.googleapis.com/auth/calendar.readonly"]

    creds = None

    # 保存済みトークンがあれば読み込む
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)

    # トークンが無効なら再認証
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_PATH.exists():
                print(f"❌ 認証ファイルが見つかりません: {CREDENTIALS_PATH}")
                print("   Google Cloud Console から OAuth 2.0 クライアントIDをダウンロードしてください")
                return None
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), SCOPES)
            creds = flow.run_local_server(port=0)

        # トークンを保存
        TOKEN_PATH.parent.mkdir(parents=True, exist_ok=True)
        TOKEN_PATH.write_text(creds.to_json())

    # カレンダーAPI呼び出し
    service = build("calendar", "v3", credentials=creds)

    # 今日の開始・終了（JST）
    start_of_day = TODAY.replace(hour=0, minute=0, second=0, microsecond=0).isoformat()
    end_of_day = TODAY.replace(hour=23, minute=59, second=59, microsecond=0).isoformat()

    result = service.events().list(
        calendarId="primary",
        timeMin=start_of_day,
        timeMax=end_of_day,
        singleEvents=True,
        orderBy="startTime",
    ).execute()

    return result.get("items", [])


def format_event_time(event):
    """イベントの時間を整形"""
    start = event.get("start", {})

    if "dateTime" in start:
        # 時間指定あり
        start_dt = datetime.fromisoformat(start["dateTime"])
        end_dt = datetime.fromisoformat(event["end"]["dateTime"])
        return f"{start_dt.strftime('%H:%M')}〜{end_dt.strftime('%H:%M')}"
    else:
        # 終日イベント
        return "終日"


def format_discord_message(events):
    """Discord送信用のメッセージを整形"""
    lines = []
    lines.append(f"# 📅 今日のスケジュール — {TODAY.strftime('%Y/%m/%d (%a)')}")
    lines.append("")

    if not events:
        lines.append("予定はありません。自由に使える一日です！")
    else:
        # 終日イベントと時間指定イベントを分離
        allday = [e for e in events if "date" in e.get("start", {})]
        timed = [e for e in events if "dateTime" in e.get("start", {})]

        if allday:
            lines.append("### 📌 終日")
            for e in allday:
                summary = e.get("summary", "（タイトルなし）")
                lines.append(f"- {summary}")
            lines.append("")

        if timed:
            lines.append("### ⏰ 時間指定")
            for e in timed:
                time_str = format_event_time(e)
                summary = e.get("summary", "（タイトルなし）")
                location = e.get("location", "")
                loc_str = f"  📍 {location}" if location else ""
                lines.append(f"- **{time_str}** {summary}{loc_str}")
            lines.append("")

        lines.append(f"📊 本日の予定: **{len(events)}件**")

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
            print("✅ Discord送信成功")
        else:
            print(f"❌ Discord送信失敗: {response.status_code} - {response.text}")


def main():
    print(f"📅 Googleカレンダー取得開始... ({TODAY.strftime('%Y-%m-%d')})")

    events = get_calendar_events()

    if events is None:
        print("⚠️ カレンダー取得に失敗しました。セットアップを確認してください。")
        return

    print(f"  予定: {len(events)}件")

    message = format_discord_message(events)

    print("\n--- 送信内容 ---")
    print(message)
    print("--- ここまで ---\n")

    send_to_discord(message)

    # LINE にも同時配信
    from line_notify import send_to_line
    send_to_line(message)


if __name__ == "__main__":
    main()
