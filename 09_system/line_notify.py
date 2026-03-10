"""
LINE Notify 共通送信モジュール
- 全配信スクリプトから共通で使用
- トークンは config/line_notify_token.txt から読み込み
"""

import requests
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
TOKEN_PATH = SCRIPT_DIR / "config" / "line_notify_token.txt"
LINE_NOTIFY_API = "https://notify-api.line.me/api/notify"


def _load_token():
    """トークンファイルからLINE Notifyトークンを読み込む"""
    if not TOKEN_PATH.exists():
        print(f"❌ LINE Notifyトークンが見つかりません: {TOKEN_PATH}")
        return None
    return TOKEN_PATH.read_text(encoding="utf-8").strip()


def send_to_line(message):
    """LINE Notifyにメッセージを送信（1000文字制限対応で分割送信）"""
    token = _load_token()
    if not token:
        return

    headers = {"Authorization": f"Bearer {token}"}

    # LINE Notifyは1000文字制限があるため分割送信
    chunks = []
    current = ""
    for line in message.split("\n"):
        if len(current) + len(line) + 1 > 950:
            chunks.append(current)
            current = line
        else:
            current += "\n" + line if current else line
    if current:
        chunks.append(current)

    for chunk in chunks:
        # LINE NotifyではMarkdown記法（# ** など）を除去してプレーンテキストにする
        clean = _strip_markdown(chunk)
        payload = {"message": clean}
        response = requests.post(LINE_NOTIFY_API, headers=headers, data=payload)
        if response.status_code == 200:
            print("✅ LINE送信成功")
        else:
            print(f"❌ LINE送信失敗: {response.status_code} - {response.text}")


def _strip_markdown(text):
    """Discord用Markdownを除去してLINE向けプレーンテキストにする"""
    import re
    # 見出し（# ## ###）を除去
    text = re.sub(r'^#{1,3}\s*', '', text, flags=re.MULTILINE)
    # 太字（**text**）→ text
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    # 斜体（_text_）→ text
    text = re.sub(r'(?<!\w)_(.+?)_(?!\w)', r'\1', text)
    # 引用（> ）はそのまま残す（LINEでも読みやすい）
    return text
