"""
KOKONOE管理者LINE 自動送信スクリプト
写真・テキスト・画像+テキストの送信に対応

使い方:
  python line_kokonoe_send.py --text "送信したいメッセージ"
  python line_kokonoe_send.py --image path/to/photo.jpg
  python line_kokonoe_send.py --text "メッセージ" --image path/to/photo.jpg
"""

import argparse
import json
import urllib.request
import urllib.error
import base64
import os
import sys

# 設定ファイルのパス
CONFIG_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config")
TOKEN_FILE = os.path.join(CONFIG_DIR, "line_kokonoe_access_token.txt")

def load_token():
    """アクセストークンを読み込む"""
    with open(TOKEN_FILE, "r") as f:
        return f.read().strip()

def send_broadcast(messages):
    """ブロードキャスト送信（全友だちに送信）"""
    token = load_token()
    data = json.dumps({"messages": messages}).encode("utf-8")

    req = urllib.request.Request(
        "https://api.line.me/v2/bot/message/broadcast",
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token}",
        },
    )

    try:
        with urllib.request.urlopen(req) as res:
            print(f"送信成功 (status: {res.status})")
            return True
    except urllib.error.HTTPError as e:
        print(f"送信失敗 (status: {e.code})")
        print(e.read().decode())
        return False

def upload_image(image_path):
    """画像をLINEサーバーにアップロードしてURLを返す（imgBB経由）"""
    # LINE Messaging APIで画像を送るにはURLが必要
    # ローカルファイルの場合、一時的な画像ホスティングが必要
    # ここではbase64でデータURIを使えないため、
    # 外部ホスティングまたはリッチメニュー用のblob APIを使う

    # 簡易的にimgur匿名アップロードを使用（無料）
    # もしくはローカルサーバーを立てる方法もある
    print(f"注意: LINE APIでは画像URLが必要です。")
    print(f"ローカルファイル '{image_path}' を直接送信するには、")
    print(f"画像ホスティングサービスが必要です。")
    return None

def send_text(text):
    """テキストメッセージを送信"""
    messages = [{"type": "text", "text": text}]
    return send_broadcast(messages)

def send_image_url(image_url, preview_url=None):
    """画像URL指定で送信"""
    if preview_url is None:
        preview_url = image_url
    messages = [
        {
            "type": "image",
            "originalContentUrl": image_url,
            "previewImageUrl": preview_url,
        }
    ]
    return send_broadcast(messages)

def send_text_and_image(text, image_url, preview_url=None):
    """テキスト + 画像を送信"""
    if preview_url is None:
        preview_url = image_url
    messages = [
        {"type": "text", "text": text},
        {
            "type": "image",
            "originalContentUrl": image_url,
            "previewImageUrl": preview_url,
        },
    ]
    return send_broadcast(messages)

def main():
    parser = argparse.ArgumentParser(description="KOKONOE管理者LINE 自動送信")
    parser.add_argument("--text", "-t", help="送信するテキストメッセージ")
    parser.add_argument("--image", "-i", help="送信する画像のURL")
    parser.add_argument("--image-file", "-f", help="送信する画像ファイル（要ホスティング）")

    args = parser.parse_args()

    if not args.text and not args.image and not args.image_file:
        parser.print_help()
        sys.exit(1)

    if args.image_file:
        print(f"ローカルファイル指定: {args.image_file}")
        upload_image(args.image_file)
        sys.exit(1)

    if args.text and args.image:
        send_text_and_image(args.text, args.image)
    elif args.text:
        send_text(args.text)
    elif args.image:
        send_image_url(args.image)

if __name__ == "__main__":
    main()
