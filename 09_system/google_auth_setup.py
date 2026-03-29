"""
Google OAuth認証セットアップ
ブラウザでGoogleログインするだけでOK。

使い方:
  python3 09_system/google_auth_setup.py
"""

import pickle
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = PROJECT_ROOT / "09_system" / "config"
CLIENT_FILE = CONFIG_DIR / "google_oauth_client.json"
TOKEN_FILE = CONFIG_DIR / "google_oauth_token.pickle"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
]


def main():
    if not CLIENT_FILE.exists():
        print(f"エラー: {CLIENT_FILE} が見つかりません")
        return

    # すでにトークンがあるか確認
    if TOKEN_FILE.exists():
        from google.auth.transport.requests import Request
        with open(TOKEN_FILE, 'rb') as f:
            creds = pickle.load(f)
        if not creds.expired:
            print("✓ 認証済み（トークン有効）")
            return
        else:
            print("トークン期限切れ → 自動更新中...")
            try:
                creds.refresh(Request())
                with open(TOKEN_FILE, 'wb') as f:
                    pickle.dump(creds, f)
                print("✓ トークン更新完了")
                return
            except Exception:
                print("自動更新失敗 → 再認証します...")

    # ブラウザでOAuth認証
    from google_auth_oauthlib.flow import InstalledAppFlow
    print("ブラウザが開きます。冴香さんのGoogleアカウントでログインしてください...")
    flow = InstalledAppFlow.from_client_secrets_file(str(CLIENT_FILE), SCOPES)
    creds = flow.run_local_server(port=8090, prompt='consent')

    with open(TOKEN_FILE, 'wb') as f:
        pickle.dump(creds, f)
    print(f"✓ 認証完了！")


if __name__ == "__main__":
    main()
