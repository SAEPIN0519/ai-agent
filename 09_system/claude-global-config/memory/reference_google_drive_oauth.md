---
name: Google Drive OAuth認証
description: SAEPINが冴香さんのGoogleアカウントでGoogle Driveファイルにアクセスする方法（10xc.jp組織内ファイル含む）
type: reference
---

## Google Drive OAuth認証（冴香さんアカウント経由）

### 背景
- サービスアカウント（sheets-reader@...）では10xc.jp組織内限定のGoogle Driveファイルにアクセスできない
- 冴香さんのOAuthトークンを使えば、冴香さんがアクセス権を持つ全てのGoogle Driveファイルを読める

### 認証情報の場所
- OAuthクライアントJSON: `09_system/config/google_oauth_client.json`
- OAuthトークン: `09_system/config/google_oauth_token.pickle`
- プロジェクト: siftai-report-analyzer (867999955472)
- クライアント名: SAEPIN（デスクトップアプリ）
- 対象: 内部（10xc.jp組織）

### 使い方（Pythonコード）
```python
import pickle
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.auth.transport.requests import Request
import io

TOKEN_FILE = 'c:/Users/npino/ai-agent/09_system/config/google_oauth_token.pickle'
with open(TOKEN_FILE, 'rb') as token:
    creds = pickle.load(token)

if creds.expired:
    creds.refresh(Request())
    # 更新されたトークンを保存
    with open(TOKEN_FILE, 'wb') as token:
        pickle.dump(creds, token)

service = build('drive', 'v3', credentials=creds)

# ファイルIDでダウンロード
file_id = 'GoogleDriveのファイルID'
request = service.files().get_media(fileId=file_id)
fh = io.BytesIO()
downloader = MediaIoBaseDownload(fh, request)
done = False
while not done:
    status, done = downloader.next_chunk()
```

### トークン期限切れ時
- トークンが完全に失効した場合は再認証が必要:
```python
from google_auth_oauthlib.flow import InstalledAppFlow
flow = InstalledAppFlow.from_client_secrets_file(CLIENT_FILE, SCOPES)
creds = flow.run_local_server(port=8090, prompt='consent')
```
- ブラウザが開くので冴香さんの10xCアカウントでログインする

### セットアップ済みAPI
- Google Sheets API: 有効
- Google Drive API: 有効（2026-03-15追加）

### セキュリティ注意
- google_oauth_client.json, google_oauth_token.pickle は .gitignore に追加すること
