@echo off
chcp 65001 >nul
echo ダッシュボード更新中...
cd /d "%~dp0"
node update_dashboard.js
if %ERRORLEVEL% NEQ 0 (
    echo エラーが発生しました
    pause
    exit /b 1
)
echo.
echo 共有フォルダにコピー中...
copy /y "更新型ダッシュボード.html" "\\192.168.1.233\技術部\新製品進捗管理\※技術管理用　新規・更新進捗表※\更新型ダッシュボード.html"
if %ERRORLEVEL% NEQ 0 (
    echo 共有フォルダへのコピーに失敗しました（ネットワーク接続を確認してね）
) else (
    echo 共有フォルダに反映しました
)
echo.
echo ブラウザで開きます...
start "" "更新型ダッシュボード.html"
timeout /t 3 >nul
