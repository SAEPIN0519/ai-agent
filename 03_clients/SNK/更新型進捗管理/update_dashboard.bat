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
echo ブラウザで開きます...
start "" "dashboard.html"
timeout /t 3 >nul
