@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: ログファイル設定
set LOGFILE=%~dp0logs\update_%date:~0,4%%date:~5,2%%date:~8,2%.log

:: logsフォルダがなければ作成
if not exist "%~dp0logs" mkdir "%~dp0logs"

echo ===== %date% %time% 更新開始 ===== >> "%LOGFILE%" 2>&1

:: 手動実行かタスクスケジューラ実行かを判定
set MANUAL=0
if "%1"=="" if "%SESSIONNAME%"=="Console" set MANUAL=1

echo ダッシュボード更新中...
node update_dashboard.js >> "%LOGFILE%" 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [エラー] node update_dashboard.js が失敗しました >> "%LOGFILE%"
    echo エラーが発生しました（ログ: %LOGFILE%）
    if "%MANUAL%"=="1" pause
    exit /b 1
)

echo.
echo 共有フォルダにコピー中...
copy /y "更新型ダッシュボード.html" "\\192.168.1.233\技術部\新製品進捗管理\※技術管理用　新規・更新進捗表※\更新型ダッシュボード.html" >> "%LOGFILE%" 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [警告] 共有フォルダへのコピーに失敗（ネットワーク未接続の可能性） >> "%LOGFILE%"
    echo 共有フォルダへのコピーに失敗しました（ネットワーク接続を確認してね）
) else (
    echo [OK] 共有フォルダに反映完了 >> "%LOGFILE%"
    echo 共有フォルダに反映しました
)

echo ===== %date% %time% 更新完了 ===== >> "%LOGFILE%" 2>&1

:: 手動実行時のみブラウザで開く
if "%MANUAL%"=="1" (
    echo.
    echo ブラウザで開きます...
    start "" "更新型ダッシュボード.html"
    timeout /t 3 >nul
)

:: 7日以上前のログを自動削除
forfiles /p "%~dp0logs" /m "update_*.log" /d -7 /c "cmd /c del @path" >nul 2>&1
