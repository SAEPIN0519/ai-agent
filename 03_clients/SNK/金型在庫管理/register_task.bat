@echo off
chcp 65001 >nul
echo ========================================
echo  タスクスケジューラ登録
echo  金型在庫管理 マスターデータ自動更新
echo ========================================
echo.

:: 管理者権限チェック
net session >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo 管理者権限が必要です。右クリック→「管理者として実行」してね
    pause
    exit /b 1
)

:: タスク登録（毎朝7:50に実行 — 更新型ダッシュボードの少し前）
schtasks /create /tn "SNK_金型在庫管理_マスター更新" /tr "\"%~dp0update_master.bat\" auto" /sc daily /st 07:50 /rl highest /f

if %ERRORLEVEL% EQU 0 (
    echo.
    echo 登録完了！毎朝 7:50 に自動更新されるよ
) else (
    echo.
    echo 登録に失敗しました
)

echo.
pause
