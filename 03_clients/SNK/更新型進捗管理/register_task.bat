@echo off
chcp 65001 >nul
echo.
echo === SNK 更新型ダッシュボード タスク登録 ===
echo.
echo 既存タスクを削除中...
schtasks /delete /tn "SNK_Dashboard_Update" /f >nul 2>&1

echo 新規タスクを登録中...
schtasks /create ^
  /tn "SNK_Dashboard_Update" ^
  /tr "\"%~dp0update_dashboard.bat\" auto" ^
  /sc daily ^
  /st 09:30 ^
  /rl highest ^
  /f

if %ERRORLEVEL% EQU 0 (
    echo.
    echo 登録完了！
    echo   タスク名: SNK_Dashboard_Update
    echo   実行時刻: 毎日 09:30
    echo   実行内容: update_dashboard.bat（Excel読込→HTML更新→共有フォルダコピー）
    echo   ログ出力: logs\update_YYYYMMDD.log
    echo.
) else (
    echo.
    echo エラー: 登録に失敗しました。右クリック→「管理者として実行」で再試行してね。
)
pause
