@echo off
chcp 65001 >nul
echo タスクスケジューラに登録中...
schtasks /create /tn "SNK_Dashboard_Update" /tr "node \"C:\Users\matsuoka\Desktop\ai-agent\03_clients\SNK\更新型進捗管理\update_dashboard.js\"" /sc daily /st 08:00 /f
if %ERRORLEVEL% EQU 0 (
    echo 登録完了！毎朝8:00に自動更新されます。
) else (
    echo エラー: 管理者権限で実行してください。
)
pause
