@echo off
REM 毎朝8:30にGoogleカレンダー → Discord配信をタスクスケジューラに登録
REM 管理者権限で実行してください

schtasks /create /tn "SAEPIN_DailyCalendar" /tr "python \"%~dp0daily_calendar_discord.py\"" /sc daily /st 08:30 /f

if %errorlevel% equ 0 (
    echo ✅ タスクスケジューラに登録しました（毎朝8:30）
    echo    タスク名: SAEPIN_DailyCalendar
) else (
    echo ❌ 登録に失敗しました。管理者権限で実行してください。
)
pause
