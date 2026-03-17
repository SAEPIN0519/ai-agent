# ============================================
# PC起動時に自動git pullを設定するスクリプト
# 管理者権限のPowerShellで実行してください
# ============================================

$scriptPath = "$env:USERPROFILE\Desktop\ai-agent\auto_git_pull.ps1"
$taskName = "GitAutoPull_aiagent"

Write-Host "=== 自動git pull セットアップ ===" -ForegroundColor Cyan

# スクリプトが存在するか確認
if (-not (Test-Path $scriptPath)) {
    Write-Host "エラー: auto_git_pull.ps1 が見つかりません。" -ForegroundColor Red
    Write-Host "先にGitHubからプロジェクトをクローンしてください。"
    pause
    exit
}

# タスクスケジューラに登録（ログイン時に自動実行）
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""

$trigger = New-ScheduledTaskTrigger -AtLogOn

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 2)

Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -RunLevel Highest `
    -Force | Out-Null

Write-Host "設定完了！" -ForegroundColor Green
Write-Host "次回Windowsにログインしたとき、自動でGitHubから最新版を取得します。"
Write-Host ""
Write-Host "今すぐテストしますか？ (Y/N)"
$answer = Read-Host
if ($answer -eq "Y" -or $answer -eq "y") {
    & $scriptPath
}
