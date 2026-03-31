# SNK ダッシュボード自動更新タスク登録
$taskName = "SNK_Dashboard_Update"
$batPath  = "C:\Users\matsuoka\Desktop\ai-agent\03_clients\SNK\更新型進捗管理\update_dashboard.bat"
$workDir  = "C:\Users\matsuoka\Desktop\ai-agent\03_clients\SNK\更新型進捗管理"

# 既存タスク削除
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

# アクション: バッチファイル実行（引数 auto = 自動実行モード）
$action = New-ScheduledTaskAction -Execute $batPath -Argument "auto" -WorkingDirectory $workDir

# トリガー: 毎日 09:30
$trigger = New-ScheduledTaskTrigger -Daily -At "09:30"

# 設定: PC起動中なら実行、バッテリーでも実行
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

# 登録
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest -Force

Write-Host ""
Write-Host "登録完了！" -ForegroundColor Green
Write-Host "  タスク名: $taskName"
Write-Host "  実行時刻: 毎日 09:30"
Write-Host "  実行内容: update_dashboard.bat auto"
Write-Host ""
