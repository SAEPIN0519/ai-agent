# ============================================
# Claude Code 記録共有セットアップ（PC2・PC3用）
# 管理者権限のPowerShellで実行してください
# ※ PC1のセットアップとOneDrive同期完了後に実行
# ============================================

$source = "$env:USERPROFILE\.claude"
$target = "$env:USERPROFILE\OneDrive\.claude"

Write-Host "=== Claude Code 同期セットアップ開始（PC2/PC3）===" -ForegroundColor Cyan

# OneDrive に .claude が存在するか確認
if (-not (Test-Path $target)) {
    Write-Host "エラー: OneDrive に '.claude' フォルダが見つかりません。" -ForegroundColor Red
    Write-Host "PC1のセットアップとOneDriveの同期が完了しているか確認してください。"
    exit
}

# 既存の .claude フォルダをバックアップ
if (Test-Path $source) {
    $backup = "$source`_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    Write-Host "既存の '.claude' をバックアップ中: $backup" -ForegroundColor Yellow
    Rename-Item -Path $source -NewName $backup
    Write-Host "バックアップ完了" -ForegroundColor Green
}

# シンボリックリンクを作成
Write-Host "シンボリックリンクを作成中..." -ForegroundColor White
New-Item -ItemType SymbolicLink -Path $source -Target $target | Out-Null
Write-Host "作成完了: $source → $target" -ForegroundColor Green

Write-Host ""
Write-Host "=== セットアップ完了！===" -ForegroundColor Cyan
Write-Host "このPCのClaude Codeが3台で記録を共有するようになりました。"
