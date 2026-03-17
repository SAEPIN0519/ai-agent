# ============================================
# Claude Code 記録共有セットアップ（PC1用）
# 管理者権限のPowerShellで実行してください
# ============================================

$source = "$env:USERPROFILE\.claude"
$target = "$env:USERPROFILE\OneDrive\.claude"

Write-Host "=== Claude Code 同期セットアップ開始 ===" -ForegroundColor Cyan

# すでにシンボリックリンクになっていないか確認
$item = Get-Item $source -ErrorAction SilentlyContinue
if ($item.LinkType -eq "SymbolicLink") {
    Write-Host "すでにシンボリックリンクが設定されています。終了します。" -ForegroundColor Yellow
    exit
}

# .claude フォルダを OneDrive へ移動
if (Test-Path $source) {
    Write-Host "'.claude' を OneDrive へ移動中..." -ForegroundColor White
    Move-Item -Path $source -Destination $target -Force
    Write-Host "移動完了: $target" -ForegroundColor Green
} else {
    Write-Host "'.claude' フォルダが見つかりません。" -ForegroundColor Red
    exit
}

# シンボリックリンクを作成
Write-Host "シンボリックリンクを作成中..." -ForegroundColor White
New-Item -ItemType SymbolicLink -Path $source -Target $target | Out-Null
Write-Host "作成完了: $source → $target" -ForegroundColor Green

Write-Host ""
Write-Host "=== PC1 セットアップ完了！===" -ForegroundColor Cyan
Write-Host "OneDrive の同期が終わったら、PC2・PC3 で sync_claude_setup_PC2_PC3.ps1 を実行してください。"
