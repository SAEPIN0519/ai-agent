# ============================================
# GitHub から最新版を自動取得するスクリプト
# ============================================

$projectPath = "$env:USERPROFILE\Desktop\ai-agent"

# プロジェクトフォルダが存在するか確認
if (-not (Test-Path $projectPath)) {
    Write-Host "エラー: $projectPath が見つかりません。" -ForegroundColor Red
    Write-Host "GitHubからクローンしてください。"
    pause
    exit
}

Set-Location $projectPath

Write-Host "GitHubから最新版をチェック中..." -ForegroundColor Cyan
$result = git pull 2>&1

if ($result -match "Already up to date") {
    Write-Host "最新版です。更新なし。" -ForegroundColor Green
} elseif ($result -match "error|fatal") {
    Write-Host "エラーが発生しました:" -ForegroundColor Red
    Write-Host $result
} else {
    Write-Host "更新しました！" -ForegroundColor Green
    Write-Host $result
}
