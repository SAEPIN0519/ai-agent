#!/bin/bash
# Claude Code グローバル設定セットアップスクリプト
# 新PCでこのスクリプトを実行すると、~/.claude/ に設定が配置される

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CLAUDE_DIR="$HOME/.claude"
MEMORY_DIR="$CLAUDE_DIR/projects/c--Users-${USER}-ai-agent/memory"

echo "=== Claude Code 設定セットアップ ==="

# ~/.claude/ フォルダ作成
mkdir -p "$CLAUDE_DIR"
mkdir -p "$MEMORY_DIR"

# CLAUDE.md コピー
cp "$SCRIPT_DIR/CLAUDE.md" "$CLAUDE_DIR/CLAUDE.md"
echo "[OK] CLAUDE.md → $CLAUDE_DIR/CLAUDE.md"

# settings.json コピー
cp "$SCRIPT_DIR/settings.json" "$CLAUDE_DIR/settings.json"
echo "[OK] settings.json → $CLAUDE_DIR/settings.json"

# メモリファイルコピー
cp "$SCRIPT_DIR/memory/"*.md "$MEMORY_DIR/"
echo "[OK] memory/ → $MEMORY_DIR/"

echo ""
echo "=== セットアップ完了 ==="
echo "注意: settings.json 内のパスが新PCのユーザー名と異なる場合は手動修正が必要"
echo "  例: npino → 新ユーザー名 に置換"
