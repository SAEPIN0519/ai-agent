# Claude Code グローバル設定

Claude Codeの `~/.claude/` 配下にある設定・メモリのバックアップ。
新PCセットアップ時に使用する。

## ファイル構成

```
claude-global-config/
├── CLAUDE.md          # グローバル指示（全プロジェクト共通）
├── settings.json      # 権限・設定
├── memory/            # プロジェクトメモリ
│   ├── MEMORY.md
│   ├── feedback_*.md
│   └── reference_*.md
├── setup.sh           # 新PCセットアップスクリプト
└── README.md
```

## 新PCでの使い方

```bash
cd ai-agent/09_system/claude-global-config
bash setup.sh
```

## 注意事項

- settings.json 内のパスにユーザー名が含まれる → 新PCのユーザー名に置換が必要な場合あり
- .credentials.json 等の認証情報はGitHubに入れない
