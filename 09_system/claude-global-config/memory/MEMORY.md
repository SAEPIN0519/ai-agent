# プロジェクトメモリ

## プロジェクト概要
- リポジトリ: c:\Users\npino\ai-agent
- ミッション: AIエージェント企業の組織構築
- 社長: SAEKA（ユーザー本人）
- COO: SAEPIN（Claude Code本体。社長SAEKAが命名）
- 社員: AIエージェント（01_company/agents/ で管理）。各プロジェクトに1名アサイン

## フォルダ構成（2026-03-18 整理済み）
```
ai-agent/
├── 01_company/              # 自社情報
│   ├── CEO/                 # 冴香のプロフィール
│   ├── COO/                 # SAEPINの情報
│   ├── agents/              # AIエージェント社員
│   ├── finance/             # 財務
│   └── marketing/           # マーケティング
├── 02_thinking/             # 冴香タイム（FLOW.md + 5思考パターン）
├── 03_clients/              # クライアント別
│   ├── SIFTAI/
│   ├── KOKONOE/
│   ├── SNK/
│   ├── AIsoraCode/          # 青嶋さんとの共同事業
│   └── 関ビズ/
├── 04_journal/YYYY-MM/      # 日次ジャーナル
├── 05_skills/               # スキル集（ai/business/technical）
├── 09_system/               # システム・設定・スクリプト
└── .claude/                 # Claude Code設定
```

## ユーザー指示（記憶事項）
- 使われていないフォルダが出てきたら適宜教えること
- ファイルを追加したフォルダに .gitkeep が残っていたら自動で削除すること（確認不要）

## 思考パターン（冴香タイム）
- 起動ワード：「冴香タイム」
- 5人格（名前つき）：くりてぃん・ストラッチ・でびるん・データりやん・クリエイティブ
- 定義ファイル: 02_thinking/FLOW.md

## ジャーナル自動更新ルール（毎回必須・自動実行）
- パス形式: 04_journal/YYYY-MM/YYYY-MM-DD.md
- 最新ジャーナル: 04_journal/2026-03/2026-03-18.md

## Discord毎朝タスク配信
- スクリプト: 09_system/毎朝タスク配信_discord.py

## GitHub リモート
- origin → SAEPIN0519/ai-agent（自社用）
- aisoracode → Aosy-Git/AIsora-Code（青嶋さんとの共通リポジトリ）

## リファレンス
- [reference_google_drive_oauth.md](reference_google_drive_oauth.md)

## フィードバック
- [feedback_no_title.md](feedback_no_title.md)
- [feedback_always_japanese.md](feedback_always_japanese.md)
- [feedback_github_push_rules.md](feedback_github_push_rules.md)
