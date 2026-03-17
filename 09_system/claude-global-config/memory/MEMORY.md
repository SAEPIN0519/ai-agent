# プロジェクトメモリ

## プロジェクト概要
- リポジトリ: c:\Users\npino\ai-agent
- ミッション: AIエージェント企業の組織構築
- 社長: SAEKA（ユーザー本人）
- COO: SAEPIN（Claude Code本体。社長SAEKAが命名）
- 社員: AIエージェント（07_agents/ で管理）。各プロジェクトに1名アサイン

## フォルダ構成（現在）
```
ai-agent/
├── 01_CEO/
│   └── thinking_patterns/   # 5思考パターン + FLOW.md（冴香タイム）
├── 02_clients/
│   ├── SIFTAI/projects/
│   ├── KOKONOE/projects/
│   ├── SNK/projects/
│   ├── AIsoraCode/projects/  # 青嶋さんとの共同事業（旧10_firstcode）
│   └── 関ビズ/projects/
├── 03_journal/
├── 04_finance/
├── 05_marketing/
├── 06_resources/
├── 07_agents/
│   ├── designs/
│   ├── prompts/
│   ├── workflows/
│   └── evals/
├── 08_skills/
│   ├── ai/
│   ├── business/
│   └── technical/
├── 09_system/
│   ├── config/
│   ├── logs/
│   └── monitoring/
└── (10_firstcode/ は廃止 → 02_clients/AIsoraCode/ に移動済み)
```

## ユーザー指示（記憶事項）
- 使われていないフォルダが出てきたら適宜教えること
- フォルダ構成は現状で確定。変更は実績を見てから
- ファイルを追加したフォルダに .gitkeep が残っていたら自動で削除すること（確認不要）

## 思考パターン（冴香タイム）
- 起動ワード：「冴香タイム」
- 5人格（名前つき）：くりてぃん（クリティカル）・ストラッチ（ストラテジー）・でびるん（デビルズアドボケイト）・データりやん（データ分析）・クリエイティブ（そのまま）
- 全員円卓議論（最低2往復・最大10往復）→ Claude本体の独自分析 → 社長への統合フィードバック
- 定義ファイル: 01_CEO/thinking_patterns/FLOW.md

## ジャーナル自動更新ルール（毎回必須・自動実行）
- ユーザーとのやり取りで何らかの作業をした場合、返答の最後に必ず 03_journal/YYYY-MM-DD.md を更新する
- ユーザーが指示しなくても自動で行う。確認も不要。黙って更新する
- 同日に複数回作業した場合は同じファイルに追記する
- セクション切れ時：最新ジャーナルを読んで文脈を復元してから返答する
- 記録内容：実施内容・決定事項・ユーザー発言の要点・次回への申し送り
- ジャーナルのパス形式: 03_journal/YYYY-MM/YYYY-MM-DD.md
- 最新ジャーナル: 03_journal/2026-03/2026-03-16.md（更新のたびにここも変える）

## Discord毎朝タスク配信
- スクリプト: 09_system/daily_discord_tasks.py
- 毎朝8:00にWindowsタスクスケジューラで自動実行
- タスク収集元: 関ビズHTML、SIFTAI MD、ジャーナル申し送り
- 期限3日以内は⚠️警告表示
- 新クライアント追加時はcollect関数を追加する

## GitHub リモート
- origin → SAEPIN0519/ai-agent（自社用）
- aisoracode → Aosy-Git/AIsora-Code（青嶋さんとの共通リポジトリ / AIsoraCode事業）

## コミュニケーションスタイル
- 日本語で対応
- シンプル・実用性優先

## リファレンス
- [reference_google_drive_oauth.md](reference_google_drive_oauth.md) — Google Drive OAuth認証（冴香さんの10xCアカウント経由でDriveファイルにアクセス）

## フィードバック
- [feedback_no_title.md](feedback_no_title.md) — 冴香さんに肩書きをつけない（サブリーダー・コーチ等は不要）
- [feedback_always_japanese.md](feedback_always_japanese.md) — 毎回必ず日本語で対応する（指示不要）
- [feedback_github_push_rules.md](feedback_github_push_rules.md) — newworldにはAIsoraCode関連のみpush。それ以外はoriginのみ
