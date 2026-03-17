---
description: クライアントフォルダ（02_clients/）の管理ルール
globs: 02_clients/**/*
---

# クライアント管理ルール

- 各クライアントは `02_clients/{クライアント名}/projects/` 配下で管理
- 新クライアント追加時は `09_system/daily_discord_tasks.py` にcollect関数を追加する
- AIsoraCode関連のファイルはリモート `aisoracode`（Aosy-Git/AIsora-Code）にもpush対象
- それ以外のファイルは `origin`（SAEPIN0519/ai-agent）のみにpush
