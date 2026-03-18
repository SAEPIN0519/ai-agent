---
description: クライアントフォルダ（03_clients/）の管理ルール
globs: 03_clients/**/*
---

# クライアント管理ルール

- 各クライアントは `03_clients/{クライアント名}/projects/` 配下で管理
- 新クライアント追加時は `09_system/毎朝タスク配信_discord.py` にcollect関数を追加する
- AIsoraCode関連のファイルはリモート `aisoracode`（Aosy-Git/AIsora-Code）にもpush対象
- それ以外のファイルは `origin`（SAEPIN0519/ai-agent）のみにpush
