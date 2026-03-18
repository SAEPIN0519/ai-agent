---
name: GitHub push ルール（AIsoraCode / SAEPIN分離）
description: newworld（AIsora-Code）にはAIsoraCode関連のみpush。それ以外は必ずorigin（SAEPIN0519）にpush。青嶋さんがAIsora-Codeに追加したものは03_clients/AIsoraCode/に取り込んでOK
type: feedback
---

## GitHubリポジトリのpushルール

### origin（SAEPIN0519/ai-agent）
- 冴香さんのデータは**すべてこちら**にpush
- クライアント情報、ジャーナル、社内システム等

### newworld（Aosy-Git/AIsora-Code）
- `03_clients/AIsoraCode/` の中身**だけ**をpushしてよい
- それ以外のデータは絶対にpushしない

### 青嶋さんからの取り込み
- 青嶋さんがAIsora-Code（GitHub上）に何かアップした場合
- その内容を `03_clients/AIsoraCode/` に取り込んでOK

### 背景
- 2026-03-15に、全データがAIsora-Codeに入っていたことが判明
- クライアント情報（SIFTAI、KOKONOE、関ビズ等）が青嶋さんに見える状態だった
- 削除対応済み。今後は再発防止のためこのルールを厳守する
