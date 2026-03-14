"""
プロプレミアムTEAM手順書 自動更新スクリプト

毎週月曜23:50に実行。
1. Slackチャンネルの直近1週間の会話を取得
2. MTG議事録の最新版を確認
3. Claude APIで手順書の変更点を分析
4. 変更がある場合、手順書を更新 + 改定履歴を追記

使い方:
  python 手順書自動更新.py              # 通常実行（Slack + MTG議事録チェック）
  python 手順書自動更新.py --dry-run    # 変更内容を表示するだけ（ファイル更新しない）
  python 手順書自動更新.py --days 14    # 直近14日分のSlackを確認
"""

import sys
import io
import os
import re
import json
import argparse
from pathlib import Path
from datetime import datetime, timedelta

# Windows環境のUTF-8対応
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# パス設定
SCRIPT_DIR = Path(__file__).parent
PROJECT_ROOT = SCRIPT_DIR.parent
MANUAL_PATH = PROJECT_ROOT / "02_clients" / "SIFTAI" / "プロプレミアムTEAM" / "プロプレミアムTEAM手順書.md"
MTG_DIR = PROJECT_ROOT / "02_clients" / "SIFTAI" / "プロプレミアムTEAM" / "MTG議事録"
CONFIG_DIR = SCRIPT_DIR / "config"
TOKEN_FILE = CONFIG_DIR / "slack_bot_token.txt"
ANTHROPIC_KEY_FILE = CONFIG_DIR / "anthropic_api_key.txt"
CHANNEL_ID = "C0AC8404FPE"  # community_ss_pro-premium_plan

# Botなど集計から除外する名前
BOT_NAMES = {"SAEPIN", "Slackbot", "PRO PREMIUM Daily Progress", "成果報告通知"}


def get_slack_client():
    """Slack APIクライアントを取得"""
    from slack_sdk import WebClient
    token = TOKEN_FILE.read_text(encoding="utf-8").strip()
    return WebClient(token=token)


def fetch_slack_conversations(days=7):
    """Slackチャンネルから直近の会話を取得し、テキストにまとめる"""
    print(f"📡 Slackから直近{days}日分の会話を取得中...")

    client = get_slack_client()
    oldest = (datetime.now() - timedelta(days=days)).timestamp()

    # メッセージ取得
    all_msgs = []
    cursor = None
    while True:
        kwargs = {"channel": CHANNEL_ID, "oldest": str(oldest), "limit": 200}
        if cursor:
            kwargs["cursor"] = cursor
        result = client.conversations_history(**kwargs)
        all_msgs.extend(result["messages"])
        cursor = result.get("response_metadata", {}).get("next_cursor")
        if not cursor:
            break

    # ユーザー名解決
    user_ids = set(m.get("user") for m in all_msgs if m.get("user"))
    user_map = {}
    for uid in user_ids:
        try:
            info = client.users_info(user=uid)
            user_map[uid] = info['user'].get('real_name', info['user'].get('name', uid))
        except:
            user_map[uid] = uid

    # スレッド返信も含めて会話をテキスト化
    conversations = []
    for m in sorted(all_msgs, key=lambda x: float(x["ts"])):
        user = user_map.get(m.get("user", ""), m.get("username", "unknown"))
        if user in BOT_NAMES:
            continue

        ts = datetime.fromtimestamp(float(m["ts"]))
        text = m.get("text", "").strip()
        if not text:
            continue

        conversations.append(f"[{ts.strftime('%m/%d %H:%M')}] {user}: {text}")

        # スレッド返信を取得
        if m.get("reply_count", 0) > 0:
            try:
                replies = client.conversations_replies(channel=CHANNEL_ID, ts=m["ts"], limit=100)
                for r in replies.get("messages", [])[1:]:
                    r_user = user_map.get(r.get("user", ""), "unknown")
                    if r_user in BOT_NAMES:
                        continue
                    r_ts = datetime.fromtimestamp(float(r["ts"]))
                    r_text = r.get("text", "").strip()
                    if r_text:
                        conversations.append(f"  └ [{r_ts.strftime('%m/%d %H:%M')}] {r_user}: {r_text}")
            except:
                pass

    print(f"  → {len(conversations)}件のメッセージを取得")
    return "\n".join(conversations)


def get_latest_mtg_minutes():
    """MTG議事録フォルダから最新の議事録を読み込む"""
    print("📋 MTG議事録の最新版を確認中...")

    # _template.md, README.md, ACTION_TRACKER.md を除いた議事録ファイルを検索
    exclude = {"_template.md", "README.md", "ACTION_TRACKER.md"}
    mtg_files = [f for f in MTG_DIR.glob("*.md") if f.name not in exclude]

    if not mtg_files:
        print("  → 議事録ファイルなし")
        return None, None

    # ファイル名の日付順でソート（2026-03-09_xxx.md 形式）
    mtg_files.sort(key=lambda f: f.name, reverse=True)
    latest = mtg_files[0]

    print(f"  → 最新議事録: {latest.name}")
    content = latest.read_text(encoding="utf-8")
    return latest.name, content


def read_current_manual():
    """現在の手順書を読み込む"""
    if not MANUAL_PATH.exists():
        return ""
    return MANUAL_PATH.read_text(encoding="utf-8")


def analyze_changes(manual_text, slack_text, mtg_name, mtg_text, dry_run=False):
    """Claude APIで変更点を分析し、更新された手順書を生成"""
    import anthropic

    print("🤖 Claude APIで変更点を分析中...")

    # APIキーをconfigファイルから読み込む
    api_key = ANTHROPIC_KEY_FILE.read_text(encoding="utf-8").strip()

    # 分析プロンプト
    prompt = f"""あなたはプロプレミアムTEAMの運営手順書を管理するAIです。

以下の情報源から、手順書に反映すべき変更点を分析してください。

## 現在の手順書
```markdown
{manual_text}
```

## 今週のSlack会話（直近1週間）
```
{slack_text[:15000] if slack_text else "（Slackデータなし）"}
```

## 最新MTG議事録（{mtg_name or "なし"}）
```markdown
{mtg_text[:10000] if mtg_text else "（新しい議事録なし）"}
```

## タスク
1. SlackやMTGの会話から「運営ルールの変更・追加・廃止」に該当するものを抽出してください
2. 単なる業務連絡や雑談は無視してください
3. 変更がある場合のみ、手順書を更新してください

## 出力形式
以下のJSON形式で回答してください:

```json
{{
  "has_changes": true/false,
  "changes": [
    {{
      "section": "変更対象のセクション名",
      "type": "追加/変更/削除",
      "description": "変更内容の要約",
      "source": "Slack会話/MTG議事録",
      "decided_date": "YYYY-MM-DD（決まった日付）"
    }}
  ],
  "updated_manual": "（変更がある場合のみ）更新後の手順書全文（markdown形式）"
}}
```

重要な注意:
- updated_manual には改定履歴セクションを含めないでください（スクリプト側で追加します）
- 変更がない場合は has_changes: false にして、updated_manual は空文字にしてください
- 手順書の既存構造・フォーマットを維持してください
- 「最終更新」の日付を今日の日付に更新してください
"""

    client = anthropic.Anthropic(api_key=api_key)
    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8000,
        messages=[{"role": "user", "content": prompt}],
    )

    # レスポンスからJSONを抽出
    response_text = response.content[0].text

    # JSONブロックを抽出
    json_match = re.search(r'```json\s*(.*?)\s*```', response_text, re.DOTALL)
    if json_match:
        result = json.loads(json_match.group(1))
    else:
        # JSONブロックがない場合、テキスト全体をJSONとしてパース
        try:
            result = json.loads(response_text)
        except json.JSONDecodeError:
            print("  → 分析結果のパースに失敗。変更なしとして扱います")
            return {"has_changes": False, "changes": [], "updated_manual": ""}

    return result


def build_revision_entry(changes):
    """改定履歴のエントリーを生成"""
    today = datetime.now().strftime("%Y-%m-%d")
    entries = []
    for c in changes:
        decided = c.get("decided_date", "不明")
        entries.append(f"  - [{c['type']}] {c['description']}（決定日: {decided} / ソース: {c['source']}）")

    return f"| {today} | " + " / ".join(
        f"{c['description']}" for c in changes
    ) + " | " + ", ".join(
        c.get("source", "") for c in changes
    ) + " |"


def update_manual(result, dry_run=False):
    """手順書を更新する"""
    if not result["has_changes"]:
        print("\n✅ 変更なし — 手順書は最新の状態です")
        return False

    changes = result["changes"]
    print(f"\n📝 {len(changes)}件の変更を検出:")
    for c in changes:
        decided = c.get("decided_date", "不明")
        print(f"  [{c['type']}] セクション「{c['section']}」: {c['description']}（決定日: {decided}）")

    if dry_run:
        print("\n🔍 --dry-run モード: ファイル更新はスキップします")
        return False

    # 手順書本文を更新
    updated_manual = result["updated_manual"]
    if not updated_manual:
        print("  → 更新後のテキストが空のため、スキップします")
        return False

    # 改定履歴セクションを構築
    today = datetime.now().strftime("%Y-%m-%d")

    # 既存の改定履歴を読み込む
    current_manual = read_current_manual()
    existing_history = ""
    history_match = re.search(
        r'## 改定履歴\s*\n\| 日付.*?\n\|---.*?\n((?:\|.*\n)*)',
        current_manual,
        re.MULTILINE
    )
    if history_match:
        existing_history = history_match.group(1)

    # 新しい改定履歴エントリー
    new_entry = build_revision_entry(changes)

    # 改定履歴セクションを手順書の末尾に追加
    revision_section = f"""

---

## 改定履歴

| 日付 | 変更内容 | ソース |
|---|---|---|
| {today} | 初版作成 | Slack定型文集・MTG議事録・README・運営業務ファイル | """ if not existing_history else ""

    if existing_history:
        # 既存の履歴テーブルに新しいエントリーを追加
        revision_section = f"""

---

## 改定履歴

| 日付 | 変更内容 | ソース |
|---|---|---|
{new_entry}
{existing_history.rstrip()}"""
    else:
        # 初回: 初版 + 今回の変更
        revision_section = f"""

---

## 改定履歴

| 日付 | 変更内容 | ソース |
|---|---|---|
{new_entry}
| 2026-03-14 | 初版作成 | Slack定型文集・MTG議事録・README・運営業務ファイル |"""

    # 手順書本文 + 改定履歴を結合して保存
    # updated_manual から既に改定履歴セクションがあれば削除
    updated_manual = re.sub(r'\n---\n\n## 改定履歴.*', '', updated_manual, flags=re.DOTALL).rstrip()

    final_text = updated_manual + "\n" + revision_section + "\n"

    MANUAL_PATH.write_text(final_text, encoding="utf-8")
    print(f"\n✅ 手順書を更新しました: {MANUAL_PATH.name}")
    return True


def main():
    parser = argparse.ArgumentParser(description="プロプレミアムTEAM手順書 自動更新")
    parser.add_argument("--dry-run", action="store_true", help="変更内容を表示するだけ（ファイル更新しない）")
    parser.add_argument("--days", type=int, default=7, help="Slackの取得日数（デフォルト: 7）")
    args = parser.parse_args()

    print("=" * 60)
    print("プロプレミアムTEAM 手順書自動更新")
    print(f"実行日時: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 60)

    # 1. 現在の手順書を読み込む
    manual_text = read_current_manual()
    if not manual_text:
        print("❌ 手順書が見つかりません")
        return

    # 2. Slack会話を取得
    try:
        slack_text = fetch_slack_conversations(days=args.days)
    except Exception as e:
        print(f"⚠️ Slack取得エラー: {e}")
        slack_text = ""

    # 3. 最新MTG議事録を取得
    mtg_name, mtg_text = get_latest_mtg_minutes()

    # 4. Claude APIで分析
    try:
        result = analyze_changes(manual_text, slack_text, mtg_name, mtg_text, dry_run=args.dry_run)
    except Exception as e:
        print(f"❌ Claude API分析エラー: {e}")
        return

    # 5. 手順書を更新
    update_manual(result, dry_run=args.dry_run)

    print("\n完了")


if __name__ == "__main__":
    main()
