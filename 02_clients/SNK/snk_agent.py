"""
SNK専用 AIエージェント
ダイカスト事業部の損益管理・経営分析をサポートします。

使い方:
    python snk_agent.py

必要なパッケージ:
    pip install anthropic
"""

import os
import csv
import json
from pathlib import Path
import anthropic

# =====================
# 設定
# =====================
MODEL = "claude-opus-4-6"
CSV_PATH = Path(__file__).parent / "ダイカスト事業部損益_2026.csv"

client = anthropic.Anthropic()  # 環境変数 ANTHROPIC_API_KEY を使用

# =====================
# ツール定義
# =====================

def read_pl_data() -> str:
    """損益CSVを読み込んで文字列で返す"""
    if not CSV_PATH.exists():
        return "損益ファイルが見つかりません。"
    with open(CSV_PATH, encoding="utf-8") as f:
        return f.read()


def update_pl_cell(科目: str, 月: str, 金額: int) -> str:
    """損益CSVの特定セルを更新する"""
    if not CSV_PATH.exists():
        return "損益ファイルが見つかりません。"

    rows = []
    with open(CSV_PATH, encoding="utf-8", newline="") as f:
        reader = csv.reader(f)
        rows = list(reader)

    # ヘッダー行から月の列インデックスを取得
    header = rows[0]
    if 月 not in header:
        return f"列 '{月}' が見つかりません。4月〜3月 または 合計 を指定してください。"
    col_idx = header.index(月)

    # 科目行を検索して更新
    updated = False
    for row in rows:
        if row[0].strip() == 科目.strip():
            row[col_idx] = str(金額)
            updated = True
            break

    if not updated:
        return f"科目 '{科目}' が見つかりません。"

    with open(CSV_PATH, encoding="utf-8", newline="") as f:
        pass  # 確認用

    with open(CSV_PATH, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(rows)

    return f"✅ {科目} / {月} を {金額:,} 円に更新しました。"


def summarize_pl() -> str:
    """損益データのサマリーを計算して返す"""
    if not CSV_PATH.exists():
        return "損益ファイルが見つかりません。"

    with open(CSV_PATH, encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    summary = {}
    key_items = ["売上高 合計", "売上原価 合計", "売上総利益", "販管費 合計", "営業利益", "経常利益"]
    for row in rows:
        科目 = row.get("科目", "").strip()
        if 科目 in key_items:
            try:
                summary[科目] = int(row.get("合計", 0) or 0)
            except ValueError:
                summary[科目] = 0

    if not summary:
        return "まだデータが入力されていません。"

    lines = ["【ダイカスト事業部 損益サマリー（年間合計）】"]
    for 科目, 金額 in summary.items():
        lines.append(f"  {科目}: {金額:,} 円")
    return "\n".join(lines)


# =====================
# ツールスキーマ
# =====================
TOOLS = [
    {
        "name": "read_pl_data",
        "description": "ダイカスト事業部の損益CSVファイルを読み込んで返します。現在のデータを確認したいときに使います。",
        "input_schema": {"type": "object", "properties": {}, "required": []}
    },
    {
        "name": "update_pl_cell",
        "description": "損益CSVの特定の科目・月の金額を更新します。",
        "input_schema": {
            "type": "object",
            "properties": {
                "科目": {"type": "string", "description": "更新する科目名。例: 製品売上、材料費"},
                "月": {"type": "string", "description": "更新する月。例: 4月、5月 または 合計"},
                "金額": {"type": "integer", "description": "設定する金額（円）"}
            },
            "required": ["科目", "月", "金額"]
        }
    },
    {
        "name": "summarize_pl",
        "description": "損益データの主要科目の年間合計サマリーを計算して返します。",
        "input_schema": {"type": "object", "properties": {}, "required": []}
    }
]

TOOL_FUNCTIONS = {
    "read_pl_data": lambda args: read_pl_data(),
    "update_pl_cell": lambda args: update_pl_cell(args["科目"], args["月"], args["金額"]),
    "summarize_pl": lambda args: summarize_pl(),
}

# =====================
# エージェントループ
# =====================
SYSTEM_PROMPT = """あなたはSNK社のダイカスト事業部専用AIアシスタントです。

主な役割:
- 損益データの読み取り・更新サポート
- 売上・原価・利益の分析と解説
- 経営改善のアドバイス

利用可能なツール:
- read_pl_data: 損益データの確認
- update_pl_cell: 数値の更新
- summarize_pl: 年間サマリーの表示

回答は日本語で、わかりやすく簡潔にお願いします。
金額は円単位で、3桁区切りで表示してください。"""

def run_agent(user_message: str) -> str:
    """ユーザーメッセージを受け取りエージェントを実行する"""
    messages = [{"role": "user", "content": user_message}]

    while True:
        response = client.messages.create(
            model=MODEL,
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            tools=TOOLS,
            messages=messages
        )

        # ツール呼び出しがない場合は完了
        if response.stop_reason == "end_turn":
            return next(
                (block.text for block in response.content if block.type == "text"),
                "（回答なし）"
            )

        # ツールを実行
        messages.append({"role": "assistant", "content": response.content})
        tool_results = []
        for block in response.content:
            if block.type == "tool_use":
                func = TOOL_FUNCTIONS.get(block.name)
                if func:
                    result = func(block.input)
                else:
                    result = f"ツール '{block.name}' が見つかりません。"
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": result
                })

        messages.append({"role": "user", "content": tool_results})


# =====================
# メインループ（対話型）
# =====================
def main():
    print("=" * 50)
    print("  SNK ダイカスト事業部 AIエージェント")
    print("=" * 50)
    print("終了するには 'exit' または 'quit' と入力してください。\n")

    while True:
        try:
            user_input = input("あなた: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n終了します。")
            break

        if not user_input:
            continue
        if user_input.lower() in ("exit", "quit", "終了"):
            print("終了します。")
            break

        print("エージェント: ", end="", flush=True)
        response = run_agent(user_input)
        print(response)
        print()


if __name__ == "__main__":
    main()
