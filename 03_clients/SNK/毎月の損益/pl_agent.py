"""
SNK ダイカスト事業部 月次損益 AI エージェント

使い方:
    python pl_agent.py

できること:
  - エクセルデータの貼り付け → CSV & HTML ダッシュボード自動更新
  - 現在の損益データの確認・分析
  - 利益率・前月比などのサマリーレポート

必要なパッケージ:
    pip install anthropic
"""

import sys
import csv
import json
from pathlib import Path
import anthropic

# =====================
# 設定
# =====================
MODEL = "claude-opus-4-6"
BASE_DIR = Path(__file__).parent
CSV_PATH = BASE_DIR / "ダイカスト事業部損益_2026.csv"
UPDATE_SCRIPT = BASE_DIR.parent / "update_dashboard.py"

client = anthropic.Anthropic()

# =====================
# ツール関数
# =====================

def read_pl_data() -> str:
    """損益CSVを読み込んで文字列で返す"""
    if not CSV_PATH.exists():
        return "損益ファイルが見つかりません。"
    with open(CSV_PATH, encoding="utf-8") as f:
        return f.read()


def update_pl_from_tsv(tsv_data: str) -> str:
    """
    タブ区切りのエクセルデータをパースして CSV と HTML を更新する。
    update_dashboard.py の parse_tsv / calc_aggregates / update_csv / generate_html を再利用。
    """
    # update_dashboard.py のロジックを import して使う
    sys.path.insert(0, str(BASE_DIR.parent))
    try:
        import update_dashboard as ud
    except ImportError as e:
        return f"update_dashboard.py の読み込みに失敗しました: {e}"

    try:
        months, data = ud.parse_tsv(tsv_data)
    except ValueError as e:
        return f"データ解析エラー: {e}"

    data = ud.calc_aggregates(data, months)

    # CSV更新
    try:
        n = ud.update_csv(months, data)
        csv_msg = f"CSV更新: {n} 科目を更新"
    except Exception as e:
        csv_msg = f"CSV更新失敗: {e}"

    # HTML生成
    try:
        ud.generate_html(months, data)
        html_msg = "HTML生成: monthly_dashboard.html を更新"
    except Exception as e:
        html_msg = f"HTML生成失敗: {e}"

    latest = months[-1]
    sales = data.get("売上高 合計", {}).get(latest, 0)
    op    = data.get("営業利益", {}).get(latest, 0)
    gross = data.get("売上総利益", {}).get(latest, 0)

    def fmt(n: int) -> str:
        if n == 0:
            return "—"
        sign = "▲" if n < 0 else ""
        return sign + f"{abs(round(n/10000)):,}万円"

    op_rate = f"{op/sales*100:.1f}%" if sales else "—"

    return (
        f"✅ 更新完了\n"
        f"  {csv_msg}\n"
        f"  {html_msg}\n\n"
        f"【{latest} サマリー】\n"
        f"  売上高    : {fmt(sales)}\n"
        f"  売上総利益: {fmt(gross)}\n"
        f"  営業利益  : {fmt(op)}（利益率 {op_rate}）\n\n"
        f"ブラウザで monthly_dashboard.html を再読み込みすると反映されます。"
    )


def summarize_pl(target_month: str = "") -> str:
    """損益データのサマリーを返す（月指定可能）"""
    if not CSV_PATH.exists():
        return "損益ファイルが見つかりません。"

    with open(CSV_PATH, encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        rows = {row["科目"]: row for row in reader}

    # 表示する主要科目
    key_items = [
        "売上高 合計", "売上原価 合計", "売上総利益",
        "販管費 合計", "営業利益", "経常利益"
    ]

    # ヘッダーから月リストを取得
    with open(CSV_PATH, encoding="utf-8") as f:
        header = f.readline().strip().split(",")
    months = [h for h in header[1:] if h and h != "合計"]

    def fmt(n: int) -> str:
        if n == 0:
            return "—"
        sign = "▲" if n < 0 else ""
        return sign + f"{abs(round(n/10000)):,}万円"

    if target_month and target_month in months:
        lines = [f"【{target_month} 損益サマリー】"]
        for item in key_items:
            if item in rows:
                try:
                    val = int(rows[item].get(target_month, 0) or 0)
                    lines.append(f"  {item}: {fmt(val)}")
                except ValueError:
                    pass
    else:
        lines = ["【年間合計 損益サマリー】"]
        for item in key_items:
            if item in rows:
                try:
                    total = int(rows[item].get("合計", 0) or 0)
                    lines.append(f"  {item}: {fmt(total)}")
                except ValueError:
                    pass
        if months:
            lines.append(f"\n入力済み月: {', '.join(m for m in months if any(rows.get(k, {}).get(m, '0') not in ('0', '') for k in key_items[:1]))}")

    return "\n".join(lines)


# =====================
# ツールスキーマ
# =====================
TOOLS = [
    {
        "name": "read_pl_data",
        "description": "現在の損益CSVデータを読み込んで返します。データ確認・分析に使います。",
        "input_schema": {"type": "object", "properties": {}, "required": []}
    },
    {
        "name": "update_pl_from_tsv",
        "description": (
            "エクセルからコピーしたタブ区切りデータを受け取り、"
            "損益CSVとHTMLダッシュボードを更新します。\n"
            "データ形式（1行目はヘッダー）:\n"
            "  科目\\t4月\\t5月\\t...\\t3月\\t合計\n"
            "  製品売上\\t1000000\\t1200000\\t...\n"
            "  材料費\\t400000\\t..."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "tsv_data": {
                    "type": "string",
                    "description": "エクセルからコピーしたタブ区切りの損益データ"
                }
            },
            "required": ["tsv_data"]
        }
    },
    {
        "name": "summarize_pl",
        "description": "損益データの主要科目サマリーを返します。月を指定することもできます。",
        "input_schema": {
            "type": "object",
            "properties": {
                "target_month": {
                    "type": "string",
                    "description": "サマリーを表示する月（例: 4月、11月）。省略すると年間合計を表示。"
                }
            },
            "required": []
        }
    }
]

TOOL_FUNCTIONS = {
    "read_pl_data":      lambda args: read_pl_data(),
    "update_pl_from_tsv": lambda args: update_pl_from_tsv(args["tsv_data"]),
    "summarize_pl":      lambda args: summarize_pl(args.get("target_month", "")),
}

# =====================
# システムプロンプト
# =====================
SYSTEM_PROMPT = """あなたはSNK社ダイカスト事業部の月次損益管理AIアシスタントです。

## 主な役割
- エクセルデータの貼り付けを受け取り、CSVとHTMLダッシュボードを更新する
- 損益データの分析・解説・アドバイス

## データ更新の手順
ユーザーがエクセルデータを貼り付けてきたら：
1. `update_pl_from_tsv` ツールを呼び出してデータを更新する
2. 更新結果と最新月のサマリーを日本語で分かりやすく報告する
3. 気になる点（赤字・利益率低下など）があれば一言コメントする

## データフォーマット
エクセルのタブ区切りコピー。1行目はヘッダー：
  科目[TAB]4月[TAB]5月[TAB]...[TAB]3月[TAB]合計

## 回答スタイル
- 日本語・簡潔・わかりやすく
- 金額は万円単位で表示
- ネガティブな数値（赤字など）は明確に指摘する"""


# =====================
# エージェントループ
# =====================
def run_agent(user_message: str) -> str:
    messages = [{"role": "user", "content": user_message}]

    while True:
        response = client.messages.create(
            model=MODEL,
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            tools=TOOLS,
            messages=messages
        )

        if response.stop_reason == "end_turn":
            return next(
                (block.text for block in response.content if block.type == "text"),
                "（回答なし）"
            )

        messages.append({"role": "assistant", "content": response.content})
        tool_results = []
        for block in response.content:
            if block.type == "tool_use":
                func = TOOL_FUNCTIONS.get(block.name)
                result = func(block.input) if func else f"ツール '{block.name}' が見つかりません。"
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": result
                })

        messages.append({"role": "user", "content": tool_results})


# =====================
# メインループ
# =====================
def main():
    print("=" * 55)
    print("  SNK ダイカスト事業部 月次損益 AI エージェント")
    print("=" * 55)
    print("終了: 'exit' または Ctrl+C\n")
    print("【使い方】")
    print("  ・エクセルのデータをそのまま貼り付けると自動更新")
    print("  ・「11月のサマリーを見せて」などの質問もOK")
    print("-" * 55 + "\n")

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
