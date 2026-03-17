"""
糸貫グループ 経営ダッシュボード 月次更新スクリプト
======================================================
毎月の損益データを入力し、ダッシュボードHTMLを自動更新します。

使い方:
    python update_dashboard.py

動作:
    - 現在の4ヶ月データの最も古い月を削除し、新しい月のデータを追加
    - HTMLファイルを上書き保存（バックアップも作成）
"""

import re
import sys
import shutil
from datetime import datetime
from pathlib import Path

# ────── 設定 ──────
HTML_PATH = Path(__file__).parent.parent / "経営ダッシュボード_糸貫グループ (1) (1) (4).html"
BACKUP_DIR = Path(__file__).parent.parent / "09_system" / "dashboard_backups"

# 各部門の入力項目定義
DIVISIONS = {
    "糸貫工場": ["売上", "原価", "粗利", "営業利益",
                 "基本給", "時間外", "雑給", "賞与引当", "法定福利", "退職金",
                 "消耗品", "修繕費", "地代家賃", "減価償却", "電力", "燃料"],
    "加工":     ["売上", "原価", "粗利", "営業利益",
                 "基本給", "時間外", "雑給", "賞与引当", "法定福利", "退職金",
                 "消耗品", "減価償却", "電力", "燃料"],
    "糸貫検査": ["売上", "原価", "粗利", "営業利益",
                 "基本給", "時間外", "雑給", "賞与引当", "法定福利", "退職金",
                 "消耗品", "減価償却", "電力", "燃料"],
    "本社検査": ["売上", "原価", "粗利", "営業利益",
                 "基本給", "時間外", "雑給", "賞与引当", "法定福利", "退職金",
                 "消耗品", "減価償却", "電力", "燃料"],
    "物流検査": ["売上", "原価", "粗利", "営業利益",
                 "基本給", "時間外", "雑給", "賞与引当", "法定福利", "退職金",
                 "消耗品", "減価償却", "電力", "燃料"],
}


def read_html() -> str:
    """HTMLファイルを読み込む"""
    if not HTML_PATH.exists():
        print(f"エラー: ファイルが見つかりません\n  {HTML_PATH}")
        sys.exit(1)
    return HTML_PATH.read_text(encoding="utf-8")


def backup_html(html: str):
    """HTMLファイルのバックアップを作成"""
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"dashboard_backup_{ts}.html"
    backup_path.write_text(html, encoding="utf-8")
    print(f"  バックアップ: {backup_path.name}")


def parse_month_arrays(html: str) -> tuple[list, list]:
    """
    const M と const Ms の配列を取得する
    戻り値: (M配列, Ms配列)  例: (['R7.10月',...], ['10月',...])
    """
    m_match = re.search(r"const M\s*=\s*\[([^\]]+)\]", html)
    ms_match = re.search(r"const Ms\s*=\s*\[([^\]]+)\]", html)
    if not m_match or not ms_match:
        print("エラー: 月配列が見つかりません")
        sys.exit(1)

    def parse_str_array(raw: str) -> list:
        return [s.strip().strip("'\"") for s in raw.split(",")]

    return parse_str_array(m_match.group(1)), parse_str_array(ms_match.group(1))


def parse_division_data(html: str, div: str, metric: str) -> list[int]:
    """
    D[div][metric] の配列を取得する
    例: parse_division_data(html, '糸貫工場', '売上') → [195656503, ...]
    """
    # 部門ブロックを取得
    pattern = rf"{re.escape(div)}:\s*\{{(.*?)\}},"
    div_match = re.search(pattern, html, re.DOTALL)
    if not div_match:
        return []

    div_block = div_match.group(1)

    # メトリクス行を取得
    metric_pattern = rf"{re.escape(metric)}:\s*\[([^\]]+)\]"
    metric_match = re.search(metric_pattern, div_block)
    if not metric_match:
        return []

    raw = metric_match.group(1)
    return [int(v.strip()) for v in raw.split(",")]


def update_month_arrays(html: str, old_m: list, old_ms: list,
                         new_m: str, new_ms: str) -> str:
    """const M と const Ms に新しい月を追加（最古の月を削除）"""
    updated_m = old_m[1:] + [new_m]
    updated_ms = old_ms[1:] + [new_ms]

    m_str = ", ".join(f"'{x}'" for x in updated_m)
    ms_str = ", ".join(f"'{x}'" for x in updated_ms)

    html = re.sub(r"const M\s*=\s*\[[^\]]+\]", f"const M = [{m_str}]", html)
    html = re.sub(r"const Ms\s*=\s*\[[^\]]+\]", f"const Ms = [{ms_str}]", html)
    return html


def update_metric_array(html: str, div: str, metric: str,
                         old_vals: list[int], new_val: int) -> str:
    """
    D[div][metric] の配列を更新（最古の値を削除、新しい値を追加）
    """
    updated = old_vals[1:] + [new_val]
    vals_str = ", ".join(f"{v:>12}" for v in updated)

    # 部門ブロック内のメトリクス行を置換
    old_vals_str = ", ".join(str(v) for v in old_vals)

    # より確実な置換：部門ブロックを特定してから置換
    div_pattern = rf"({re.escape(div)}:\s*\{{)(.*?)(\}},)"
    def replace_in_div(m):
        prefix, block, suffix = m.group(1), m.group(2), m.group(3)
        metric_pattern = rf"({re.escape(metric)}:\s*\[)[^\]]+(\])"
        new_block = re.sub(
            metric_pattern,
            rf"\g<1>{vals_str}\2",
            block
        )
        return prefix + new_block + suffix

    html = re.sub(div_pattern, replace_in_div, html, flags=re.DOTALL)
    return html


def update_header(html: str, new_m_list: list, latest_month: str) -> str:
    """ヘッダーの期間表示と最新月を更新"""
    # 最古月と最新月
    first = new_m_list[0]
    last  = new_m_list[-1]

    html = re.sub(
        r"糸貫工場 / 加工 / 糸貫検査 / 本社検査 / 物流検査｜[^<]+",
        f"糸貫工場 / 加工 / 糸貫検査 / 本社検査 / 物流検査｜{first}〜{last}",
        html
    )
    html = re.sub(
        r"<strong>最新月：[^<]+</strong>",
        f"<strong>最新月：{latest_month}</strong>",
        html
    )
    return html


def get_input_int(prompt: str, allow_empty: bool = False) -> int | None:
    """整数入力を取得（空欄→0）"""
    while True:
        raw = input(prompt).strip().replace(",", "").replace("，", "")
        if raw == "" and allow_empty:
            return None
        if raw == "":
            return 0
        try:
            return int(float(raw))  # 小数点入力にも対応
        except ValueError:
            print("  ※ 数値を入力してください（未入力→0）")


def main():
    print("=" * 60)
    print("  糸貫グループ 経営ダッシュボード 月次更新ツール")
    print("=" * 60)

    # HTMLを読み込む
    html = read_html()
    m_list, ms_list = parse_month_arrays(html)

    print(f"\n現在の期間: {m_list[0]} 〜 {m_list[-1]}")
    print(f"更新後: {m_list[1]} 〜 新しい月  （{m_list[0]} を削除）\n")

    # 新しい月ラベルを入力
    print("【新しい月のラベルを入力してください】")
    print("  例: R8.2月  →  表示ラベル（長）")
    new_m  = input("  長ラベル (例: R8.2月): ").strip()
    new_ms = input("  短ラベル (例: 2月):    ").strip()
    if not new_m or not new_ms:
        print("エラー: 月ラベルを入力してください。")
        sys.exit(1)

    # 最新月の表示（ヘッダー用）
    latest_label = input(f"  ヘッダー表示 (例: R8年2月、空欄→{new_m}): ").strip()
    if not latest_label:
        latest_label = new_m

    print()

    # 各部門のデータ入力
    new_data: dict[str, dict[str, int]] = {}

    for div, metrics in DIVISIONS.items():
        print(f"{'─'*50}")
        print(f"  【{div}】")
        print(f"  ※ 未入力(Enterのみ) = 0円\n")
        new_data[div] = {}
        for metric in metrics:
            val = get_input_int(f"    {metric}: ")
            new_data[div][metric] = val

    # バックアップ作成
    print(f"\n{'─'*50}")
    print("バックアップを作成中...")
    backup_html(html)

    # M / Ms 更新
    html = update_month_arrays(html, m_list, ms_list, new_m, new_ms)
    new_m_list = m_list[1:] + [new_m]

    # 各部門・各メトリクスを更新
    print("データを更新中...")
    for div, metrics in DIVISIONS.items():
        for metric in metrics:
            old_vals = parse_division_data(html, div, metric)
            if not old_vals:
                continue
            new_val = new_data[div].get(metric, 0)
            html = update_metric_array(html, div, metric, old_vals, new_val)

    # ヘッダー更新
    html = update_header(html, new_m_list, latest_label)

    # 保存
    HTML_PATH.write_text(html, encoding="utf-8")
    print(f"\n✅ 更新完了！")
    print(f"   ファイル: {HTML_PATH.name}")
    print(f"   期間: {new_m_list[0]} 〜 {new_m_list[-1]}")


if __name__ == "__main__":
    main()
