"""
テスト03: 経費・売上の集計補助
対象: CSVデータの読み込みと簡易レポート生成機能
"""
import pytest
import io
import csv


def parse_expense_csv(csv_text: str) -> list[dict]:
    """経費CSVをパースしてリストで返す"""
    reader = csv.DictReader(io.StringIO(csv_text))
    return [{"date": row["日付"], "category": row["カテゴリ"], "amount": int(row["金額"])} for row in reader]


def summarize_by_category(expenses: list[dict]) -> dict:
    """カテゴリ別に合計する"""
    summary = {}
    for expense in expenses:
        cat = expense["category"]
        summary[cat] = summary.get(cat, 0) + expense["amount"]
    return summary


SAMPLE_CSV = """日付,カテゴリ,金額
2026-03-01,食材,5000
2026-03-02,光熱費,8000
2026-03-03,食材,3000
2026-03-04,消耗品,1500
"""


# テスト11: CSVが正しくパースされる
def test_csv_parsed_correctly():
    expenses = parse_expense_csv(SAMPLE_CSV)
    assert len(expenses) == 4
    assert expenses[0]["amount"] == 5000


def test_category_summary_correct():
    expenses = parse_expense_csv(SAMPLE_CSV)
    summary = summarize_by_category(expenses)
    assert summary["食材"] == 8000
    assert summary["光熱費"] == 8000
    assert summary["消耗品"] == 1500


def test_total_expense_calculation():
    expenses = parse_expense_csv(SAMPLE_CSV)
    total = sum(e["amount"] for e in expenses)
    assert total == 17500


def test_empty_csv_returns_empty_list():
    empty_csv = "日付,カテゴリ,金額\n"
    expenses = parse_expense_csv(empty_csv)
    assert expenses == []
