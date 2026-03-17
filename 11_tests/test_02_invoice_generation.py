"""
テスト02: 請求書・見積書の自動生成
対象: テンプレートに顧客情報・金額を自動入力して書類を生成する機能
"""
import pytest
from datetime import date


def generate_invoice(client_name: str, items: list[dict], tax_rate: float = 0.1) -> dict:
    """請求書データを生成するダミー関数"""
    subtotal = sum(item["unit_price"] * item["quantity"] for item in items)
    tax = int(subtotal * tax_rate)
    total = subtotal + tax
    return {
        "client_name": client_name,
        "items": items,
        "subtotal": subtotal,
        "tax": tax,
        "total": total,
        "issue_date": str(date.today()),
    }


# テスト6: 請求書の合計金額が正しく計算される
def test_invoice_total_is_correct():
    items = [{"name": "コーヒー豆", "unit_price": 1000, "quantity": 3}]
    invoice = generate_invoice("テスト商店", items)
    assert invoice["subtotal"] == 3000
    assert invoice["tax"] == 300
    assert invoice["total"] == 3300


# テスト7: 複数品目の合計が正しく計算される
def test_invoice_multiple_items():
    items = [
        {"name": "ドリップバッグ", "unit_price": 500, "quantity": 10},
        {"name": "フィルター", "unit_price": 200, "quantity": 5},
    ]
    invoice = generate_invoice("カフェA", items)
    assert invoice["subtotal"] == 6000
    assert invoice["total"] == 6600


# テスト8: 顧客名が請求書に含まれる
def test_invoice_includes_client_name():
    invoice = generate_invoice("田中商店", [{"name": "商品", "unit_price": 100, "quantity": 1}])
    assert invoice["client_name"] == "田中商店"


# テスト9: 発行日が自動設定される
def test_invoice_issue_date_is_set():
    invoice = generate_invoice("テスト", [{"name": "商品", "unit_price": 100, "quantity": 1}])
    assert invoice["issue_date"] == str(date.today())


# テスト10: 品目が空の場合は合計ゼロ
def test_invoice_empty_items_total_zero():
    invoice = generate_invoice("テスト商店", [])
    assert invoice["subtotal"] == 0
    assert invoice["total"] == 0
