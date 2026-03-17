# 11_tests — テストフォルダ

小規模事業者向けAI支援ツールの自動テスト集。

## 実行方法

```bash
# 全テスト実行
pytest 11_tests/ -v

# 特定ファイルのみ
pytest 11_tests/test_01_sns_post_generation.py -v
```

## テスト一覧

| ファイル | 機能 | テスト数 |
|---|---|---|
| test_01_sns_post_generation.py | SNS投稿文の自動生成 | 5 |
| test_02_invoice_generation.py | 請求書・見積書の自動生成 | 5 |
| test_03_expense_summary.py | 経費・売上の集計補助 | 4 |

**合計: 14テスト（うち11テストがコアテスト）**

## テスト番号対応

- テスト01〜05: SNS投稿文生成
- テスト06〜10: 請求書生成
- テスト11〜: 経費集計
