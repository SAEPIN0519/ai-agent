# SharePoint リスト定義書 — 金型在庫管理システム

## リストA：「更新型一覧表」（既存リストへの追加カラム）

既存の更新型マスターデータに以下のカラムを追加する。

| カラム名 | 内部名 | 型 | 必須 | 説明 |
|---|---|---|---|---|
| CurrentLocation | CurrentLocation | 1行テキスト | No | 現在の保管場所（例: `D1-A3`）。空欄＝在庫なし |
| QRCodeURL | QRCodeURL | ハイパーリンク | No | アプリへのディープリンク（自動生成） |
| LastMovedDate | LastMovedDate | 日付と時刻 | No | 最後に入庫/出庫した日時 |
| LastMovedBy | LastMovedBy | 1行テキスト | No | 最後に操作した担当者名 |

### 運用ルール
- `CurrentLocation` が空欄 → 金型は出庫済み（ラック外）
- `CurrentLocation` に値あり → 金型はラック内に保管中
- `QRCodeURL` は金型登録時に自動生成される

---

## リストB：「ラック在庫」（新規作成）

ラックの各番地と金型の紐付けを管理するリスト。

| カラム名 | 内部名 | 型 | 必須 | 説明 |
|---|---|---|---|---|
| Title | Title | 1行テキスト | Yes | 一意キー（例: `D1-A3`） |
| RackName | RackName | 選択肢 | Yes | ラック名: D1, D2, F1, F2, S1, S2 |
| Address | Address | 1行テキスト | Yes | 番地（例: A1, A2, B1...） |
| MoldNo | MoldNo | 1行テキスト | No | 格納中の金型No。空欄＝空き番地 |
| MoldInfo | MoldInfo | 複数行テキスト | No | 金型の補足情報（客先・品番等） |
| LastUpdated | LastUpdated | 日付と時刻 | No | 最終更新日時 |
| UpdatedBy | UpdatedBy | 1行テキスト | No | 更新した担当者名 |

### インデックス設定
- `Title`（既定）
- `RackName` + `Address` の複合ビュー
- `MoldNo` でフィルタ用ビュー

### 初期データ投入

各ラックの番地をあらかじめ全件登録しておく（MoldNoは空欄）。

```
ラック名: D1, D2, F1, F2, S1, S2
番地例: A1, A2, A3, ... B1, B2, B3, ...
```

番地の命名規則は現場のラック構成に合わせてカスタマイズする。

---

## リスト作成手順（SharePoint管理画面）

### リストB「ラック在庫」の作成

1. SharePointサイト → 「サイトコンテンツ」→「新規」→「リスト」
2. リスト名: `ラック在庫`（内部名: `RackInventory`）
3. 上記カラムを順に追加
4. `RackName` の選択肢: `D1; D2; F1; F2; S1; S2`
5. ビュー「ラック別」を作成: RackName でグループ化

### リストA「更新型一覧表」への追加

1. 既存リストの設定 → 「列の追加」
2. `CurrentLocation`, `QRCodeURL`, `LastMovedDate`, `LastMovedBy` を追加
3. 既存ビューに `CurrentLocation` 列を表示追加

---

## Graph API エンドポイント

```
サイトID取得:
GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site-path}

リスト取得:
GET https://graph.microsoft.com/v1.0/sites/{siteId}/lists

リストアイテム取得:
GET https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items?expand=fields

アイテム作成:
POST https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items

アイテム更新:
PATCH https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items/{itemId}/fields
```

---

## Azure AD アプリ登録（Graph API認証用）

1. Azure Portal → Azure Active Directory → アプリの登録 → 新規登録
2. 名前: `金型在庫管理アプリ`
3. リダイレクトURI: `https://{SharePointサイトURL}` （SPAを選択）
4. APIのアクセス許可:
   - `Sites.ReadWrite.All`（委任）
   - `User.Read`（委任）
5. クライアントID をアプリの設定に記入
