# 金型在庫管理 Power Apps 構築手順

## 前提
- Microsoft 365 ライセンス（Power Apps含む）
- SharePoint「技術部進捗管理部屋」へのアクセス権

---

## Step 1: SharePointリストを作成

### リスト「ラック在庫」を新規作成

「技術部進捗管理部屋」→「サイトコンテンツ」→「新規」→「リスト」

| カラム名 | 型 | 必須 | 説明 |
|---|---|---|---|
| Title | 1行テキスト | Yes | スロットID（例: `旧Dラック/⑤番扉/左/3段目`） |
| RackName | 選択肢 | Yes | `旧Dラック; 新Dラック; F北側ラック; F南側ラック; G北側ラック` |
| DoorName | 1行テキスト | Yes | 扉名（例: `⑤番扉`） |
| Position | 選択肢 | Yes | `左; 中央; 右` |
| Level | 選択肢 | Yes | `3段目; 2段目; 1段目奥; 1段目手前` |
| UketsukeNo | 1行テキスト | No | 更新型受付No |
| Hinban | 1行テキスト | No | 品番 |
| Hinmei | 1行テキスト | No | 品名 |
| KoushinNo | 1行テキスト | No | 更新型No |
| Customer | 1行テキスト | No | 客先 |
| IsOccupied | はい/いいえ | No | 使用中フラグ（デフォルト: いいえ） |
| LastUpdated | 日付と時刻 | No | 最終更新日時 |
| UpdatedBy | 1行テキスト | No | 更新者 |

### 初期データ投入

全スロットを空の状態で登録する。以下の組み合わせで全件作成:

**旧Dラック**: ⑤④③②①番扉 × 左/中央/右 × 3段目/2段目/1段目奥/1段目手前 = 60件
**新Dラック**: ③②①番扉 × 左/中央/右 × 3段目/2段目/1段目奥/1段目手前 = 36件
**F北側ラック**: ⑦⑥⑤番扉 × 左/中央/右 × 3段目/2段目/1段目奥/1段目手前 = 36件
**F南側ラック**: ③②①番扉 × 左/中央/右 × 3段目/2段目/1段目奥/1段目手前 = 36件
**G北側ラック**: ④③②番扉 × 左/中央/右 × 2段目/1段目奥/1段目手前 = 27件

**合計: 195件**

---

## Step 2: SharePointリスト「更新型マスター」を作成

更新型一覧表のデータを検索するためのリスト。

| カラム名 | 型 | 説明 |
|---|---|---|
| Title | 1行テキスト | 管理番号（受付No） |
| Hinban | 1行テキスト | 品番 |
| Hinmei | 1行テキスト | 品名 |
| KoushinNo | 1行テキスト | 更新型No |
| Customer | 1行テキスト | 客先 |
| Maker | 1行テキスト | 型メーカー |
| Status | 1行テキスト | ステータス |

※ `mold_master.json`（928件）のデータをCSVに変換してインポートする

---

## Step 3: Power Apps アプリを作成

### 3-1. アプリの起動

1. https://make.powerapps.com にアクセス
2. 「作成」→「空のアプリ」→「空のキャンバスアプリ」
3. アプリ名: `金型在庫管理`
4. 形式: **電話**（スマホ向け）

### 3-2. データソース接続

1. 左メニュー「データ」→「データの追加」
2. 「SharePoint」を選択
3. 「技術部進捗管理部屋」サイトを選択
4. 「ラック在庫」と「更新型マスター」の両方を追加

### 3-3. 画面構成（5画面）

#### 画面1: ホーム（メニュー）
- 「入庫」「出庫・検索」「ダッシュボード」「棚卸」の4ボタン
- 各ボタンに Navigate() で画面遷移を設定

#### 画面2: 入庫画面
**上部: 受付No入力**
- TextInput コントロール: `txtUketsukeNo`
- 入力変更時のアクション:
```
UpdateContext({
    selectedMold: LookUp(更新型マスター, Title = txtUketsukeNo.Text)
})
```
- 情報表示ラベル: `selectedMold.Hinban`, `selectedMold.Hinmei`, `selectedMold.KoushinNo`, `selectedMold.Customer`

**中部: ラック選択**
- Gallery コントロール: ラック名ボタンを横並び
- Items: `["旧Dラック","新Dラック","F北側ラック","F南側ラック","G北側ラック"]`
- OnSelect: `UpdateContext({selectedRack: ThisItem.Value})`

**下部: 配置図（スロット選択）**
- Gallery コントロール（入れ子 or HTML表示）
- Items: `Filter(ラック在庫, RackName = selectedRack)`
- 空きスロット（IsOccupied = false）は緑、使用中は赤
- OnSelect:
```
If(
    !ThisItem.IsOccupied,
    Patch(
        ラック在庫,
        ThisItem,
        {
            UketsukeNo: txtUketsukeNo.Text,
            Hinban: selectedMold.Hinban,
            Hinmei: selectedMold.Hinmei,
            KoushinNo: selectedMold.KoushinNo,
            Customer: selectedMold.Customer,
            IsOccupied: true,
            LastUpdated: Now(),
            UpdatedBy: User().FullName
        }
    );
    Notify("入庫完了: " & selectedMold.Hinban, NotificationType.Success)
)
```

#### 画面3: 出庫・検索画面
- 受付No入力 → `Filter(ラック在庫, UketsukeNo = input)` で検索
- 検索結果に保管場所を表示
- 「出庫」ボタン:
```
Patch(
    ラック在庫,
    selectedSlot,
    {
        UketsukeNo: "",
        Hinban: "",
        Hinmei: "",
        KoushinNo: "",
        Customer: "",
        IsOccupied: false,
        LastUpdated: Now(),
        UpdatedBy: User().FullName
    }
)
```

#### 画面4: ダッシュボード
- 各ラックの使用率を表示
- 数式例:
```
CountRows(Filter(ラック在庫, RackName = "旧Dラック" && IsOccupied)) / CountRows(Filter(ラック在庫, RackName = "旧Dラック")) * 100
```
- Chartコントロール（棒グラフ）で可視化

#### 画面5: 棚卸画面
- ラック選択 → スロット一覧表示
- 各スロットに「OK」「NG」ボタン

### 3-4. バーコード/QRスキャン

Power Appsには**バーコードスキャナーコントロール**が標準搭載:
1. 「挿入」→「メディア」→「バーコードスキャナー」
2. OnScan:
```
UpdateContext({scannedValue: BarcodeScanner1.Value});
// 受付No欄に自動入力
UpdateContext({selectedMold: LookUp(更新型マスター, Title = scannedValue)})
```

---

## Step 4: アプリの公開

1. 「ファイル」→「保存」→「発行」
2. 「共有」→ SNK技術部のメンバーを追加
3. スマホで「Power Apps」アプリをインストール → ログイン → 金型在庫管理が表示される

---

## Step 5: マスターデータのインポート

### CSVでの一括インポート手順

1. `mold_master.json` を CSVに変換（変換スクリプトを使用）
2. SharePointリスト「更新型マスター」→「クイック編集」→ Excelからコピペ
   または Power Automate でJSONから自動投入

---

## 補足: HTMLアプリからの移行ポイント

| HTMLアプリの機能 | Power Appsでの実現方法 |
|---|---|
| LocalStorage保存 | SharePointリスト |
| 配置図タップ式 | Gallery + 条件付き書式 |
| 受付No自動検索 | LookUp関数 |
| QRスキャン | バーコードスキャナーコントロール |
| ダッシュボード | Chartコントロール |
| オフライン対応 | Power Appsオフラインモード（自動キャッシュ） |
