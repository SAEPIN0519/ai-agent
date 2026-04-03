# Office 365 社内アプリケーション開発スキル

## 概要
Microsoft 365（SharePoint / Power Apps / Graph API）を活用して、社内限定のWebアプリを開発・配置するスキル。
外部公開不要でセキュア、スマホ対応、追加コストなし。

## 前提条件
- クライアントがMicrosoft 365を契約していること
- SharePointサイトへのアクセス権があること

---

## アーキテクチャパターン

### パターンA: HTMLアプリ + SharePoint配置（SPFx）
**最もUI自由度が高い方式。** HTML/CSS/JSで作ったアプリをSharePoint上で動かす。

| 項目 | 内容 |
|---|---|
| フロントエンド | HTML/CSS/JavaScript（1ファイル完結推奨） |
| データ保存 | LocalStorage（単体）or SharePointリスト（共有） |
| 配置方法 | SPFxパッケージ(.sppkg)としてアプリカタログに登録 |
| 認証 | SharePointの認証がそのまま適用 |
| スマホ対応 | SharePointページ経由でアクセス可能 |

**注意:** SPFxパッケージのアップロードにはテナント管理者 or アプリカタログ管理者の権限が必要。

### パターンB: Power Apps
**ノーコード/ローコードで構築。** SharePointリストをバックエンドにする。

| 項目 | 内容 |
|---|---|
| UI構築 | Power Apps Studio（GUI操作） |
| データ保存 | SharePointリスト |
| 配置方法 | Power Appsで発行→共有 |
| 認証 | Microsoft 365アカウント |
| スマホ対応 | Power Appsモバイルアプリ |

**注意:** 複雑なUIレイアウト（グリッド配置図等）は苦手。Copilotはコントロール自動配置不可。

### パターンC: SharePointリスト + ビュー
**最もシンプル。** SharePointリストの標準機能だけで管理。

| 項目 | 内容 |
|---|---|
| UI構築 | SharePointリストのビュー・フォームカスタマイズ |
| データ保存 | SharePointリスト |
| 配置方法 | 設定不要 |
| 認証 | SharePointの認証 |
| スマホ対応 | SharePointモバイルアプリ |

---

## SharePointリストの作成

### Excelからインポート
SharePointの「新規」→「リスト」→「Excelから」でインポート可能。
**重要:** Excelにテーブル定義が必要。ExcelJSの `addTable()` でテーブルを追加してからインポートする。

```javascript
const ExcelJS = require('exceljs');
const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet('シート名');
ws.addTable({
    name: 'TableName',
    ref: 'A1',
    headerRow: true,
    columns: [
        { name: 'Title' },
        { name: 'Column1' },
        // ...
    ],
    rows: [
        ['値1', '値2'],
        // ...
    ],
});
await wb.xlsx.writeFile('output.xlsx');
```

### CSVインポート
CSVは直接インポートできないが、Excelに変換すれば可能。BOM付きUTF-8で文字化け防止。

---

## SPFxパッケージの作成（HTMLアプリ配置用）

### 方法1: 正式な方法（SPFx開発環境）
```bash
npm install -g yo @microsoft/generator-sharepoint
yo @microsoft/sharepoint
# Webパーツを作成 → gulp bundle → gulp package-solution
```

### 方法2: 簡易パッケージ生成（Node.jsスクリプト）
sppkgはZIPファイル。内部にマニフェストJSONとJSバンドルを配置する。
HTMLアプリをiframeで表示するWebパーツとしてパッケージ化する。

必要なファイル:
- `AppManifest.xml` — アプリ定義
- `[Content_Types].xml` — コンテンツタイプ
- `{GUID}/Feature.xml` — 機能定義
- `{GUID}/{WebPartGUID}.manifest.json` — Webパーツマニフェスト
- `{GUID}/bundle.js` — JSバンドル（iframe内にHTMLを展開）

### アプリカタログへのアップロード
1. `https://{tenant}.sharepoint.com/sites/appcatalog` にアクセス
2. 「SharePoint用アプリ」→ ファイルをアップロード
3. 「展開」をクリック
4. 対象サイトの「サイトコンテンツ」→「アプリの追加」で有効化

**権限:** アプリカタログの「SharePoint用アプリ」への投稿権限が必要。閲覧権限だけではアップロード不可。

---

## HTMLアプリのSharePoint配置で学んだこと

### NG方法（動かない）
| 方法 | 結果 |
|---|---|
| ドキュメントライブラリにHTML直アップ | ダウンロードされる |
| .aspx拡張子に変更 | ファイルが見つからないエラー |
| サイトページの「埋め込み」Webパーツ | `<script>`タグ非対応 |
| iframeでドキュメントライブラリのHTMLを参照 | iframe内でもダウンロードされる |
| URLに`?web=1`パラメータ追加 | 効果なし |

### OK方法
| 方法 | 条件 |
|---|---|
| SPFxパッケージとして配置 | アプリカタログ管理者権限が必要 |
| Power Appsでネイティブ構築 | 複雑なUIは困難 |

---

## マスターデータの自動更新

OneDrive同期されたExcelからHTMLアプリのデータを自動更新する仕組み。

### 構成
```
OneDrive同期フォルダ
  └ Excel（更新型一覧表等）
      ↓ xlsxパッケージで読み込み
  Node.jsスクリプト（extract_master.js）
      ↓ HTML内の const DATA = [...]; を書き換え
  standalone.html（配布用）
      ↓ 共有フォルダにコピー
  \\192.168.1.233\...（現場からアクセス）
```

### 自動実行
Windowsタスクスケジューラで毎朝定時実行。
```bat
schtasks /create /tn "タスク名" /tr "cmd /c update.bat auto" /sc daily /st 07:50 /f
```

---

## Power Apps開発メモ

### データソース接続
1. 「データの追加」→ SharePoint → サイト選択 → リスト選択

### 主要な関数
```
LookUp(リスト名, Title = 入力値)          // 1件検索
Filter(リスト名, Column = 値)             // 複数件フィルタ
Patch(リスト名, レコード, {列: 値})        // 更新
CountRows(Filter(リスト名, 条件))         // 件数カウント
Navigate(画面名, ScreenTransition.Fade)   // 画面遷移
```

### Copilotの限界
- 数式の説明・提案はしてくれる
- コントロールの自動配置はしてくれない
- 画面の自動生成も不可
- 手動でコントロールを1つずつ配置する必要がある

---

## セキュリティ比較

| 観点 | SharePoint/Power Apps | Google Drive | GitHub Pages |
|---|---|---|---|
| 認証 | Microsoft 365アカウント必須 | 個人アカウント混在リスク | 認証なし（公開） |
| データ所在 | テナント内 | Google管理 | 外部公開 |
| 退職者対応 | AD無効化で即遮断 | 回収困難 | URLを知っていれば見れる |
| 監査ログ | あり | ビジネス版のみ | なし |

**クライアントデータを含む場合は SharePoint / Power Apps 一択。**

---

## 実績
- SNK金型在庫管理システム（2026-04）
  - SharePointリスト2つ作成（ラック在庫195件 + 更新型マスター928件）
  - SPFxパッケージ生成（IT管理者へのアップロード依頼待ち）
  - Power Apps環境セットアップ（データソース接続済み）
  - HTMLアプリ + OneDrive自動更新で暫定運用中
