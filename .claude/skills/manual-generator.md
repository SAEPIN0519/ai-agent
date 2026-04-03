---
description: HTMLアプリの操作マニュアルをスクショ付きで自動生成する。「マニュアル作って」「手順書作って」「説明書作って」で起動。
---

# 操作マニュアル自動生成

## 実行フロー

1. **対象アプリの確認**: どのHTMLアプリのマニュアルを作るか確認
2. **Puppeteerスクショスクリプト作成**:
   - 対象アプリをPuppeteerで開く
   - デモデータをLocalStorageにセット
   - 各操作画面を自動操作してスクショを撮影
   - `manual_screenshots/` フォルダに保存
3. **マニュアルHTML生成**:
   - 表紙・目次・各操作手順をステップ形式で記載
   - スクショを各ステップに埋め込み
   - `@media print` でPDF変換対応のCSS設定
   - 「PDF保存/印刷する」ボタンを設置
4. **ブラウザで開いてCtrl+PでPDF保存**

## Puppeteerスクショの撮り方

```javascript
const puppeteer = require('puppeteer');

const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: { width: 500, height: 900 },
});
const page = await browser.newPage();
await page.goto('file:///パス/アプリ.html', { waitUntil: 'networkidle2' });

// デモデータセット
await page.evaluate((data) => {
    localStorage.setItem('キー', JSON.stringify(data));
}, demoData);
await page.reload({ waitUntil: 'networkidle2' });

// 操作してスクショ
await page.type('#input', '値', { delay: 100 });
await page.screenshot({ path: 'manual_screenshots/01_name.png' });

// JS関数を直接実行
await page.evaluate(() => { appFunction(); });
await page.screenshot({ path: 'manual_screenshots/02_name.png' });

await browser.close();
```

## マニュアルHTMLのテンプレート構成

- 表紙（タイトル・日付・部署名）
- 目次（アンカーリンク付き）
- 各章:
  - ステップ番号付きの操作手順（`.step-box` + `.step-num`）
  - スクショ画像（`<img src="manual_screenshots/xx.png">`）
  - Tips（`.tip-box`）/ 注意（`.warn-box`）
- よくある質問（FAQ）
- PDF保存ボタン

## 必要パッケージ
- `puppeteer`（`npm install puppeteer`）

## 参考スキル
- `05_skills/technical/auto-manual-generator.md` に詳細なコードサンプルあり
