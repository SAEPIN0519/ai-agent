# 操作マニュアル自動生成スキル

## 概要
Puppeteer（ヘッドレスブラウザ）でWebアプリを自動操作し、各操作画面のスクショを撮影。
スクショ付きのHTML操作マニュアルを自動生成する。ブラウザのCtrl+PでPDF変換可能。

## 使いどころ
- HTMLアプリの操作マニュアルを作りたい時
- クライアントへの納品物として手順書が必要な時
- アプリ更新のたびにスクショを撮り直したい時（スクリプト再実行で全自動）

## 必要なもの
- Node.js
- Puppeteer（`npm install puppeteer`）

## 手順

### 1. スクショ撮影スクリプトを作る

```javascript
const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');

const APP_PATH = path.join(__dirname, 'アプリ.html');
const SCREENSHOT_DIR = path.join(__dirname, 'manual_screenshots');

async function main() {
    if (!fs.existsSync(SCREENSHOT_DIR)) fs.mkdirSync(SCREENSHOT_DIR);

    const browser = await puppeteer.launch({
        headless: false,  // true にすると画面非表示で高速
        defaultViewport: { width: 500, height: 900 }, // スマホサイズ
    });

    const page = await browser.newPage();
    await page.goto('file:///' + APP_PATH.replace(/\\/g, '/'), { waitUntil: 'networkidle2' });

    // デモデータをLocalStorageにセット（必要に応じて）
    await page.evaluate((data) => {
        localStorage.setItem('キー名', JSON.stringify(data));
    }, デモデータ);
    await page.reload({ waitUntil: 'networkidle2' });

    // スクショ撮影: 操作→スクショ→操作→スクショ の繰り返し
    // 例: 入力欄にテキスト入力
    await page.type('#inputId', '値', { delay: 100 });
    await sleep(500);
    await page.screenshot({ path: path.join(SCREENSHOT_DIR, '01_入力後.png') });

    // 例: ボタンクリック
    await page.click('#buttonId');
    await sleep(500);
    await page.screenshot({ path: path.join(SCREENSHOT_DIR, '02_クリック後.png') });

    // 例: JS関数を直接実行
    await page.evaluate(() => { アプリの関数(); });
    await sleep(500);
    await page.screenshot({ path: path.join(SCREENSHOT_DIR, '03_関数実行後.png') });

    await browser.close();
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
main().catch(console.error);
```

### 2. マニュアルHTMLを作る

```html
<div class="step-box">
    <span class="step-num">1</span>
    <span class="step-title">操作の説明</span>
    <p>詳細な手順をここに書く</p>
</div>
<div class="screenshot">
    <img src="manual_screenshots/01_入力後.png" alt="説明">
</div>
```

### 3. PDF変換
- ブラウザでマニュアルHTMLを開く
- Ctrl+P → 「PDFとして保存」
- `@media print` でCSSを調整しておくと綺麗に出力される

## ポイント
- `headless: false` で実際のブラウザが開くのでデバッグしやすい
- `page.evaluate()` でアプリ内のJS関数を直接呼べる（ボタンクリックより確実）
- `sleep()` を入れないとアニメーション中にスクショが撮れる
- `defaultViewport` でスマホ/PC両方のスクショが撮れる
- `@page { size: A4; margin: 15mm; }` でPDF化時の用紙サイズを指定

## 実績
- SNK金型在庫管理システム: 13枚のスクショを全自動撮影（2026-04-03）
