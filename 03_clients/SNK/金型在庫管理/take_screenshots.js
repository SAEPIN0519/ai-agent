/**
 * Puppeteerで金型在庫管理アプリの操作画面スクショを自動撮影
 * マニュアル用の画像を生成する
 */
const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');

const APP_PATH = path.join(__dirname, '金型在庫管理_standalone.html');
const SCREENSHOT_DIR = path.join(__dirname, 'manual_screenshots');

// デモ用の在庫データ
const DEMO_INVENTORY = {
    "旧Dラック/⑤番扉/左/3段目": { moldNo: "1", hinban: "44", hinmei: "スリーブ", koushinNo: "5 17～20", customer: "名光", updatedAt: "2026-04-03" },
    "旧Dラック/⑤番扉/中央/2段目": { moldNo: "5", hinban: "460", hinmei: "スリーブ", koushinNo: "5 17～20", customer: "光", split: true, item2: { moldNo: "6", hinban: "460", hinmei: "スリーブ", koushinNo: "6 21～25", customer: "光" }, updatedAt: "2026-04-03" },
    "旧Dラック/⑤番扉/左/1段目奥": { moldNo: "2", hinban: "TG-719-02", hinmei: "PALTE", koushinNo: "3", customer: "住友理工", updatedAt: "2026-04-03" },
    "旧Dラック/④番扉/中央/3段目": { moldNo: "192", hinban: "ST-6800", hinmei: "巻き取り体R", koushinNo: "5", customer: "シマノ", updatedAt: "2026-04-03" },
    "新Dラック/③番扉/右/2段目": { moldNo: "M-1712345", hinban: "", hinmei: "NA11 インゴット", koushinNo: "リターン材", customer: "手動登録", updatedAt: "2026-04-03" },
};

async function main() {
    if (!fs.existsSync(SCREENSHOT_DIR)) fs.mkdirSync(SCREENSHOT_DIR, { recursive: true });

    console.log('ブラウザ起動中...');
    const browser = await puppeteer.launch({
        headless: false, // 画面表示あり（デバッグ用）
        defaultViewport: { width: 500, height: 900 }, // スマホサイズ
        args: ['--no-sandbox']
    });

    const page = await browser.newPage();
    const fileUrl = 'file:///' + APP_PATH.replace(/\\/g, '/');
    await page.goto(fileUrl, { waitUntil: 'networkidle2', timeout: 30000 });

    // デモデータをLocalStorageにセット
    await page.evaluate((data) => {
        localStorage.setItem('moldInventory', JSON.stringify(data));
    }, DEMO_INVENTORY);
    await page.reload({ waitUntil: 'networkidle2' });
    await sleep(1000);

    // ===== 1. トップ画面（入庫タブ） =====
    console.log('1. トップ画面');
    await screenshot(page, '01_top.png');

    // ===== 2. 受付No入力後の品番表示 =====
    console.log('2. 受付No入力');
    await page.type('#receiveUketsuke', '192', { delay: 100 });
    await sleep(500);
    await screenshot(page, '02_uketsuke_input.png');

    // ===== 3. ラック選択 =====
    console.log('3. ラック選択');
    await page.evaluate(() => { window.scrollTo(0, 300); });
    await sleep(300);
    // 旧Dラックをクリック
    const rackCards = await page.$$('.rack-card');
    if (rackCards.length > 0) await rackCards[0].click();
    await sleep(500);
    await screenshot(page, '03_rack_select.png');

    // ===== 4. 配置図（空きと使用中混在） =====
    console.log('4. 配置図');
    await page.evaluate(() => { window.scrollTo(0, 500); });
    await sleep(300);
    await screenshot(page, '04_rack_map.png');

    // ===== 5. 空きセルクリック → 1個/2個選択 =====
    console.log('5. 個数選択モーダル');
    const emptyCells = await page.$$('.slot-cell.empty');
    if (emptyCells.length > 0) await emptyCells[0].click();
    await sleep(500);
    await screenshot(page, '05_count_modal.png');

    // ===== 6. 「1個」選択 → セル青ハイライト =====
    console.log('6. 1個選択');
    // 1個ボタンをクリック
    await page.evaluate(() => { confirmSlotCount(1); });
    await sleep(500);
    await page.evaluate(() => { window.scrollTo(0, 500); });
    await sleep(300);
    await screenshot(page, '06_selected.png');

    // ===== 7. 入庫完了後 =====
    console.log('7. 入庫完了');
    await page.evaluate(() => { window.scrollTo(0, 0); });
    await sleep(300);
    await page.evaluate(() => { handleReceive(); });
    await sleep(500);
    await page.evaluate(() => { window.scrollTo(0, 400); });
    await sleep(300);
    await screenshot(page, '07_receive_done.png');

    // ===== 8. 手動入力モード =====
    console.log('8. 手動入力');
    await page.evaluate(() => { window.scrollTo(0, 0); });
    await sleep(300);
    await page.evaluate(() => { switchInputMode('manual'); });
    await sleep(300);
    await screenshot(page, '08_manual_input.png');

    // ===== 9. 2個置きの入庫（分割セル） =====
    console.log('9. 2個入りセル');
    await page.evaluate(() => { window.scrollTo(0, 500); });
    await sleep(300);
    await screenshot(page, '09_split_cell.png');

    // ===== 10. 使用中セルタップ → メニュー =====
    console.log('10. スロットメニュー');
    const occupiedCells = await page.$$('.slot-cell.occupied');
    if (occupiedCells.length > 0) await occupiedCells[0].click();
    await sleep(500);
    await screenshot(page, '10_slot_menu.png');

    // メニュー閉じる
    await page.evaluate(() => { closeSlotMenu(); });
    await sleep(300);

    // ===== 11. 出庫・検索タブ =====
    console.log('11. 出庫・検索');
    await page.evaluate(() => { switchTab('ship'); });
    await sleep(300);
    await screenshot(page, '11_ship_tab.png');

    // ===== 12. 2個入りの出庫メニュー =====
    console.log('12. 2個入り出庫メニュー');
    // 分割セルを探してメニュー表示
    await page.evaluate(() => {
        switchTab('receive');
        // 旧Dラックを選択
        const cards = document.querySelectorAll('.rack-card');
        if (cards[0]) cards[0].click();
    });
    await sleep(500);
    await page.evaluate(() => {
        // 分割セルを見つけてメニュー表示
        const splitCells = document.querySelectorAll('[data-split="true"]');
        if (splitCells.length > 0) splitCells[0].click();
    });
    await sleep(500);
    await screenshot(page, '12_split_menu.png');

    // ===== 13. 2個目入力待ち（黄色帯） =====
    console.log('13. 2個目待ち');
    await page.evaluate(() => { closeSlotMenu(); });
    await sleep(200);
    // 通常セルをタップして2分割
    await page.evaluate(() => {
        const cells = document.querySelectorAll('.slot-cell.occupied');
        for (const c of cells) {
            if (!c.dataset.split) { showSlotMenu(c); break; }
        }
    });
    await sleep(500);
    await page.evaluate(() => { slotMenuAction('split'); });
    await sleep(500);
    await page.evaluate(() => { window.scrollTo(0, 0); });
    await sleep(300);
    await screenshot(page, '13_waiting_item2.png');

    console.log('\n===== 完了 =====');
    console.log(`${SCREENSHOT_DIR} にスクショが保存されました`);

    await browser.close();
}

async function screenshot(page, filename) {
    await page.screenshot({
        path: path.join(SCREENSHOT_DIR, filename),
        fullPage: false
    });
    console.log(`  ✓ ${filename}`);
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

main().catch(e => { console.error('エラー:', e.message); process.exit(1); });
