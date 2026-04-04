/**
 * 更新型一覧表（Excel）からマスターデータを抽出し、
 * 金型在庫管理_standalone.html に直接埋め込むスクリプト
 *
 * 更新型ダッシュボードと同じ仕組み:
 * OneDrive同期Excel → HTML内のデータ部分を書き換え
 */
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// === 設定 ===
const EXCEL_PATH = path.join(
    'C:', 'Users', 'matsuoka',
    'OneDrive - 新日本金属工業株式会社',
    '技術部進捗管理部屋 - 修正',
    '【技術】更新型進捗管理一覧表.xlsx'
);
const HTML_PATH = path.join(__dirname, '金型在庫管理_standalone.html');
const SHEET_NAME = '更新型';

function main() {
    console.log('--- 金型在庫管理 マスターデータ更新 ---');
    console.log('Excel: ' + EXCEL_PATH);

    // Excel存在チェック
    if (!fs.existsSync(EXCEL_PATH)) {
        console.error('エラー: Excelファイルが見つかりません。OneDrive同期を確認してください。');
        process.exit(1);
    }

    // Excel読み込み
    const wb = XLSX.readFile(EXCEL_PATH);
    const ws = wb.Sheets[SHEET_NAME];
    if (!ws) {
        console.error('エラー: シート「' + SHEET_NAME + '」が見つかりません。');
        process.exit(1);
    }

    const rawData = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false });
    const rows = rawData.slice(2); // ヘッダー2行スキップ

    // データ変換
    const data = [];
    for (const r of rows) {
        const no = r[0] != null ? String(r[0]).trim() : '';
        if (!no || no === '管理番号') continue;
        const hinban = String(r[6] || '').trim();
        const hinmei = String(r[7] || '').trim().replace(/\n/g, ' ');

        data.push({
            no,
            customer: String(r[3] || '').trim(),
            hinban,
            hinmei,
            koushinNo: String(r[8] || '').trim().replace(/\n/g, ' '),
            maker: String(r[9] || '').trim(),
            status: String(r[12] || '').trim(),
        });
    }

    console.log('抽出件数: ' + data.length);

    // まず元HTMLからstandalone版を再生成（本体コードの変更も反映）
    const SOURCE_HTML = path.join(__dirname, '金型在庫管理.html');
    if (!fs.existsSync(SOURCE_HTML)) {
        console.error('エラー: 金型在庫管理.html が見つかりません。');
        process.exit(1);
    }

    let html = fs.readFileSync(SOURCE_HTML, 'utf8');
    const masterJs = 'const MOLD_MASTER_DATA = ' + JSON.stringify(data) + ';';
    html = html.replace(
        '<script id="mold-master-embed">/* マスターデータは末尾に埋め込み */</script>',
        '<script>' + masterJs + '</script>'
    );

    // 初期在庫データも埋め込み
    const invPath = path.join(__dirname, 'initial_inventory.json');
    if (fs.existsSync(invPath)) {
        const invData = JSON.parse(fs.readFileSync(invPath, 'utf8'));
        const invJs = 'const INITIAL_INVENTORY = ' + JSON.stringify(invData) + ';';
        html = html.replace(
            '<script id="initial-inventory-embed">/* 初期在庫データは末尾に埋め込み */</script>',
            '<script>' + invJs + '</script>'
        );
        console.log('初期在庫データ埋め込み: ' + Object.keys(invData).length + '件');
    }

    fs.writeFileSync(HTML_PATH, html);

    console.log('金型在庫管理_standalone.html 更新完了 (' + Math.round(fs.statSync(HTML_PATH).size / 1024) + 'KB)');
    console.log('--- 完了 ---');
}

main();
