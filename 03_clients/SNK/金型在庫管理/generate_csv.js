/**
 * SharePointリストインポート用のCSVを生成するスクリプト
 * 使い方: node generate_csv.js
 *
 * 生成物:
 *   - rack_slots.csv    ... ラック在庫リストの初期データ（195件、全スロット空）
 *   - mold_master.csv   ... 更新型マスターリストのデータ（928件）
 */

const fs = require('path');
const { writeFileSync, readFileSync } = require('fs');
const path = require('path');

// ラック構造定義
const RACK_DEFS = {
    '旧Dラック': {
        doors: ['⑤番扉','④番扉','③番扉','②番扉','①番扉'],
        positions: ['左','中央','右'],
        levels: ['3段目','2段目','1段目奥','1段目手前'],
    },
    '新Dラック': {
        doors: ['③番扉','②番扉','①番扉'],
        positions: ['左','中央','右'],
        levels: ['3段目','2段目','1段目奥','1段目手前'],
    },
    'F北側ラック': {
        doors: ['⑦番扉','⑥番扉','⑤番扉'],
        positions: ['左','中央','右'],
        levels: ['3段目','2段目','1段目奥','1段目手前'],
    },
    'F南側ラック': {
        doors: ['③番扉','②番扉','①番扉'],
        positions: ['左','中央','右'],
        levels: ['3段目','2段目','1段目奥','1段目手前'],
    },
    'G北側ラック': {
        doors: ['④番扉','③番扉','②番扉'],
        positions: ['左','中央','右'],
        levels: ['2段目','1段目奥','1段目手前'],
    },
};

// === ラック在庫CSV生成 ===
function generateRackCSV() {
    const rows = ['Title,RackName,DoorName,Position,Level,UketsukeNo,Hinban,Hinmei,KoushinNo,Customer,IsOccupied'];

    for (const [rackName, def] of Object.entries(RACK_DEFS)) {
        for (const door of def.doors) {
            for (const pos of def.positions) {
                for (const level of def.levels) {
                    const title = `${rackName}/${door}/${pos}/${level}`;
                    rows.push(`"${title}","${rackName}","${door}","${pos}","${level}","","","","","",No`);
                }
            }
        }
    }

    const outPath = path.join(__dirname, 'rack_slots.csv');
    // BOM付きUTF-8（Excelで文字化けしない）
    writeFileSync(outPath, '\ufeff' + rows.join('\n'));
    console.log(`✓ ${outPath}  (${rows.length - 1}件)`);
}

// === 更新型マスターCSV生成 ===
function generateMasterCSV() {
    const data = JSON.parse(readFileSync(path.join(__dirname, 'mold_master.json'), 'utf8'));
    const rows = ['Title,Hinban,Hinmei,KoushinNo,Customer,Maker,Status'];

    for (const item of data) {
        const esc = (s) => `"${(s || '').replace(/"/g, '""').replace(/\n/g, ' ')}"`;
        rows.push([
            esc(item.no),
            esc(item.hinban),
            esc(item.hinmei),
            esc(item.koushinNo),
            esc(item.customer),
            esc(item.maker),
            esc(item.status),
        ].join(','));
    }

    const outPath = path.join(__dirname, 'mold_master.csv');
    writeFileSync(outPath, '\ufeff' + rows.join('\n'));
    console.log(`✓ ${outPath}  (${rows.length - 1}件)`);
}

generateRackCSV();
generateMasterCSV();
console.log('\n完了。CSVをSharePointリストにインポートしてね。');
