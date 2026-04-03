/**
 * 技術金型ラック配置表の既存データを更新型マスターと照合し、
 * 金型在庫管理アプリのLocalStorage初期データを生成するスクリプト
 */
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const SOURCE_FILE = path.join(__dirname, '技術金型ラック配置表_original.xlsx');
const MASTER_FILE = path.join(__dirname, 'mold_master.json');
const OUTPUT_FILE = path.join(__dirname, 'initial_inventory.json');

// ラック構造定義（HTMLアプリと同じ）
const RACK_DEFS = {
    'D(旧ラック)': { appName: '旧Dラック', doors: ['⑤番扉','④番扉','③番扉','②番扉','①番扉'], levels: ['3段目','2段目','1段目奥','1段目手前'] },
    'D(新ラック)': { appName: '新Dラック', doors: ['③番扉','②番扉','①番扉'], levels: ['3段目','2段目','1段目奥','1段目手前'] },
    'F(北側ラック)': { appName: 'F北側ラック', doors: ['⑦番扉','⑥番扉','⑤番扉'], levels: ['3段目','2段目','1段目奥','1段目手前'] },
    'F(南側ラック)': { appName: 'F南側ラック', doors: ['③番扉','②番扉','①番扉'], levels: ['3段目','2段目','1段目奥','1段目手前'] },
    'G(北側ラック)': { appName: 'G北側ラック', doors: ['④番扉','③番扉','②番扉'], levels: ['2段目','1段目奥','1段目手前'] },
};

const POSITIONS = ['左', '中央', '右'];

async function main() {
    // マスターデータ読み込み
    const masterData = JSON.parse(fs.readFileSync(MASTER_FILE, 'utf8'));

    // 品番で検索用マップ（大文字・ハイフンなし正規化）
    const byHinban = {};
    for (const m of masterData) {
        if (m.hinban) {
            const key = normalizeHinban(m.hinban);
            if (!byHinban[key]) byHinban[key] = [];
            byHinban[key].push(m);
        }
    }

    // Excel読み込み
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(SOURCE_FILE);

    const inventory = {};
    let matchCount = 0;
    let manualCount = 0;
    let emptyCount = 0;

    for (const [sheetName, rackDef] of Object.entries(RACK_DEFS)) {
        const ws = wb.getWorksheet(sheetName);
        if (!ws) { console.log(`シート「${sheetName}」が見つかりません`); continue; }

        console.log(`\n=== ${sheetName} → ${rackDef.appName} ===`);

        // Excelの構造: Row6-7=扉、Row8-9=位置(左/中央/右)、Row10以降=データ(4行ずつ段)
        const dataStartCol = 7;
        const dataStartRow = 10;

        // 扉の列範囲を解析
        const doorRanges = parseDoorRanges(ws, 6, dataStartCol);

        // 段の行範囲を解析
        const levelRanges = parseLevelRanges(ws, dataStartRow, 2);

        // 各扉の位置範囲を解析
        const posRanges = parsePosRanges(ws, 8, dataStartCol);

        // ドアインデックスとアプリの扉名をマッピング
        let doorIdx = 0;
        for (const doorRange of doorRanges) {
            if (doorIdx >= rackDef.doors.length) break;
            const doorName = rackDef.doors[doorIdx];

            // このドア内の位置を解析
            const positionsInDoor = posRanges.filter(p =>
                p.startCol >= doorRange.startCol && p.endCol <= doorRange.endCol
            );

            let posIdx = 0;
            for (const posRange of positionsInDoor) {
                if (posIdx >= POSITIONS.length) break;
                const posName = POSITIONS[posIdx];

                let levelIdx = 0;
                for (const levelRange of levelRanges) {
                    if (levelIdx >= rackDef.levels.length) break;
                    const levelName = rackDef.levels[levelIdx];

                    // セルの値を取得（最初のデータ行、最初の列）
                    const cellText = getCellText(ws, levelRange.startRow, posRange.startCol);

                    if (cellText && cellText !== '空' && cellText !== '') {
                        // 品番部分を抽出（改行の前の部分）
                        const lines = cellText.split('\n').map(l => l.trim()).filter(l => l);
                        const hinbanPart = lines[0] || '';
                        const descPart = lines.slice(1).join(' ');

                        // マスターと照合
                        const match = findMatch(hinbanPart, byHinban);

                        const key = `${rackDef.appName}/${doorName}/${posName}/${levelName}`;

                        if (match) {
                            inventory[key] = {
                                moldNo: match.no,
                                hinban: match.hinban,
                                hinmei: match.hinmei,
                                koushinNo: match.koushinNo,
                                customer: match.customer,
                                updatedAt: new Date().toISOString(),
                            };
                            matchCount++;
                        } else {
                            // マスターにない → 手動入力扱い
                            inventory[key] = {
                                moldNo: 'M-' + Date.now() + '-' + manualCount,
                                hinban: hinbanPart,
                                hinmei: descPart || hinbanPart,
                                koushinNo: '',
                                customer: '手動登録（初期データ）',
                                updatedAt: new Date().toISOString(),
                            };
                            manualCount++;
                        }
                    } else {
                        emptyCount++;
                    }

                    levelIdx++;
                }
                posIdx++;
            }
            doorIdx++;
        }
    }

    // 結果保存
    fs.writeFileSync(OUTPUT_FILE, JSON.stringify(inventory, null, 2));

    console.log('\n=== 結果 ===');
    console.log(`マスター照合成功: ${matchCount}件`);
    console.log(`手動登録（マスターに該当なし）: ${manualCount}件`);
    console.log(`空きスロット: ${emptyCount}件`);
    console.log(`合計登録: ${matchCount + manualCount}件`);
    console.log(`\n保存先: ${OUTPUT_FILE}`);
}

function normalizeHinban(s) {
    return s.toUpperCase().replace(/[\s\-\.]/g, '');
}

function findMatch(hinbanPart, byHinban) {
    if (!hinbanPart) return null;
    const key = normalizeHinban(hinbanPart);
    if (!key) return null;

    // 完全一致
    if (byHinban[key]) return byHinban[key][0];

    // 前方一致（品番の先頭部分で一致）
    for (const [mKey, vals] of Object.entries(byHinban)) {
        if (mKey.startsWith(key) || key.startsWith(mKey)) {
            return vals[0];
        }
    }

    return null;
}

function getCellText(ws, row, col) {
    const cell = ws.getCell(row, col);
    const v = cell.value;
    if (v === null || v === undefined) return '';
    if (typeof v === 'object' && v.richText) return v.richText.map(t => t.text).join('').trim();
    if (typeof v === 'object' && v.result !== undefined) return String(v.result).trim();
    return String(v).trim();
}

function parseDoorRanges(ws, row, startCol) {
    const ranges = [];
    let current = '';
    let start = startCol;
    for (let c = startCol; c <= ws.columnCount; c++) {
        const text = getCellText(ws, row, c);
        if (text && text !== current) {
            if (current) ranges.push({ name: current, startCol: start, endCol: c - 1 });
            current = text;
            start = c;
        }
    }
    if (current) ranges.push({ name: current, startCol: start, endCol: ws.columnCount });
    return ranges;
}

function parsePosRanges(ws, row, startCol) {
    const ranges = [];
    let current = '';
    let start = startCol;
    for (let c = startCol; c <= ws.columnCount; c++) {
        const text = getCellText(ws, row, c);
        if (text && text !== current) {
            if (current) ranges.push({ name: current, startCol: start, endCol: c - 1 });
            current = text;
            start = c;
        }
    }
    if (current) ranges.push({ name: current, startCol: start, endCol: ws.columnCount });
    return ranges;
}

function parseLevelRanges(ws, startRow, labelCol) {
    const ranges = [];
    let current = '';
    let start = startRow;
    for (let r = startRow; r <= ws.rowCount; r++) {
        const text = getCellText(ws, r, labelCol);
        if (text && text !== current) {
            if (current) ranges.push({ name: current, startRow: start, endRow: r - 1 });
            current = text;
            start = r;
        }
    }
    if (current) ranges.push({ name: current, startRow: start, endRow: ws.rowCount });
    return ranges;
}

main().catch(console.error);
