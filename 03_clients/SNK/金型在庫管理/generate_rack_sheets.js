/**
 * 技術金型ラック配置表から、ラック毎のExcelファイルを生成するスクリプト
 * 使い方: node generate_rack_sheets.js
 *
 * 元データ: 技術金型ラック配置表_original.xlsx（ネットワークドライブからコピー済み）
 */

const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const SOURCE_FILE = path.join(__dirname, '技術金型ラック配置表_original.xlsx');
const OUTPUT_DIR = path.join(__dirname, 'racks');

// 色設定
const COLORS = {
    header: 'FF2563EB',
    headerFont: 'FFFFFFFF',
    empty: 'FFE8F5E9',
    occupied: 'FFFFF3E0',
    border: 'FFB0BEC5',
    title: 'FF1E3A5F',
    doorLabel: 'FF1E40AF',
    posLabel: 'FF6B7280',
    levelLabel: 'FF374151',
};

/**
 * 元Excelからラック構造とデータを解析する
 */
async function parseSourceExcel() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(SOURCE_FILE);

    const racks = {};

    wb.eachSheet((ws, id) => {
        const sheetName = ws.name;
        console.log(`解析中: ${sheetName} (${ws.rowCount}行 x ${ws.columnCount}列)`);

        // 型整備シートは別形式
        if (sheetName === '型整備') {
            racks['型整備'] = parseMoldMaintenance(ws);
            return;
        }

        racks[sheetName] = parseRackSheet(ws);
    });

    return racks;
}

/**
 * ラックシートの解析
 * 構造: Row6-7=扉ラベル, Row8-9=位置(左/中央/右), Row10以降=データ（4行ずつ段が変わる）
 */
function parseRackSheet(ws) {
    const rack = {
        title: '',
        doors: [],      // 扉情報（番号, 開始列, 終了列）
        positions: [],   // 位置情報（左/中央/右, 開始列, 終了列）
        levels: [],      // 段情報（名前, 開始行, 終了行）
        slots: [],       // 各スロットのデータ
    };

    // タイトル取得
    rack.title = getCellText(ws, 2, 2);

    // 扉と位置の解析（Row6, Row8から）
    const doorRow = 6;
    const posRow = 8;
    const dataStartCol = 7; // データは7列目から

    // 扉ラベル解析
    let currentDoor = '';
    let doorStart = dataStartCol;
    for (let c = dataStartCol; c <= ws.columnCount; c++) {
        const doorText = getCellText(ws, doorRow, c);
        if (doorText && doorText !== currentDoor) {
            if (currentDoor) {
                rack.doors.push({ name: currentDoor, startCol: doorStart, endCol: c - 1 });
            }
            currentDoor = doorText;
            doorStart = c;
        }
    }
    if (currentDoor) {
        rack.doors.push({ name: currentDoor, startCol: doorStart, endCol: ws.columnCount });
    }

    // 位置ラベル解析（各扉内の左/中央/右）
    let currentPos = '';
    let posStart = dataStartCol;
    let currentDoorForPos = '';
    for (let c = dataStartCol; c <= ws.columnCount; c++) {
        const posText = getCellText(ws, posRow, c);
        const doorText = getCellText(ws, doorRow, c);

        // 扉が変わったらリセット
        if (doorText && doorText !== currentDoorForPos) {
            if (currentPos) {
                rack.positions.push({ door: currentDoorForPos, pos: currentPos, startCol: posStart, endCol: c - 1 });
            }
            currentDoorForPos = doorText;
            currentPos = posText;
            posStart = c;
        } else if (posText && posText !== currentPos) {
            if (currentPos) {
                rack.positions.push({ door: currentDoorForPos, pos: currentPos, startCol: posStart, endCol: c - 1 });
            }
            currentPos = posText;
            posStart = c;
        }
    }
    if (currentPos) {
        rack.positions.push({ door: currentDoorForPos, pos: currentPos, startCol: posStart, endCol: ws.columnCount });
    }

    // 段ラベル解析（B列=Col2の値から）
    const dataStartRow = 10;
    let currentLevel = '';
    let levelStart = dataStartRow;
    for (let r = dataStartRow; r <= ws.rowCount; r++) {
        const levelText = getCellText(ws, r, 2);
        if (levelText && levelText !== currentLevel) {
            if (currentLevel) {
                rack.levels.push({ name: currentLevel, startRow: levelStart, endRow: r - 1 });
            }
            currentLevel = levelText;
            levelStart = r;
        }
    }
    if (currentLevel) {
        rack.levels.push({ name: currentLevel, startRow: levelStart, endRow: ws.rowCount });
    }

    // スロットデータ解析
    for (const pos of rack.positions) {
        for (const level of rack.levels) {
            // スロットのテキストを取得（最初のデータ行のセルから）
            const cellText = getCellText(ws, level.startRow, pos.startCol);
            const moldInfo = cellText || '空';

            rack.slots.push({
                door: pos.door,
                position: pos.pos,
                level: level.name,
                moldInfo: moldInfo,
                isEmpty: moldInfo === '空' || moldInfo === '',
            });
        }
    }

    return rack;
}

/**
 * 型整備シートの解析
 */
function parseMoldMaintenance(ws) {
    const items = [];
    for (let r = 5; r <= ws.rowCount; r++) {
        const hinban = getCellText(ws, r, 1);
        const kata = getCellText(ws, r, 2);
        const status = getCellText(ws, r, 3);
        if (hinban) {
            items.push({ hinban, kata, status });
        }
    }
    return { type: 'maintenance', items };
}

/**
 * セルの値をテキストとして取得
 */
function getCellText(ws, row, col) {
    const cell = ws.getCell(row, col);
    const v = cell.value;
    if (v === null || v === undefined) return '';
    if (typeof v === 'object' && v.richText) {
        return v.richText.map(t => t.text).join('').trim();
    }
    if (typeof v === 'object' && v.result !== undefined) {
        return String(v.result).trim();
    }
    return String(v).trim();
}

function getBorder() {
    return {
        top: { style: 'thin', color: { argb: COLORS.border } },
        left: { style: 'thin', color: { argb: COLORS.border } },
        bottom: { style: 'thin', color: { argb: COLORS.border } },
        right: { style: 'thin', color: { argb: COLORS.border } },
    };
}

/**
 * ラック毎のExcelファイルを生成
 */
async function generateRackExcel(rackName, rackData) {
    if (rackData.type === 'maintenance') {
        return generateMaintenanceExcel(rackName, rackData);
    }

    const wb = new ExcelJS.Workbook();
    wb.creator = '金型在庫管理システム';

    // === シート1: 配置図 ===
    const ws = wb.addWorksheet('配置図', { properties: { defaultColWidth: 16 } });

    // タイトル
    const title = rackData.title || rackName;
    ws.mergeCells(1, 1, 1, 10);
    const titleCell = ws.getCell(1, 1);
    titleCell.value = title;
    titleCell.font = { bold: true, size: 16, color: { argb: COLORS.title } };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(1).height = 40;

    // 更新日
    ws.mergeCells(2, 1, 2, 10);
    const dateCell = ws.getCell(2, 1);
    dateCell.value = `最終更新: ${new Date().toLocaleDateString('ja-JP')}  ※元データから自動生成`;
    dateCell.font = { size: 10, color: { argb: 'FF999999' } };
    dateCell.alignment = { horizontal: 'right' };

    // 扉ごとにセクションを作成
    let currentRow = 4;

    // ユニークな扉リストを取得
    const doorNames = [...new Set(rackData.doors.map(d => d.name))];
    const levelNames = [...new Set(rackData.levels.map(l => l.name))];
    const posNames = ['左', '中央', '右'];

    for (const doorName of doorNames) {
        // 扉ヘッダー
        ws.mergeCells(currentRow, 1, currentRow, posNames.length + 1);
        const doorCell = ws.getCell(currentRow, 1);
        doorCell.value = `  ${doorName}`;
        doorCell.font = { bold: true, size: 13, color: { argb: COLORS.doorLabel } };
        doorCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDBEAFE' } };
        doorCell.alignment = { vertical: 'middle' };
        ws.getRow(currentRow).height = 32;
        currentRow++;

        // 位置ヘッダー
        ws.getCell(currentRow, 1).value = '段';
        ws.getCell(currentRow, 1).font = { bold: true, size: 10 };
        ws.getCell(currentRow, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.header } };
        ws.getCell(currentRow, 1).font = { bold: true, color: { argb: COLORS.headerFont } };
        ws.getCell(currentRow, 1).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getCell(currentRow, 1).border = getBorder();

        for (let p = 0; p < posNames.length; p++) {
            const cell = ws.getCell(currentRow, p + 2);
            cell.value = posNames[p];
            cell.font = { bold: true, color: { argb: COLORS.headerFont } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.header } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = getBorder();
        }
        ws.getRow(currentRow).height = 28;
        currentRow++;

        // 各段のデータ
        for (const levelName of levelNames) {
            const levelCell = ws.getCell(currentRow, 1);
            levelCell.value = levelName;
            levelCell.font = { bold: true, size: 10, color: { argb: COLORS.levelLabel } };
            levelCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            levelCell.border = getBorder();

            for (let p = 0; p < posNames.length; p++) {
                const slot = rackData.slots.find(s =>
                    s.door === doorName && s.position === posNames[p] && s.level === levelName
                );
                const cell = ws.getCell(currentRow, p + 2);
                const moldText = slot ? slot.moldInfo : '';
                cell.value = moldText || '空';
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                cell.border = getBorder();

                if (!moldText || moldText === '空') {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.empty } };
                    cell.font = { color: { argb: 'FF9CA3AF' } };
                } else {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.occupied } };
                    cell.font = { size: 10 };
                }
            }
            ws.getRow(currentRow).height = 50;
            currentRow++;
        }

        currentRow++; // 扉間の空行
    }

    // 列幅
    ws.getColumn(1).width = 14;
    for (let c = 2; c <= posNames.length + 1; c++) {
        ws.getColumn(c).width = 22;
    }

    // === シート2: 一覧表（検索しやすい形式） ===
    const listWs = wb.addWorksheet('一覧表');
    listWs.columns = [
        { header: '扉', key: 'door', width: 12 },
        { header: '位置', key: 'pos', width: 8 },
        { header: '段', key: 'level', width: 14 },
        { header: '金型情報', key: 'moldInfo', width: 30 },
        { header: '状態', key: 'status', width: 8 },
    ];

    // ヘッダースタイル
    listWs.getRow(1).eachCell(cell => {
        cell.font = { bold: true, color: { argb: COLORS.headerFont } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.header } };
        cell.alignment = { horizontal: 'center' };
    });
    listWs.getRow(1).height = 28;

    // データ行
    for (const slot of rackData.slots) {
        listWs.addRow({
            door: slot.door,
            pos: slot.position,
            level: slot.level,
            moldInfo: slot.moldInfo,
            status: slot.isEmpty ? '空き' : '使用中',
        });
    }

    // オートフィルター
    listWs.autoFilter = { from: 'A1', to: 'E1' };

    // === シート3: 入出庫履歴 ===
    const histWs = wb.addWorksheet('入出庫履歴');
    histWs.columns = [
        { header: '日時', key: 'datetime', width: 20 },
        { header: '操作', key: 'action', width: 10 },
        { header: '金型No', key: 'moldNo', width: 20 },
        { header: '扉', key: 'door', width: 12 },
        { header: '位置', key: 'pos', width: 8 },
        { header: '段', key: 'level', width: 14 },
        { header: '担当者', key: 'operator', width: 12 },
        { header: '備考', key: 'note', width: 25 },
    ];
    histWs.getRow(1).eachCell(cell => {
        cell.font = { bold: true, color: { argb: COLORS.headerFont } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.header } };
        cell.alignment = { horizontal: 'center' };
    });
    histWs.getRow(1).height = 28;

    // 保存
    const safeName = rackName.replace(/[()（）]/g, '_').replace(/\s+/g, '');
    const filePath = path.join(OUTPUT_DIR, `${safeName}.xlsx`);
    await wb.xlsx.writeFile(filePath);

    // 統計
    const total = rackData.slots.length;
    const occupied = rackData.slots.filter(s => !s.isEmpty).length;
    console.log(`  ✓ ${filePath}  (${occupied}/${total}スロット使用中)`);

    return { safeName, filePath, total, occupied };
}

/**
 * 型整備シート用Excel生成
 */
async function generateMaintenanceExcel(rackName, rackData) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('型整備');
    ws.columns = [
        { header: '品番', key: 'hinban', width: 25 },
        { header: '型', key: 'kata', width: 10 },
        { header: '状態', key: 'status', width: 25 },
    ];
    ws.getRow(1).eachCell(cell => {
        cell.font = { bold: true, color: { argb: COLORS.headerFont } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.header } };
    });
    for (const item of rackData.items) {
        ws.addRow(item);
    }

    const filePath = path.join(OUTPUT_DIR, '型整備.xlsx');
    await wb.xlsx.writeFile(filePath);
    console.log(`  ✓ ${filePath}  (${rackData.items.length}件)`);
    return { safeName: '型整備', filePath, total: rackData.items.length, occupied: rackData.items.length };
}

async function main() {
    if (!fs.existsSync(OUTPUT_DIR)) {
        fs.mkdirSync(OUTPUT_DIR, { recursive: true });
    }

    console.log('=== 技術金型ラック配置表を解析中 ===\n');
    const racks = await parseSourceExcel();

    console.log('\n=== ラック別Excelを生成中 ===\n');
    const results = [];
    for (const [name, data] of Object.entries(racks)) {
        const result = await generateRackExcel(name, data);
        results.push(result);
    }

    console.log('\n=== 完了 ===');
    console.log(`生成ファイル数: ${results.length}`);
    console.log(`出力先: ${OUTPUT_DIR}\n`);

    // ラック名一覧をJSONで保存（QRコード生成用）
    const rackInfo = results.map(r => ({
        name: r.safeName,
        total: r.total,
        occupied: r.occupied,
    }));
    fs.writeFileSync(path.join(OUTPUT_DIR, 'rack_info.json'), JSON.stringify(rackInfo, null, 2));
    console.log('rack_info.json を保存しました');
}

main().catch(console.error);
