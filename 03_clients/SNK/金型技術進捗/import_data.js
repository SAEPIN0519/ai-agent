// 金型技術 製作進捗管理Excel生成（ガントチャート付き）
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const path = require('path');

// Excelシリアル日付 → Date
function serialToDate(serial) {
  if (!serial) return null;
  if (typeof serial === 'string') {
    const d = new Date(serial);
    return isNaN(d.getTime()) ? null : d;
  }
  if (typeof serial !== 'number' || serial < 100) return null;
  const epoch = new Date(1899, 11, 30);
  return new Date(epoch.getTime() + serial * 86400000);
}

function guessStatus(row) {
  if (row.shippedDate) return '出荷済';
  if (row.completionDate) return '制作完了';
  return '未着手';
}

async function main() {
  // ===== 元データ読み込み =====
  const srcWb = XLSX.readFile('C:/Users/matsuoka/Desktop/コピー57期 製作予定.xlsx');
  const allRows = [];

  srcWb.SheetNames.forEach(sheetName => {
    const ws = srcWb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (data.length < 2) return;

    const header = data[0].map(h => String(h).trim());
    const colMap = {};
    header.forEach((h, idx) => {
      if (h === '受注番号') colMap.order = idx;
      if (h === '製品名') colMap.product = idx;
      if (h === '製作アイテム') colMap.item = idx;
      if (h === '種類') colMap.type = idx;
      if (h === '見積り金額') colMap.estimate = idx;
      if (h === '完成予定') colMap.completionDate = idx;
      if (h === '備考') colMap.note = idx;
      if (h === '出荷日' || h === '10月' || h === '11月') {
        if (!colMap.shippedDate) colMap.shippedDate = idx;
      }
    });

    data.slice(1).forEach(row => {
      const order = row[colMap.order];
      const product = row[colMap.product];
      if (!order && !product) return;
      if (!product && !row[colMap.item]) return;

      allRows.push({
        order: order || '',
        product: product || '',
        item: row[colMap.item] || '',
        type: row[colMap.type] || '',
        estimate: row[colMap.estimate] || '',
        completionDate: colMap.completionDate !== undefined ? row[colMap.completionDate] : '',
        note: row[colMap.note] || '',
        shippedDate: colMap.shippedDate !== undefined ? row[colMap.shippedDate] : '',
        sheet: sheetName,
      });
    });
  });

  console.log(`読み込みデータ: ${allRows.length}件`);

  // ===== Excel生成 =====
  const wb = new ExcelJS.Workbook();
  wb.creator = 'SNK金型技術';
  wb.created = new Date();

  const C = {
    headerBg: '1E3A5F',
    headerFont: 'FFFFFF',
    designHeader: '6D28D9',
    makeHeader: '0E7490',
    statusHeader: 'B45309',
    lightGray: 'F3F4F6',
    border: 'D1D5DB',
    designBar: 'C4B5FD',   // 設計バー（薄紫）
    makeBar: '67E8F9',      // 制作バー（薄シアン）
  };

  const thinBorder = {
    top: { style: 'thin', color: { argb: C.border } },
    bottom: { style: 'thin', color: { argb: C.border } },
    left: { style: 'thin', color: { argb: C.border } },
    right: { style: 'thin', color: { argb: C.border } },
  };

  function styleHeader(row, colCount) {
    row.height = 28;
    row.font = { bold: true, color: { argb: C.headerFont }, size: 11 };
    row.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    for (let i = 1; i <= colCount; i++) {
      row.getCell(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.headerBg } };
      row.getCell(i).border = thinBorder;
    }
  }

  function addDropdown(ws, cellRef, list) {
    ws.getCell(cellRef).dataValidation = {
      type: 'list', allowBlank: true,
      formulae: ['"' + list.join(',') + '"'],
      showErrorMessage: true, errorTitle: '入力エラー', error: 'リストから選択してください',
    };
  }

  const typeList = ['Die', 'TP', 'Tem', 'Mill', 'GC', 'Ins', 'S', 'Other'];
  const designStatusList = ['未着手', '進行中', '完了'];
  const makeStatusList = ['未着手', '進行中', '完了'];
  const overallStatusList = ['未着手', '設計中', '設計完了', '制作中', '制作完了', '出荷済'];

  const EXTRA_ROWS = 50;

  // =============================================================
  //  シート1: 進捗管理（設計開始日・制作開始日を追加）
  // =============================================================
  const ws1 = wb.addWorksheet('進捗管理', {
    views: [{ state: 'frozen', ySplit: 1, xSplit: 3 }],
  });

  // 列構成（設計開始日・制作開始日を追加）
  // A:受注番号 B:製品名 C:製作アイテム D:種類
  // E:設計開始日 F:設計期限 G:設計ステータス
  // H:制作開始日 I:制作期限 J:制作ステータス
  // K:全体ステータス L:依頼元 M:見積り金額 N:登録月 O:備考
  ws1.columns = [
    { header: '受注番号',       key: 'order',      width: 12 },
    { header: '製品名',         key: 'product',    width: 26 },
    { header: '製作アイテム',   key: 'item',       width: 26 },
    { header: '種類',           key: 'type',       width: 10 },
    { header: '設計開始日',     key: 'designStart', width: 14 },
    { header: '設計期限',       key: 'designDL',   width: 14 },
    { header: '設計ステータス', key: 'designSt',   width: 14 },
    { header: '制作開始日',     key: 'makeStart',  width: 14 },
    { header: '制作期限',       key: 'makeDL',     width: 14 },
    { header: '制作ステータス', key: 'makeSt',     width: 14 },
    { header: '全体ステータス', key: 'overallSt',  width: 14 },
    { header: '依頼元',         key: 'requester',  width: 12 },
    { header: '見積り金額',     key: 'estimate',   width: 14 },
    { header: '登録月',         key: 'regMonth',   width: 10 },
    { header: '備考',           key: 'note',       width: 34 },
  ];

  styleHeader(ws1.getRow(1), 15);
  // 設計列ヘッダー色（E,F,G）
  for (let c = 5; c <= 7; c++) {
    ws1.getRow(1).getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.designHeader } };
  }
  // 制作列ヘッダー色（H,I,J）
  for (let c = 8; c <= 10; c++) {
    ws1.getRow(1).getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.makeHeader } };
  }
  // 全体ステータス（K）
  ws1.getRow(1).getCell(11).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.statusHeader } };

  // データ挿入
  allRows.sort((a, b) => String(a.product).localeCompare(String(b.product)));

  allRows.forEach((item, idx) => {
    const r = idx + 2;
    const row = ws1.getRow(r);
    const completionDate = serialToDate(item.completionDate);
    const shippedDate = serialToDate(item.shippedDate);
    const status = guessStatus(item);

    let designSt = '未着手', makeSt = '未着手';
    if (status === '出荷済' || status === '制作完了') { designSt = '完了'; makeSt = '完了'; }

    row.getCell(1).value = item.order;
    row.getCell(2).value = item.product;
    row.getCell(3).value = item.item;
    row.getCell(4).value = item.type;
    row.getCell(5).value = null;              // 設計開始日（手動入力）
    row.getCell(6).value = completionDate;    // 設計期限
    row.getCell(7).value = designSt;
    row.getCell(8).value = null;              // 制作開始日（手動入力）
    row.getCell(9).value = completionDate;    // 制作期限
    row.getCell(10).value = makeSt;
    row.getCell(11).value = status;
    row.getCell(12).value = '';
    row.getCell(13).value = item.estimate || null;
    row.getCell(14).value = item.sheet;
    row.getCell(15).value = item.note;

    row.height = 22;
    row.alignment = { vertical: 'middle' };
  });

  const lastDataRow = allRows.length + 1 + EXTRA_ROWS;

  // 全行スタイル＆ドロップダウン
  for (let r = 2; r <= lastDataRow; r++) {
    const row = ws1.getRow(r);
    for (let c = 1; c <= 15; c++) {
      row.getCell(c).border = thinBorder;
    }
    if (r % 2 === 0) {
      for (let c = 1; c <= 15; c++) {
        row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lightGray } };
      }
    }
    addDropdown(ws1, `D${r}`, typeList);
    addDropdown(ws1, `G${r}`, designStatusList);
    addDropdown(ws1, `J${r}`, makeStatusList);
    addDropdown(ws1, `K${r}`, overallStatusList);

    ws1.getCell(`E${r}`).numFmt = 'yyyy/mm/dd';
    ws1.getCell(`F${r}`).numFmt = 'yyyy/mm/dd';
    ws1.getCell(`H${r}`).numFmt = 'yyyy/mm/dd';
    ws1.getCell(`I${r}`).numFmt = 'yyyy/mm/dd';
    ws1.getCell(`M${r}`).numFmt = '#,##0';
  }

  // 条件付き書式: 全体ステータス色分け（K列）
  ws1.addConditionalFormatting({
    ref: `K2:K${lastDataRow}`,
    rules: [
      { type: 'cellIs', operator: 'equal', formulae: ['"未着手"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'F3F4F6' } }, font: { color: { argb: '6B7280' } } }, priority: 1 },
      { type: 'cellIs', operator: 'equal', formulae: ['"設計中"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'EDE9FE' } }, font: { color: { argb: '6D28D9' } } }, priority: 2 },
      { type: 'cellIs', operator: 'equal', formulae: ['"設計完了"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'DBEAFE' } }, font: { color: { argb: '1D4ED8' } } }, priority: 3 },
      { type: 'cellIs', operator: 'equal', formulae: ['"制作中"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'CFFAFE' } }, font: { color: { argb: '0E7490' } } }, priority: 4 },
      { type: 'cellIs', operator: 'equal', formulae: ['"制作完了"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'DCFCE7' } }, font: { color: { argb: '15803D' } } }, priority: 5 },
      { type: 'cellIs', operator: 'equal', formulae: ['"出荷済"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'E5E7EB' } }, font: { color: { argb: '9CA3AF' } } }, priority: 6 },
    ],
  });
  // 設計ステータス（G列）
  ws1.addConditionalFormatting({
    ref: `G2:G${lastDataRow}`,
    rules: [
      { type: 'cellIs', operator: 'equal', formulae: ['"進行中"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'EDE9FE' } }, font: { color: { argb: '6D28D9' } } }, priority: 7 },
      { type: 'cellIs', operator: 'equal', formulae: ['"完了"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'DCFCE7' } }, font: { color: { argb: '15803D' } } }, priority: 8 },
    ],
  });
  // 制作ステータス（J列）
  ws1.addConditionalFormatting({
    ref: `J2:J${lastDataRow}`,
    rules: [
      { type: 'cellIs', operator: 'equal', formulae: ['"進行中"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'CFFAFE' } }, font: { color: { argb: '0E7490' } } }, priority: 9 },
      { type: 'cellIs', operator: 'equal', formulae: ['"完了"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'DCFCE7' } }, font: { color: { argb: '15803D' } } }, priority: 10 },
    ],
  });
  // 期限超過アラート
  ws1.addConditionalFormatting({
    ref: `F2:F${lastDataRow}`,
    rules: [
      { type: 'expression', formulae: ['AND(F2<>"",F2<TODAY(),G2<>"完了")'], style: { font: { color: { argb: 'DC2626' }, bold: true } }, priority: 11 },
      { type: 'expression', formulae: ['AND(F2<>"",F2-TODAY()<=3,F2-TODAY()>=0,G2<>"完了")'], style: { font: { color: { argb: 'D97706' }, bold: true } }, priority: 12 },
    ],
  });
  ws1.addConditionalFormatting({
    ref: `I2:I${lastDataRow}`,
    rules: [
      { type: 'expression', formulae: ['AND(I2<>"",I2<TODAY(),J2<>"完了")'], style: { font: { color: { argb: 'DC2626' }, bold: true } }, priority: 13 },
      { type: 'expression', formulae: ['AND(I2<>"",I2-TODAY()<=3,I2-TODAY()>=0,J2<>"完了")'], style: { font: { color: { argb: 'D97706' }, bold: true } }, priority: 14 },
    ],
  });

  ws1.autoFilter = 'A1:O1';
  ws1.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1 };


  // =============================================================
  //  シート2: ガントチャート
  // =============================================================
  const wsGantt = wb.addWorksheet('ガントチャート', {
    views: [{ state: 'frozen', ySplit: 2, xSplit: 5 }],
  });

  // ガントチャート期間: 2026/3/1 〜 2027/2/28（1年間）
  const ganttStart = new Date(2026, 2, 1);  // 2026/3/1
  const ganttEnd = new Date(2027, 1, 28);   // 2027/2/28

  // 日付列を生成
  const dates = [];
  const d = new Date(ganttStart);
  while (d <= ganttEnd) {
    dates.push(new Date(d));
    d.setDate(d.getDate() + 1);
  }

  // 左側の固定列
  // A:受注番号 B:製品名 C:アイテム D:全体ステータス E:種類
  const fixedCols = 5;
  const ganttColStart = fixedCols + 1; // F列から日付

  // ヘッダー行1: 月名（結合）
  const headerRow1 = wsGantt.getRow(1);
  headerRow1.height = 20;

  // 固定列ヘッダー（2行目に書く）
  const headerRow2 = wsGantt.getRow(2);
  headerRow2.height = 22;

  const fixedHeaders = ['受注番号', '製品名', '製作アイテム', 'ステータス', '種類'];
  const fixedWidths = [10, 22, 22, 12, 8];

  for (let i = 0; i < fixedHeaders.length; i++) {
    const col = i + 1;
    wsGantt.getColumn(col).width = fixedWidths[i];

    // 行1と行2を結合
    wsGantt.mergeCells(1, col, 2, col);
    const cell = headerRow1.getCell(col);
    cell.value = fixedHeaders[i];
    cell.font = { bold: true, color: { argb: C.headerFont }, size: 10 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.headerBg } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = thinBorder;
  }

  // 日付列ヘッダー
  const monthNames = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
  let currentMonth = -1;
  let monthStartCol = -1;

  // 月ごとの背景色（交互）
  const monthColors = ['E8F0FE', 'F3F4F6'];

  dates.forEach((date, idx) => {
    const col = ganttColStart + idx;
    const month = date.getMonth();

    // 列幅を細く（ガントバー用）
    wsGantt.getColumn(col).width = 2.5;

    // 行2に日付を入れる
    const cell2 = headerRow2.getCell(col);
    cell2.value = date;
    cell2.numFmt = 'd';
    cell2.font = { size: 7, color: { argb: '6B7280' } };
    cell2.alignment = { horizontal: 'center', vertical: 'middle' };
    cell2.border = thinBorder;

    // 日曜日は背景を薄くする
    if (date.getDay() === 0) {
      cell2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FEE2E2' } };
    } else if (date.getDay() === 6) {
      cell2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'DBEAFE' } };
    } else {
      cell2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lightGray } };
    }

    // 月が変わったら行1に月名を書く（後で結合）
    if (month !== currentMonth) {
      if (currentMonth !== -1 && monthStartCol !== -1) {
        // 前の月を結合
        if (col - 1 > monthStartCol) {
          wsGantt.mergeCells(1, monthStartCol, 1, col - 1);
        }
      }
      currentMonth = month;
      monthStartCol = col;

      const cell1 = headerRow1.getCell(col);
      cell1.value = monthNames[month];
      cell1.font = { bold: true, size: 10, color: { argb: C.headerFont } };
      cell1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.headerBg } };
      cell1.alignment = { horizontal: 'center', vertical: 'middle' };
      cell1.border = thinBorder;
    }
  });
  // 最後の月を結合
  if (monthStartCol !== -1) {
    const lastCol = ganttColStart + dates.length - 1;
    if (lastCol > monthStartCol) {
      wsGantt.mergeCells(1, monthStartCol, 1, lastCol);
    }
  }

  // データ行（進捗管理シートから数式参照）
  const ganttDataRows = allRows.length + EXTRA_ROWS;

  for (let r = 3; r <= ganttDataRows + 2; r++) {
    const row = wsGantt.getRow(r);
    row.height = 18;

    const dataRowNum = r - 1; // 進捗管理シートの行番号（ヘッダー=1, データ=2〜）

    // A: 受注番号 = 進捗管理!A{dataRowNum}
    row.getCell(1).value = { formula: `IF('進捗管理'!A${dataRowNum}="","",'進捗管理'!A${dataRowNum})` };
    row.getCell(1).font = { size: 10 };
    row.getCell(1).border = thinBorder;

    // B: 製品名
    row.getCell(2).value = { formula: `IF('進捗管理'!B${dataRowNum}="","",'進捗管理'!B${dataRowNum})` };
    row.getCell(2).font = { size: 10 };
    row.getCell(2).border = thinBorder;

    // C: 製作アイテム
    row.getCell(3).value = { formula: `IF('進捗管理'!C${dataRowNum}="","",'進捗管理'!C${dataRowNum})` };
    row.getCell(3).font = { size: 10 };
    row.getCell(3).border = thinBorder;

    // D: 全体ステータス
    row.getCell(4).value = { formula: `IF('進捗管理'!K${dataRowNum}="","",'進捗管理'!K${dataRowNum})` };
    row.getCell(4).font = { size: 9 };
    row.getCell(4).border = thinBorder;

    // E: 種類
    row.getCell(5).value = { formula: `IF('進捗管理'!D${dataRowNum}="","",'進捗管理'!D${dataRowNum})` };
    row.getCell(5).font = { size: 9 };
    row.getCell(5).border = thinBorder;

    // 日付セルの罫線（薄く）
    for (let c = ganttColStart; c < ganttColStart + dates.length; c++) {
      row.getCell(c).border = {
        left: { style: 'hair', color: { argb: 'E5E7EB' } },
        right: { style: 'hair', color: { argb: 'E5E7EB' } },
      };
    }
  }

  // ===== ガントチャートの条件付き書式（自動バー表示） =====
  // 日付列の範囲を計算
  const firstDateCol = String.fromCharCode(64 + ganttColStart); // F
  // 列番号→列文字変換（26超え対応）
  function colToLetter(c) {
    let s = '';
    while (c > 0) {
      c--;
      s = String.fromCharCode(65 + (c % 26)) + s;
      c = Math.floor(c / 26);
    }
    return s;
  }
  const lastDateCol = colToLetter(ganttColStart + dates.length - 1);
  const ganttRef = `${firstDateCol}3:${lastDateCol}${ganttDataRows + 2}`;

  console.log(`ガントチャート範囲: ${ganttRef}`);
  console.log(`日付列数: ${dates.length}日`);

  // 進捗管理シートの列: E=設計開始日, F=設計期限, H=制作開始日, I=制作期限
  // ガントチャートのデータ行rに対応する進捗管理の行は r-1
  // ガントの日付はヘッダー行2にある

  // 設計バー（紫）: 設計開始日〜設計期限
  wsGantt.addConditionalFormatting({
    ref: ganttRef,
    rules: [
      {
        type: 'expression',
        formulae: [`AND('進捗管理'!$E${2}<>"",'進捗管理'!$F${2}<>"",${firstDateCol}$2>='進捗管理'!$E${2},${firstDateCol}$2<='進捗管理'!$F${2})`],
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.designBar } },
        },
        priority: 20,
      },
    ],
  });

  // 制作バー（水色）: 制作開始日〜制作期限（設計バーを上書き）
  wsGantt.addConditionalFormatting({
    ref: ganttRef,
    rules: [
      {
        type: 'expression',
        formulae: [`AND('進捗管理'!$H${2}<>"",'進捗管理'!$I${2}<>"",${firstDateCol}$2>='進捗管理'!$H${2},${firstDateCol}$2<='進捗管理'!$I${2})`],
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.makeBar } },
        },
        priority: 19,
      },
    ],
  });

  // 今日の線（赤い縦線の代わりに、今日の列を薄赤にする）
  wsGantt.addConditionalFormatting({
    ref: ganttRef,
    rules: [
      {
        type: 'expression',
        formulae: [`${firstDateCol}$2=TODAY()`],
        style: {
          border: {
            left: { style: 'medium', color: { argb: 'DC2626' } },
            right: { style: 'medium', color: { argb: 'DC2626' } },
          },
        },
        priority: 18,
      },
    ],
  });

  // ステータス色分け（D列）
  wsGantt.addConditionalFormatting({
    ref: `D3:D${ganttDataRows + 2}`,
    rules: [
      { type: 'cellIs', operator: 'equal', formulae: ['"出荷済"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'E5E7EB' } }, font: { color: { argb: '9CA3AF' }, size: 9 } }, priority: 30 },
      { type: 'cellIs', operator: 'equal', formulae: ['"制作完了"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'DCFCE7' } }, font: { color: { argb: '15803D' }, size: 9 } }, priority: 31 },
      { type: 'cellIs', operator: 'equal', formulae: ['"制作中"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'CFFAFE' } }, font: { color: { argb: '0E7490' }, size: 9 } }, priority: 32 },
      { type: 'cellIs', operator: 'equal', formulae: ['"設計中"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'EDE9FE' } }, font: { color: { argb: '6D28D9' }, size: 9 } }, priority: 33 },
    ],
  });

  // 凡例行を右上あたりに（ガントチャートシートの上部に説明追加）
  // → 使い方シートに記載する

  wsGantt.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: 5 } };
  wsGantt.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1 };


  // =============================================================
  //  シート3: 製品別サマリー
  // =============================================================
  const ws2 = wb.addWorksheet('製品別サマリー', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });

  ws2.columns = [
    { header: '製品名',       width: 26 },
    { header: 'アイテム数',   width: 12 },
    { header: '設計完了',     width: 12 },
    { header: '制作完了',     width: 12 },
    { header: '出荷済',       width: 10 },
    { header: '進捗率',       width: 10 },
    { header: '最終制作期限', width: 16 },
    { header: '備考',         width: 30 },
  ];
  styleHeader(ws2.getRow(1), 8);

  const uniqueProducts = [...new Set(allRows.map(r => r.product))].sort();

  uniqueProducts.forEach((prod, idx) => {
    const r = idx + 2;
    ws2.getCell(`A${r}`).value = prod;
    ws2.getCell(`B${r}`).value = { formula: `COUNTIF('進捗管理'!B:B,A${r})` };
    ws2.getCell(`C${r}`).value = { formula: `COUNTIFS('進捗管理'!B:B,A${r},'進捗管理'!G:G,"完了")` };
    ws2.getCell(`D${r}`).value = { formula: `COUNTIFS('進捗管理'!B:B,A${r},'進捗管理'!J:J,"完了")` };
    ws2.getCell(`E${r}`).value = { formula: `COUNTIFS('進捗管理'!B:B,A${r},'進捗管理'!K:K,"出荷済")` };
    ws2.getCell(`F${r}`).value = { formula: `IF(B${r}=0,"",D${r}/B${r})` };
    ws2.getCell(`F${r}`).numFmt = '0%';
    ws2.getCell(`G${r}`).value = { formula: `IFERROR(MAXIFS('進捗管理'!I:I,'進捗管理'!B:B,A${r}),"")` };
    ws2.getCell(`G${r}`).numFmt = 'yyyy/mm/dd';

    const row = ws2.getRow(r);
    row.height = 22;
    for (let c = 1; c <= 8; c++) row.getCell(c).border = thinBorder;
    if (r % 2 === 0) {
      for (let c = 1; c <= 8; c++) {
        row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lightGray } };
      }
    }
  });

  ws2.addConditionalFormatting({
    ref: `F2:F${uniqueProducts.length + 1}`,
    rules: [
      { type: 'cellIs', operator: 'equal', formulae: ['1'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'DCFCE7' } }, font: { color: { argb: '15803D' }, bold: true } }, priority: 15 },
      { type: 'cellIs', operator: 'greaterThanOrEqual', formulae: ['0.5'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'DBEAFE' } }, font: { color: { argb: '1D4ED8' } } }, priority: 16 },
    ],
  });
  ws2.autoFilter = 'A1:H1';


  // =============================================================
  //  シート4: 使い方
  // =============================================================
  const ws3 = wb.addWorksheet('使い方');
  const guide = [
    ['金型技術 製作進捗管理 — 使い方ガイド'],
    [''],
    ['■ 進捗管理シート（メイン）'],
    ['  1. 新規依頼 → 受注番号・製品名・製作アイテム・種類を入力'],
    ['  2. 設計開始日・設計期限、制作開始日・制作期限を入力'],
    ['  3. 開始日と期限を入力すると、ガントチャートに自動でバーが表示されます'],
    ['  4. ステータスはドロップダウンから選択（色が自動で変わります）'],
    [''],
    ['■ ガントチャートシート'],
    ['  - 進捗管理シートの開始日・期限から自動でバーが表示されます'],
    ['  - 紫のバー = 設計期間（設計開始日〜設計期限）'],
    ['  - 水色のバー = 制作期間（制作開始日〜制作期限）'],
    ['  - 今日の日付に赤い縦線が表示されます'],
    ['  - 左5列は固定。右にスクロールしても製品名が見えます'],
    ['  - ヘッダーのフィルターでステータスや種類で絞り込めます'],
    [''],
    ['■ 新規登録 → ガントチャート自動表示の流れ'],
    ['  ① 進捗管理シートに新しい行を追加'],
    ['  ② 設計開始日・設計期限を入力 → ガントチャートに紫バーが出る'],
    ['  ③ 制作開始日・制作期限を入力 → ガントチャートに水色バーが出る'],
    ['  ※ ガントチャートシートは触らなくてOK。進捗管理だけ更新すれば自動反映'],
    [''],
    ['■ 製品別サマリーシート'],
    ['  - 進捗管理シートのデータから自動集計（COUNTIFS関数）'],
    [''],
    ['■ 種類の略称'],
    ['  Die=金型 | TP=トリム型 | Tem=ベント落とし | Mill=加工治具'],
    ['  GC=切断治具 | Ins=検査治具 | S=強度試験 | Other=その他'],
  ];
  guide.forEach((row, i) => {
    ws3.getCell(`A${i+1}`).value = row[0];
    if (i === 0) ws3.getCell('A1').font = { size: 14, bold: true, color: { argb: C.headerBg } };
  });
  ws3.getColumn(1).width = 80;


  // ===== 保存 =====
  const outPath = path.join(__dirname, '金型技術_製作進捗管理.xlsx');
  await wb.xlsx.writeFile(outPath);
  console.log(`\n生成完了: ${outPath}`);
  console.log(`データ: ${allRows.length}件 / 製品: ${uniqueProducts.length}件`);
  console.log(`ガントチャート: ${dates.length}日分（2026/3/1〜2027/2/28）`);
}

main().catch(err => { console.error('エラー:', err); process.exit(1); });
