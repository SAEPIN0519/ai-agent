// 金型技術 製作進捗管理Excel生成スクリプト
const ExcelJS = require('exceljs');
const path = require('path');

async function main() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'SNK金型技術';
  wb.created = new Date();

  // 色定義
  const C = {
    headerBg: '1E3A5F',
    headerFont: 'FFFFFF',
    designBg: 'EDE9FE',    // 設計系（紫系）
    makeBg: 'CFFAFE',      // 制作系（シアン系）
    statusBg: 'FEF3C7',    // ステータス（黄系）
    lightGray: 'F3F4F6',
    border: 'D1D5DB',
    dangerFont: 'DC2626',
    warningFont: 'D97706',
  };

  // 共通スタイル関数
  function styleHeader(row, colCount) {
    row.height = 28;
    row.font = { bold: true, color: { argb: C.headerFont }, size: 11 };
    row.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    for (let i = 1; i <= colCount; i++) {
      row.getCell(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.headerBg } };
      row.getCell(i).border = {
        top: { style: 'thin', color: { argb: C.border } },
        bottom: { style: 'thin', color: { argb: C.border } },
        left: { style: 'thin', color: { argb: C.border } },
        right: { style: 'thin', color: { argb: C.border } },
      };
    }
  }

  function styleDataRows(ws, startRow, endRow, colCount) {
    for (let r = startRow; r <= endRow; r++) {
      const row = ws.getRow(r);
      row.height = 22;
      row.alignment = { vertical: 'middle' };
      for (let c = 1; c <= colCount; c++) {
        row.getCell(c).border = {
          top: { style: 'thin', color: { argb: C.border } },
          bottom: { style: 'thin', color: { argb: C.border } },
          left: { style: 'thin', color: { argb: C.border } },
          right: { style: 'thin', color: { argb: C.border } },
        };
        // 偶数行に薄い背景
        if (r % 2 === 0) {
          row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lightGray } };
        }
      }
    }
  }

  // ドロップダウン設定
  function addDropdown(ws, cellRef, list) {
    ws.getCell(cellRef).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"' + list.join(',') + '"'],
      showErrorMessage: true,
      errorTitle: '入力エラー',
      error: 'リストから選択してください',
    };
  }

  // 種類リスト
  const typeList = ['Die', 'TP', 'Tem', 'Mill', 'GC', 'Ins', 'S', 'Other'];
  const designStatusList = ['未着手', '進行中', '完了'];
  const makeStatusList = ['未着手', '進行中', '完了'];
  const overallStatusList = ['未着手', '設計中', '設計完了', '制作中', '制作完了', '出荷済'];

  const DATA_ROWS = 100; // データ入力行数

  // =============================================================
  //  シート1: 進捗管理（メインシート）
  // =============================================================
  const ws1 = wb.addWorksheet('進捗管理', {
    views: [{ state: 'frozen', ySplit: 1 }], // ヘッダー行固定
  });

  ws1.columns = [
    { header: '受注番号',     key: 'order',        width: 12 },
    { header: '製品名',       key: 'product',      width: 26 },
    { header: '製作アイテム', key: 'item',         width: 26 },
    { header: '種類',         key: 'type',         width: 10 },
    { header: '設計期限',     key: 'designDL',     width: 14 },
    { header: '設計ステータス', key: 'designSt',   width: 16 },
    { header: '制作期限',     key: 'makeDL',       width: 14 },
    { header: '制作ステータス', key: 'makeSt',     width: 16 },
    { header: '全体ステータス', key: 'overallSt',  width: 16 },
    { header: '依頼元',       key: 'requester',    width: 12 },
    { header: '見積り金額',   key: 'estimate',     width: 14 },
    { header: '登録日',       key: 'regDate',      width: 14 },
    { header: '備考',         key: 'note',         width: 34 },
  ];

  // ヘッダースタイル
  styleHeader(ws1.getRow(1), 13);

  // 設計列の背景色（E,F列ヘッダー）
  ws1.getRow(1).getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '6D28D9' } };
  ws1.getRow(1).getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '6D28D9' } };
  // 制作列の背景色（G,H列ヘッダー）
  ws1.getRow(1).getCell(7).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0E7490' } };
  ws1.getRow(1).getCell(8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0E7490' } };
  // 全体ステータス
  ws1.getRow(1).getCell(9).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'B45309' } };

  // データ行スタイル
  styleDataRows(ws1, 2, DATA_ROWS + 1, 13);

  // ドロップダウン設定（各行に）
  for (let r = 2; r <= DATA_ROWS + 1; r++) {
    addDropdown(ws1, `D${r}`, typeList);
    addDropdown(ws1, `F${r}`, designStatusList);
    addDropdown(ws1, `H${r}`, makeStatusList);
    addDropdown(ws1, `I${r}`, overallStatusList);

    // 日付列の書式
    ws1.getCell(`E${r}`).numFmt = 'yyyy/mm/dd';
    ws1.getCell(`G${r}`).numFmt = 'yyyy/mm/dd';
    ws1.getCell(`L${r}`).numFmt = 'yyyy/mm/dd';

    // 金額列の書式
    ws1.getCell(`K${r}`).numFmt = '#,##0';
  }

  // 条件付き書式：全体ステータス列の色分け
  ws1.addConditionalFormatting({
    ref: `I2:I${DATA_ROWS + 1}`,
    rules: [
      { type: 'cellIs', operator: 'equal', formulae: ['"未着手"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'F3F4F6' } }, font: { color: { argb: '6B7280' } } }, priority: 1 },
      { type: 'cellIs', operator: 'equal', formulae: ['"設計中"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'EDE9FE' } }, font: { color: { argb: '6D28D9' } } }, priority: 2 },
      { type: 'cellIs', operator: 'equal', formulae: ['"設計完了"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'DBEAFE' } }, font: { color: { argb: '1D4ED8' } } }, priority: 3 },
      { type: 'cellIs', operator: 'equal', formulae: ['"制作中"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'CFFAFE' } }, font: { color: { argb: '0E7490' } } }, priority: 4 },
      { type: 'cellIs', operator: 'equal', formulae: ['"制作完了"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'DCFCE7' } }, font: { color: { argb: '15803D' } } }, priority: 5 },
      { type: 'cellIs', operator: 'equal', formulae: ['"出荷済"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'E5E7EB' } }, font: { color: { argb: '9CA3AF' } } }, priority: 6 },
    ],
  });

  // 条件付き書式：設計ステータス列
  ws1.addConditionalFormatting({
    ref: `F2:F${DATA_ROWS + 1}`,
    rules: [
      { type: 'cellIs', operator: 'equal', formulae: ['"進行中"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'EDE9FE' } }, font: { color: { argb: '6D28D9' } } }, priority: 7 },
      { type: 'cellIs', operator: 'equal', formulae: ['"完了"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'DCFCE7' } }, font: { color: { argb: '15803D' } } }, priority: 8 },
    ],
  });

  // 条件付き書式：制作ステータス列
  ws1.addConditionalFormatting({
    ref: `H2:H${DATA_ROWS + 1}`,
    rules: [
      { type: 'cellIs', operator: 'equal', formulae: ['"進行中"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'CFFAFE' } }, font: { color: { argb: '0E7490' } } }, priority: 9 },
      { type: 'cellIs', operator: 'equal', formulae: ['"完了"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'DCFCE7' } }, font: { color: { argb: '15803D' } } }, priority: 10 },
    ],
  });

  // 条件付き書式：設計期限超過（期限 < 今日 かつ 設計ステータスが完了でない）
  ws1.addConditionalFormatting({
    ref: `E2:E${DATA_ROWS + 1}`,
    rules: [
      { type: 'expression', formulae: ['AND(E2<>"",E2<TODAY(),F2<>"完了")'], style: { font: { color: { argb: 'DC2626' }, bold: true } }, priority: 11 },
      { type: 'expression', formulae: ['AND(E2<>"",E2-TODAY()<=3,E2-TODAY()>=0,F2<>"完了")'], style: { font: { color: { argb: 'D97706' }, bold: true } }, priority: 12 },
    ],
  });

  // 条件付き書式：制作期限超過
  ws1.addConditionalFormatting({
    ref: `G2:G${DATA_ROWS + 1}`,
    rules: [
      { type: 'expression', formulae: ['AND(G2<>"",G2<TODAY(),H2<>"完了")'], style: { font: { color: { argb: 'DC2626' }, bold: true } }, priority: 13 },
      { type: 'expression', formulae: ['AND(G2<>"",G2-TODAY()<=3,G2-TODAY()>=0,H2<>"完了")'], style: { font: { color: { argb: 'D97706' }, bold: true } }, priority: 14 },
    ],
  });

  // オートフィルター
  ws1.autoFilter = 'A1:M1';

  // 印刷設定
  ws1.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1 };


  // =============================================================
  //  シート2: 製品別サマリー（COUNTIFS関数で自動集計）
  // =============================================================
  const ws2 = wb.addWorksheet('製品別サマリー', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });

  ws2.columns = [
    { header: '製品名',     key: 'product',     width: 26 },
    { header: 'アイテム数', key: 'itemCount',   width: 12 },
    { header: '設計完了',   key: 'designDone',  width: 12 },
    { header: '制作完了',   key: 'makeDone',    width: 12 },
    { header: '出荷済',     key: 'shipped',     width: 10 },
    { header: '進捗率',     key: 'progress',    width: 10 },
    { header: '最終制作期限', key: 'lastDL',    width: 16 },
    { header: '備考',       key: 'note',        width: 30 },
  ];

  styleHeader(ws2.getRow(1), 8);

  // 数式でデータを自動集計（製品名を入力すると集計される）
  for (let r = 2; r <= 51; r++) {
    // アイテム数: 進捗管理シートの製品名列をCOUNTIF
    ws2.getCell(`B${r}`).value = { formula: `IF(A${r}="","",COUNTIF('進捗管理'!B:B,A${r}))` };
    // 設計完了数
    ws2.getCell(`C${r}`).value = { formula: `IF(A${r}="","",COUNTIFS('進捗管理'!B:B,A${r},'進捗管理'!F:F,"完了"))` };
    // 制作完了数
    ws2.getCell(`D${r}`).value = { formula: `IF(A${r}="","",COUNTIFS('進捗管理'!B:B,A${r},'進捗管理'!H:H,"完了"))` };
    // 出荷済数
    ws2.getCell(`E${r}`).value = { formula: `IF(A${r}="","",COUNTIFS('進捗管理'!B:B,A${r},'進捗管理'!I:I,"出荷済"))` };
    // 進捗率（制作完了 / アイテム数）
    ws2.getCell(`F${r}`).value = { formula: `IF(B${r}="","",IF(B${r}=0,"",D${r}/B${r}))` };
    ws2.getCell(`F${r}`).numFmt = '0%';
    // 最終制作期限
    ws2.getCell(`G${r}`).value = { formula: `IF(A${r}="","",IFERROR(MAXIFS('進捗管理'!G:G,'進捗管理'!B:B,A${r}),""))` };
    ws2.getCell(`G${r}`).numFmt = 'yyyy/mm/dd';
  }

  styleDataRows(ws2, 2, 51, 8);

  // 進捗率の条件付き書式
  ws2.addConditionalFormatting({
    ref: 'F2:F51',
    rules: [
      { type: 'cellIs', operator: 'equal', formulae: ['1'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'DCFCE7' } }, font: { color: { argb: '15803D' }, bold: true } }, priority: 15 },
      { type: 'cellIs', operator: 'greaterThanOrEqual', formulae: ['0.5'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'DBEAFE' } }, font: { color: { argb: '1D4ED8' } } }, priority: 16 },
      { type: 'cellIs', operator: 'lessThan', formulae: ['0.5'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FEF3C7' } }, font: { color: { argb: 'D97706' } } }, priority: 17 },
    ],
  });


  // =============================================================
  //  シート3: 使い方ガイド
  // =============================================================
  const ws3 = wb.addWorksheet('使い方');

  const guide = [
    ['金型技術 製作進捗管理 — 使い方ガイド'],
    [''],
    ['■ 進捗管理シート（メイン）'],
    ['  1. 技術部が依頼を登録するとき → 受注番号・製品名・製作アイテム・種類を入力'],
    ['  2. 設計期限・制作期限を入力'],
    ['  3. 設計ステータス（未着手/進行中/完了）をドロップダウンから選択'],
    ['  4. 制作ステータス（未着手/進行中/完了）をドロップダウンから選択'],
    ['  5. 全体ステータス（未着手〜出荷済）をドロップダウンから選択'],
    ['  ※ ステータスに応じてセルの色が自動で変わります'],
    ['  ※ 期限超過は赤字、3日以内は橙字で表示されます'],
    [''],
    ['■ 製品別サマリーシート'],
    ['  - A列に製品名を入力すると、進捗管理シートから自動集計されます'],
    ['  - アイテム数・設計完了数・制作完了数・出荷済数・進捗率が自動計算'],
    [''],
    ['■ フィルター機能'],
    ['  - ヘッダー行の▼ボタンで、ステータスや種類で絞り込みができます'],
    [''],
    ['■ 種類の略称'],
    ['  Die = 金型  |  TP = トリム型  |  Tem = ベント落とし'],
    ['  Mill = 加工治具  |  GC = 切断治具  |  Ins = 検査治具'],
    ['  S = 強度試験  |  Other = その他'],
    [''],
    ['■ ステータスの流れ'],
    ['  未着手 → 設計中 → 設計完了 → 制作中 → 制作完了 → 出荷済'],
  ];

  guide.forEach((row, i) => {
    ws3.getCell(`A${i+1}`).value = row[0];
    if (i === 0) {
      ws3.getCell('A1').font = { size: 14, bold: true, color: { argb: C.headerBg } };
    }
  });
  ws3.getColumn(1).width = 80;

  // ===== 保存 =====
  const outPath = path.join(__dirname, '金型技術_製作進捗管理.xlsx');
  await wb.xlsx.writeFile(outPath);
  console.log('生成完了:', outPath);
}

main().catch(err => { console.error('エラー:', err); process.exit(1); });
