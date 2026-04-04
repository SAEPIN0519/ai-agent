// ============================================================
// 金型技術 製作進捗管理システム — メインスクリプト
// Google Apps Script (GAS) 版
// ============================================================

// ---- シート名定数 ----
var SHEET_PROGRESS = '進捗管理';
var SHEET_PRODUCT  = '製品マスター';
var SHEET_LEADTIME = 'リードタイム';

// ---- ヘッダー定義 ----
var HEADERS_PROGRESS = [
  'No', '受注番号', '製品品番', '製作アイテム', '種類',
  '設計担当', '製作担当', '依頼部署', '納期',
  '設計開始日', '設計完了日', '設計ST',
  '制作開始日', '制作完了日', '制作ST',
  '全体ST', 'メモ'
];
var HEADERS_PRODUCT = [
  '品番', '設計担当者', '製作担当者', '依頼担当者名', '依頼部署', '納期', 'メモ', '登録日'
];
var HEADERS_LEADTIME = [
  '種類コード', '種類名', '設計工数（日）', '製作工数（日）'
];

// ---- リードタイム初期データ ----
var LEADTIME_DATA = [
  ['Die',   'ダイ',       30, 45],
  ['TP',    'トリムパンチ', 20, 30],
  ['Tem',   'テンプレート', 15, 20],
  ['Mill',  'ミーリング',  15, 25],
  ['GC',    'ゲージ',      15, 20],
  ['Ins',   'インサート',  10, 15],
  ['S',     '標準',        10, 15],
  ['Other', 'その他',      10, 15]
];

// ============================================================
// メニュー追加（スプレッドシート起動時に自動実行）
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('金型技術')
    .addItem('製品登録', 'showProductDialog')
    .addItem('アイテム追加', 'showItemDialog')
    .addSeparator()
    .addItem('初期設定', 'setupSheets')
    .addToUi();
}

// ============================================================
// ダイアログ表示
// ============================================================

// 製品登録ダイアログを表示
function showProductDialog() {
  var html = HtmlService.createHtmlOutputFromFile('登録画面')
    .setWidth(560)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '製品・アイテム登録');
}

// アイテム追加ダイアログを表示（製品登録と同じHTMLを流用）
function showItemDialog() {
  var html = HtmlService.createHtmlOutputFromFile('登録画面')
    .setWidth(560)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '製品・アイテム登録');
}

// ============================================================
// GASサーバー側関数（HTMLから呼び出す）
// ============================================================

// 製品マスターの全品番リストを返す
function getProductList() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PRODUCT);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map(function(row) { return row[0]; }).filter(function(v) { return v !== ''; });
}

// 製品マスターから品番に対応する情報を返す
function getProductInfo(productCode) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PRODUCT);
  if (!sheet) return null;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(productCode)) {
      return {
        productCode:  data[i][0],
        designPerson: data[i][1],
        makePerson:   data[i][2],
        requesterName: data[i][3],
        department:   data[i][4],
        dueDate:      data[i][5] ? Utilities.formatDate(new Date(data[i][5]), 'Asia/Tokyo', 'yyyy/MM/dd') : '',
        memo:         data[i][6]
      };
    }
  }
  return null;
}

// 製品マスターに新しい製品を登録する
function registerProduct(formData) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PRODUCT);
  if (!sheet) {
    return { success: false, message: 'シートが見つかりません。初期設定を先に実行してください。' };
  }

  // 重複チェック
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i][0]) === String(formData.productCode)) {
        return { success: false, message: '品番「' + formData.productCode + '」はすでに登録されています。' };
      }
    }
  }

  var dueDate = formData.dueDate ? new Date(formData.dueDate) : '';
  var today   = new Date();

  sheet.appendRow([
    formData.productCode,
    formData.designPerson,
    formData.makePerson,
    formData.requesterName,
    formData.department,
    dueDate,
    formData.memo,
    today
  ]);

  // 日付セルのフォーマット設定
  var newRow = sheet.getLastRow();
  if (dueDate) {
    sheet.getRange(newRow, 6).setNumberFormat('yyyy/MM/dd');
  }
  sheet.getRange(newRow, 8).setNumberFormat('yyyy/MM/dd');

  return { success: true, message: '製品「' + formData.productCode + '」を登録しました。' };
}

// 進捗管理シートにアイテムを追加する
function addProgressItem(formData) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sheet    = ss.getSheetByName(SHEET_PROGRESS);
  if (!sheet) {
    return { success: false, message: 'シートが見つかりません。初期設定を先に実行してください。' };
  }

  // 製品マスターから担当者・部署・納期を取得
  var productInfo = getProductInfo(formData.productCode);
  if (!productInfo) {
    return { success: false, message: '品番「' + formData.productCode + '」が製品マスターに見つかりません。' };
  }

  var newRow   = sheet.getLastRow() + 1;
  var rowIndex = newRow; // 数式用の行番号

  // A列（No）: 同一品番内の連番数式
  var noFormula = '=COUNTIF($C$2:C' + rowIndex + ',C' + rowIndex + ')';

  // K列（設計完了日）の数式
  var designEndFormula = '=IF(OR(J' + rowIndex + '="",E' + rowIndex + '=""),"",J' + rowIndex +
    '+IFERROR(VLOOKUP(E' + rowIndex + ',リードタイム!$A$2:$C$9,3,FALSE),0))';

  // N列（制作完了日）の数式
  var makeEndFormula = '=IF(OR(M' + rowIndex + '="",E' + rowIndex + '=""),"",M' + rowIndex +
    '+IFERROR(VLOOKUP(E' + rowIndex + ',リードタイム!$A$2:$D$9,4,FALSE),0))';

  var dueDate = productInfo.dueDate ? new Date(productInfo.dueDate) : '';

  // 行データをセット（数式は別途セット）
  sheet.appendRow([
    '',                       // A: No（数式）
    formData.orderNo,         // B: 受注番号
    formData.productCode,     // C: 製品品番
    formData.itemName,        // D: 製作アイテム
    formData.itemType,        // E: 種類
    productInfo.designPerson, // F: 設計担当
    productInfo.makePerson,   // G: 製作担当
    productInfo.department,   // H: 依頼部署
    dueDate,                  // I: 納期
    '',                       // J: 設計開始日
    '',                       // K: 設計完了日（数式）
    '未着手',                 // L: 設計ST
    '',                       // M: 制作開始日
    '',                       // N: 制作完了日（数式）
    '未着手',                 // O: 制作ST
    '未着手',                 // P: 全体ST
    formData.memo             // Q: メモ
  ]);

  // 数式をセット
  sheet.getRange(rowIndex, 1).setFormula(noFormula);
  sheet.getRange(rowIndex, 11).setFormula(designEndFormula);
  sheet.getRange(rowIndex, 14).setFormula(makeEndFormula);

  // 日付フォーマット
  if (dueDate) {
    sheet.getRange(rowIndex, 9).setNumberFormat('yyyy/MM/dd');
  }
  sheet.getRange(rowIndex, 10).setNumberFormat('yyyy/MM/dd');
  sheet.getRange(rowIndex, 11).setNumberFormat('yyyy/MM/dd');
  sheet.getRange(rowIndex, 13).setNumberFormat('yyyy/MM/dd');
  sheet.getRange(rowIndex, 14).setNumberFormat('yyyy/MM/dd');

  return { success: true, message: 'アイテム「' + formData.itemName + '」を追加しました。' };
}

// ============================================================
// 初期設定（シート作成・スタイル設定）
// ============================================================
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // 3シートを作成（既存なら作成しない）
  createSheetIfNotExists(ss, SHEET_LEADTIME);
  createSheetIfNotExists(ss, SHEET_PRODUCT);
  createSheetIfNotExists(ss, SHEET_PROGRESS);

  // 各シートを設定
  setupLeadtimeSheet(ss);
  setupProductSheet(ss);
  setupProgressSheet(ss);

  ui.alert('初期設定が完了しました。\n\n・リードタイムシート\n・製品マスターシート\n・進捗管理シート\n\n以上3つのシートを設定しました。');
}

// シートが存在しなければ作成する
function createSheetIfNotExists(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    ss.insertSheet(sheetName);
  }
}

// ---- リードタイムシート設定 ----
function setupLeadtimeSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_LEADTIME);

  // ヘッダー設定
  sheet.clearContents();
  sheet.getRange(1, 1, 1, HEADERS_LEADTIME.length).setValues([HEADERS_LEADTIME]);

  // ヘッダースタイル（ダークネイビー）
  var headerRange = sheet.getRange(1, 1, 1, HEADERS_LEADTIME.length);
  headerRange.setBackground('#1E3A5F')
             .setFontColor('#FFFFFF')
             .setFontWeight('bold')
             .setHorizontalAlignment('center');

  // 初期データ投入（データがなければ）
  if (sheet.getLastRow() < 2) {
    sheet.getRange(2, 1, LEADTIME_DATA.length, 4).setValues(LEADTIME_DATA);
  }

  // 列幅設定
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);

  // 1行目固定
  sheet.setFrozenRows(1);
}

// ---- 製品マスターシート設定 ----
function setupProductSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_PRODUCT);

  // ヘッダー設定
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS_PRODUCT.length).setValues([HEADERS_PRODUCT]);
  }

  // ヘッダースタイル
  var headerRange = sheet.getRange(1, 1, 1, HEADERS_PRODUCT.length);
  headerRange.setBackground('#1E3A5F')
             .setFontColor('#FFFFFF')
             .setFontWeight('bold')
             .setHorizontalAlignment('center');

  // 列幅設定
  var colWidths = [100, 80, 80, 100, 100, 100, 160, 100];
  for (var i = 0; i < colWidths.length; i++) {
    sheet.setColumnWidth(i + 1, colWidths[i]);
  }

  // 1行目固定
  sheet.setFrozenRows(1);
}

// ---- 進捗管理シート設定（メイン） ----
function setupProgressSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_PROGRESS);

  // ヘッダー設定
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS_PROGRESS.length).setValues([HEADERS_PROGRESS]);
  }

  // ---- ヘッダー色分け ----
  // 基本情報（A-I）: ダークネイビー
  sheet.getRange(1, 1, 1, 9)
       .setBackground('#1E3A5F')
       .setFontColor('#FFFFFF')
       .setFontWeight('bold')
       .setHorizontalAlignment('center');

  // 設計系（J-L）: 紫
  sheet.getRange(1, 10, 1, 3)
       .setBackground('#6D28D9')
       .setFontColor('#FFFFFF')
       .setFontWeight('bold')
       .setHorizontalAlignment('center');

  // 制作系（M-O）: シアン
  sheet.getRange(1, 13, 1, 3)
       .setBackground('#0E7490')
       .setFontColor('#FFFFFF')
       .setFontWeight('bold')
       .setHorizontalAlignment('center');

  // 全体ST（P）: アンバー
  sheet.getRange(1, 16, 1, 1)
       .setBackground('#B45309')
       .setFontColor('#FFFFFF')
       .setFontWeight('bold')
       .setHorizontalAlignment('center');

  // メモ（Q）: ダークネイビー
  sheet.getRange(1, 17, 1, 1)
       .setBackground('#1E3A5F')
       .setFontColor('#FFFFFF')
       .setFontWeight('bold')
       .setHorizontalAlignment('center');

  // ---- 列幅設定 ----
  var colWidths = [
    40,  // A: No
    90,  // B: 受注番号
    100, // C: 製品品番
    160, // D: 製作アイテム
    60,  // E: 種類
    70,  // F: 設計担当
    70,  // G: 製作担当
    90,  // H: 依頼部署
    100, // I: 納期
    100, // J: 設計開始日
    100, // K: 設計完了日
    80,  // L: 設計ST
    100, // M: 制作開始日
    100, // N: 制作完了日
    80,  // O: 制作ST
    80,  // P: 全体ST
    200  // Q: メモ
  ];
  for (var i = 0; i < colWidths.length; i++) {
    sheet.setColumnWidth(i + 1, colWidths[i]);
  }

  // ---- 条件付き書式設定 ----
  setupConditionalFormats(sheet);

  // ---- データ入力規則（ドロップダウン） ----
  setupDataValidation(sheet);

  // ---- オートフィルター設定 ----
  var lastCol = HEADERS_PROGRESS.length;
  sheet.getRange(1, 1, 1, lastCol).createFilter();

  // ---- 1行目固定 ----
  sheet.setFrozenRows(1);
}

// 条件付き書式を設定する
function setupConditionalFormats(sheet) {
  // 既存の条件付き書式をクリア
  sheet.clearConditionalFormatRules();
  var rules = [];
  var maxRow = 1000; // 適用対象の最大行数

  // ---- 全体ST（P列=16列目）の色分け ----
  var stRange = sheet.getRange(2, 16, maxRow, 1);
  var stColors = [
    { value: '未着手',  bg: '#9CA3AF', font: '#FFFFFF' }, // グレー
    { value: '設計中',  bg: '#7C3AED', font: '#FFFFFF' }, // 紫
    { value: '設計完了', bg: '#2563EB', font: '#FFFFFF' }, // 青
    { value: '制作中',  bg: '#0891B2', font: '#FFFFFF' }, // シアン
    { value: '制作完了', bg: '#059669', font: '#FFFFFF' }, // 緑
    { value: '出荷済',  bg: '#6B7280', font: '#FFFFFF' }  // 灰
  ];
  stColors.forEach(function(c) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(c.value)
        .setBackground(c.bg)
        .setFontColor(c.font)
        .setRanges([stRange])
        .build()
    );
  });

  // ---- 設計ST（L列=12列目）の色分け ----
  var designStRange = sheet.getRange(2, 12, maxRow, 1);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('進行中')
      .setBackground('#7C3AED')
      .setFontColor('#FFFFFF')
      .setRanges([designStRange])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('完了')
      .setBackground('#059669')
      .setFontColor('#FFFFFF')
      .setRanges([designStRange])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('未着手')
      .setBackground('#9CA3AF')
      .setFontColor('#FFFFFF')
      .setRanges([designStRange])
      .build()
  );

  // ---- 制作ST（O列=15列目）の色分け ----
  var makeStRange = sheet.getRange(2, 15, maxRow, 1);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('進行中')
      .setBackground('#0891B2')
      .setFontColor('#FFFFFF')
      .setRanges([makeStRange])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('完了')
      .setBackground('#059669')
      .setFontColor('#FFFFFF')
      .setRanges([makeStRange])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('未着手')
      .setBackground('#9CA3AF')
      .setFontColor('#FFFFFF')
      .setRanges([makeStRange])
      .build()
  );

  // ---- 納期超過（I列=9列目）: 今日より前なら赤文字 ----
  var dueDateRange = sheet.getRange(2, 9, maxRow, 1);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenDateBefore(SpreadsheetApp.RelativeDate.TODAY)
      .setFontColor('#DC2626')
      .setRanges([dueDateRange])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}

// データ入力規則（ドロップダウン）を設定する
function setupDataValidation(sheet) {
  var maxRow = 1000;

  // 種類（E列=5列目）: リードタイムシートの種類コードからドロップダウン
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Die', 'TP', 'Tem', 'Mill', 'GC', 'Ins', 'S', 'Other'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 5, maxRow, 1).setDataValidation(typeRule);

  // 全体ST（P列=16列目）: ドロップダウン
  var overallStRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未着手', '設計中', '設計完了', '制作中', '制作完了', '出荷済'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 16, maxRow, 1).setDataValidation(overallStRule);

  // 設計ST（L列=12列目）: ドロップダウン
  var designStRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未着手', '進行中', '完了'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 12, maxRow, 1).setDataValidation(designStRule);

  // 制作ST（O列=15列目）: ドロップダウン
  var makeStRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未着手', '進行中', '完了'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 15, maxRow, 1).setDataValidation(makeStRule);
}
