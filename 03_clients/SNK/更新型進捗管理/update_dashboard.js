/**
 * 更新型進捗管理ダッシュボード — 自動更新スクリプト
 *
 * OneDrive同期されたExcelを読み込み、dashboard.htmlのデータ部分を更新する。
 *
 * 使い方:
 *   node update_dashboard.js
 *   または update_dashboard.bat をダブルクリック
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
const HTML_PATH = path.join(__dirname, 'dashboard.html');
const SHEET_NAME = '更新型';

// === ユーティリティ ===
function toDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'number') return new Date((v - 25569) * 86400000);
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

function fmtDate(d) {
  return d ? d.toISOString().split('T')[0] : '';
}

// === メイン処理 ===
function main() {
  console.log('--- 更新型ダッシュボード更新 ---');
  console.log('Excel: ' + EXCEL_PATH);

  // Excelが存在するか確認
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

  const data = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false });
  const rows = data.slice(2); // ヘッダー2行をスキップ
  const today = new Date();

  // データ変換
  const allData = rows.filter(r => r[0] != null && r[0] !== '').map(r => {
    const kijitsu = toDate(r[14]);
    const okure = kijitsu && kijitsu < today ? Math.floor((today - kijitsu) / 86400000) : 0;
    const tag = String(r[37] || '').trim();

    let kanryo;
    if (tag === '完了') kanryo = '完了';
    else if (tag === '進行中') kanryo = '進行中';
    else kanryo = '未設定';

    return {
      no: r[0],
      kyakusaki: String(r[3] || ''),
      tantou: String(r[5] || '').trim(),
      hinban: String(r[6] || ''),
      hinmei: String(r[7] || ''),
      koushinGata: String(r[8] || ''),
      status: String(r[12] || '').trim() || '(なし)',
      shinchoku: String(r[13] || '').replace(/\n/g, ' ').substring(0, 100),
      kijitsu: fmtDate(kijitsu),
      okure: okure,
      kanryo: kanryo,
    };
  });

  // 集計表示
  const shinko = allData.filter(r => r.kanryo === '進行中');
  const kanryo = allData.filter(r => r.kanryo === '完了');
  console.log('全案件: ' + allData.length);
  console.log('進行中: ' + shinko.length);
  console.log('完了: ' + kanryo.length);
  console.log('期日遅れ(進行中): ' + shinko.filter(r => r.okure > 0).length);

  // HTML更新
  if (!fs.existsSync(HTML_PATH)) {
    console.error('エラー: dashboard.html が見つかりません。');
    process.exit(1);
  }

  let html = fs.readFileSync(HTML_PATH, 'utf8');
  const jsonStr = JSON.stringify(allData);
  html = html.replace(/const RAW_DATA = \[.*?\];/s, 'const RAW_DATA = ' + jsonStr + ';');
  fs.writeFileSync(HTML_PATH, html);

  console.log('dashboard.html 更新完了 (' + Math.round(fs.statSync(HTML_PATH).size / 1024) + 'KB)');
  console.log('--- 完了 ---');
}

main();
