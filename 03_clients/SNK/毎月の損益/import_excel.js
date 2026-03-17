/**
 * SNK ダイカスト事業部 Excel → ダッシュボード更新スクリプト
 *
 * 使い方:
 *   node import_excel.js "C:\Users\matsuoka\Desktop\R8.2月　損益.xlsx"
 *   ※ 引数省略時は Desktop\ai-agent 内の最新の損益.xlsx を自動検索
 *
 * 更新対象:
 *   ① 02_clients/SNK/diecast_dashboard.html（5部門グラフダッシュボード）
 *   ② 毎月の損益/monthly_dashboard.html（単体損益ダッシュボード）
 */

import { readFileSync, writeFileSync, readdirSync } from "fs";
import { read, utils } from "../../../node_modules/xlsx/xlsx.mjs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SNK_DIR   = path.resolve(__dirname, "..");
const HTML_PATH = path.join(SNK_DIR, "diecast_dashboard.html");

// =====================
// 列インデックス（Excel ヘッダー行より）
// DC本部[1] 営業部[2] 品質管理部[3] 本社検査[4] 物流検査[5]
// 技術部[6] 生産管理部[7] 型技術課[8] ダイカスト計[9]
// 糸貫工場[10] 糸貫検査[11] 加工[12] ダイカスト合計[13] ...
// =====================
const DIV_COL = {
  糸貫工場: 10,
  糸貫検査: 11,
  加工:     12,
  本社検査:  4,
  物流検査:  5,
};

// Excel の科目名 → ダッシュボードのキー名
const ROW_MAP = {
  "売上総合計":         "売上",
  "売上原価":           "原価",
  "売上総利益":         "粗利",
  "営業利益":           "営業利益",
  "給与基本給":         "基本給",
  "時間外・深夜手当":   "時間外",
  "雑給":               "雑給",
  "賞与引当金":         "賞与引当",
  "法定福利費":         "法定福利",
  "退職金":             "退職金",
  "消耗品費":           "消耗品",
  "修繕費":             "修繕費",
  "地代家賃":           "地代家賃",
  "減価償却":           "減価償却",
  "電力料+水道光熱費":  "電力",
  "燃料費":             "燃料",
};

// ダッシュボードに必要な全キー（0 で初期化）
const ALL_KEYS = [
  "売上","原価","粗利","営業利益",
  "基本給","時間外","雑給","賞与引当","法定福利","退職金",
  "消耗品","修繕費","地代家賃","減価償却","電力","燃料",
];

// =====================
// Excel 解析
// =====================
function parseExcel(filePath) {
  const wb = read(readFileSync(filePath));
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = utils.sheet_to_json(ws, { header: 1, defval: "" });

  // 月ラベルを取得（行0: 令和8年2月 → "2月"）
  const title = String(rows[0][1] || "");
  const monthMatch = title.match(/(\d+)月/);
  const yearMatch  = title.match(/令和(\d+)年/);
  const month  = monthMatch ? monthMatch[1] + "月" : "不明";
  const nengo  = yearMatch  ? `R${yearMatch[1]}.${month}` : month;

  console.log(`📄 読み込み: ${path.basename(filePath)}`);
  console.log(`📅 月: ${nengo}（表示名: ${month}）`);

  // 各部門のデータを抽出
  const extracted = {};
  for (const div of Object.keys(DIV_COL)) {
    extracted[div] = {};
    for (const key of ALL_KEYS) extracted[div][key] = 0;
  }

  // 科目行が重複する場合（給与基本給が製造と販管で2行ある）→ 製造側（先に出現）を使う
  const seen = new Set();

  for (let i = 2; i < rows.length; i++) {
    const rowLabel = String(rows[i][0]).trim();
    const dashKey  = ROW_MAP[rowLabel];
    if (!dashKey) continue;

    // 同じキーが既に登録済みならスキップ（製造の給与基本給を優先）
    if (seen.has(dashKey) && ["基本給","時間外","雑給","賞与引当","退職金","法定福利"].includes(dashKey)) {
      continue;
    }
    seen.add(dashKey);

    for (const [div, col] of Object.entries(DIV_COL)) {
      const raw = rows[i][col];
      extracted[div][dashKey] = (typeof raw === "number") ? Math.round(raw) : 0;
    }
  }

  return { month, nengo, data: extracted };
}

// =====================
// ダッシュボード HTML 更新
// =====================
function updateDashboardHtml(month, nengo, newData) {
  let html = readFileSync(HTML_PATH, "utf-8");

  // 現在の Ms と D を取得
  const msMatch = html.match(/const Ms = \[(.*?)\];/);
  const dMatch  = html.match(/const D = \{([\s\S]*?)\};/);
  if (!msMatch || !dMatch) throw new Error("HTMLのデータ部分が見つかりません。");

  const currentMs = msMatch[1].match(/'([^']+)'/g)?.map(s => s.replace(/'/g, "")) ?? [];

  // 既に存在する月かチェック
  const exists = currentMs.includes(month);
  const idx    = exists ? currentMs.indexOf(month) : currentMs.length;
  const newMs  = exists ? currentMs : [...currentMs, month];

  console.log(`\n📊 ダッシュボード更新:`);
  console.log(`  現在の月: [${currentMs.join(", ")}]`);
  console.log(`  → ${exists ? `${month}のデータを上書き` : `${month}を追加`}`);
  console.log(`  更新後:   [${newMs.join(", ")}]`);

  // D 構造を再構築
  // まず既存 D をパース
  const existingD = {};
  const divPattern = /(\S+): \{([^{}]+)\}/g;
  let m;
  while ((m = divPattern.exec(dMatch[1])) !== null) {
    const div = m[1];
    const body = m[2];
    existingD[div] = {};
    const itemPattern = /(\S+):\s*\[([\d\s,\-]+)\]/g;
    let im;
    while ((im = itemPattern.exec(body)) !== null) {
      existingD[div][im[1]] = im[2].split(",").map(v => parseInt(v.trim(), 10));
    }
  }

  // 新しい D を構築
  const DIVS = ["糸貫工場", "加工", "糸貫検査", "本社検査", "物流検査"];
  let dStr = "const D = {\n";
  for (const div of DIVS) {
    dStr += `  ${div}: {\n`;
    for (const key of ALL_KEYS) {
      const existing = existingD[div]?.[key] ?? [];
      // 配列を新しい月数に合わせる
      while (existing.length < idx) existing.push(0);
      if (exists) {
        existing[idx] = newData[div]?.[key] ?? 0;
      } else {
        existing.push(newData[div]?.[key] ?? 0);
      }
      const padded = existing.slice(0, newMs.length);
      dStr += `    ${key}:   [${padded.join(", ")}],\n`;
    }
    dStr += "  },\n";
  }
  dStr += "};";

  // Ms を更新
  const msStr = `const Ms = [${newMs.map(m => `'${m}'`).join(", ")}];`;
  html = html.replace(/const Ms = \[.*?\];/, msStr);

  // M（長い月ラベル配列）も更新 - nengo を使う
  const currentM = (html.match(/const M = \[(.*?)\];/)?.[1].match(/'([^']+)'/g)?.map(s => s.replace(/'/g, "")) ?? []);
  if (exists) {
    currentM[idx] = nengo;
  } else {
    currentM.push(nengo);
  }
  const mStr = `const M = [${currentM.map(m => `'${m}'`).join(", ")}];`;
  html = html.replace(/const M = \[.*?\];/, mStr);

  // D を更新
  html = html.replace(/const D = \{[\s\S]*?\};/, dStr);

  writeFileSync(HTML_PATH, html, "utf-8");
  console.log(`✅ ${HTML_PATH.split("\\").pop()} を更新しました`);
}

// =====================
// メイン
// =====================
async function main() {
  // 引数またはデフォルトパス
  let filePath = process.argv[2];

  if (!filePath) {
    // ai-agent フォルダ内の損益.xlsx を自動検索
    const base = path.resolve(__dirname, "../../..");
    try {
      const files = readdirSync(base).filter(f => f.includes("損益") && f.endsWith(".xlsx"));
      if (files.length > 0) {
        filePath = path.join(base, files[0]);
        console.log(`📁 自動検出: ${filePath}`);
      }
    } catch (e) { /* ignore */ }
  }

  if (!filePath) {
    console.error("❌ Excelファイルが見つかりません。");
    console.error("使い方: node import_excel.js \"C:\\パス\\損益.xlsx\"");
    process.exit(1);
  }

  // 解析
  let parsed;
  try {
    parsed = parseExcel(filePath);
  } catch (e) {
    console.error(`❌ Excel読み込みエラー: ${e.message}`);
    process.exit(1);
  }

  const { month, nengo, data } = parsed;

  // プレビュー
  console.log("\n【抽出データ プレビュー（売上・営業利益）】");
  for (const div of ["糸貫工場","加工","糸貫検査","本社検査","物流検査"]) {
    const s = data[div].売上;
    const o = data[div].営業利益;
    const fmt = n => (n === 0 ? "—" : (n < 0 ? "▲" : "") + Math.abs(Math.round(n/10000)).toLocaleString() + "万");
    console.log(`  ${div.padEnd(5)}: 売上 ${fmt(s)} / 営業利益 ${fmt(o)}`);
  }

  // 確認
  const readline = await import("readline");
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  await new Promise(resolve => {
    rl.question(`\nこの内容でダッシュボードを更新しますか？ [y/n]: `, ans => {
      rl.close();
      if (!["y", "yes", ""].includes(ans.trim().toLowerCase())) {
        console.log("キャンセルしました。");
        process.exit(0);
      }
      resolve();
    });
  });

  // HTML 更新
  try {
    updateDashboardHtml(month, nengo, data);
  } catch (e) {
    console.error(`❌ HTML更新エラー: ${e.message}`);
    process.exit(1);
  }

  console.log("\nブラウザで diecast_dashboard.html を F5 すると反映されます。");
}

main();
