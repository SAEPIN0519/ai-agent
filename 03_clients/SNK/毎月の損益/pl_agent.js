/**
 * SNK ダイカスト事業部 月次損益 AI エージェント
 *
 * 使い方: node pl_agent.js
 *
 * できること:
 *   - エクセルデータの貼り付け → CSV & HTML ダッシュボード自動更新
 *   - 損益データの確認・分析・サマリー表示
 */

import Anthropic from "@anthropic-ai/sdk";
import fs from "fs";
import path from "path";
import readline from "readline";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CSV_PATH = path.join(__dirname, "ダイカスト事業部損益_2026.csv");
const HTML_PATH = path.join(__dirname, "monthly_dashboard.html");

const client = new Anthropic();
const MODEL = "claude-opus-4-6";

// =====================
// ユーティリティ
// =====================

function fmt(n) {
  if (n === 0) return "—";
  const sign = n < 0 ? "▲" : "";
  return sign + Math.abs(Math.round(n / 10000)).toLocaleString() + "万円";
}

// =====================
// データ解析
// =====================

/**
 * タブ区切りのエクセルデータをパース
 * @param {string} text - エクセルからコピーしたタブ区切りデータ
 * @returns {{ months: string[], data: Object }}
 */
function parseTsv(text) {
  const lines = text
    .trim()
    .split("\n")
    .filter((l) => l.trim());
  if (lines.length === 0) throw new Error("データが空です。");

  const rows = lines.map((l) => l.split("\t"));
  const header = rows[0];

  // 月リスト（合計列は除外）
  const months = [];
  const monthIndices = [];
  for (let i = 1; i < header.length; i++) {
    const h = header[i].trim();
    if (h && h !== "合計") {
      months.push(h);
      monthIndices.push(i);
    }
  }

  if (months.length === 0)
    throw new Error(
      "月が見つかりません。ヘッダー行に「4月」「5月」などを入れてください。"
    );

  // データ: { 科目: { 月: 数値 } }
  const data = {};
  for (const row of rows.slice(1)) {
    const item = row[0]?.trim();
    if (!item) continue;
    const values = {};
    for (let j = 0; j < months.length; j++) {
      const raw = (row[monthIndices[j]] || "").trim();
      const cleaned = raw.replace(/[,¥￥\s\u00a0]/g, "");
      values[months[j]] = cleaned ? Math.round(parseFloat(cleaned)) || 0 : 0;
    }
    data[item] = values;
  }

  return { months, data };
}

/**
 * 売上高合計・営業利益などの集計値を計算
 */
function calcAggregates(data, months) {
  const get = (item, m) => data[item]?.[m] ?? 0;

  const CALC = {
    "売上高 合計": (m) =>
      get("　製品売上", m) + get("　その他売上", m),
    "売上原価 合計": (m) =>
      get("　材料費", m) + get("　労務費", m) + get("　製造経費", m),
    売上総利益: (m) => get("売上高 合計", m) - get("売上原価 合計", m),
    "販管費 合計": (m) =>
      ["　人件費", "　地代家賃", "　減価償却費", "　光熱費", "　通信費", "　その他販管費"].reduce(
        (a, k) => a + get(k, m),
        0
      ),
    営業利益: (m) => get("売上総利益", m) - get("販管費 合計", m),
    経常利益: (m) =>
      get("営業利益", m) +
      get("　受取利息", m) +
      get("　その他営業外収益", m) -
      get("　支払利息", m) -
      get("　その他営業外費用", m),
  };

  for (const [item, fn] of Object.entries(CALC)) {
    data[item] = {};
    for (const m of months) {
      data[item][m] = fn(m);
    }
  }
  return data;
}

// =====================
// CSV 更新
// =====================

function updateCsv(months, data) {
  if (!fs.existsSync(CSV_PATH)) return "CSVファイルが見つかりません。";

  const content = fs.readFileSync(CSV_PATH, "utf-8");
  const lines = content.split("\n");
  const header = lines[0].split(",");

  // 月の列インデックスを取得
  const monthCol = {};
  for (let i = 0; i < header.length; i++) {
    const h = header[i].trim();
    if (months.includes(h)) monthCol[h] = i;
  }
  const totalIdx = header.findIndex((h) => h.trim() === "合計");

  let updated = 0;
  const newLines = lines.map((line, rowIdx) => {
    if (rowIdx === 0) return line;
    const cols = line.split(",");
    const item = cols[0]?.trim();
    if (!item || !data[item]) return line;

    let rowTotal = 0;
    for (const [m, colI] of Object.entries(monthCol)) {
      const val = data[item][m] ?? 0;
      cols[colI] = String(val);
      rowTotal += val;
    }
    if (totalIdx >= 0) cols[totalIdx] = String(rowTotal);
    updated++;
    return cols.join(",");
  });

  fs.writeFileSync(CSV_PATH, newLines.join("\n"), "utf-8");
  return updated;
}

// =====================
// HTML 生成
// =====================

function generateHtml(months, data) {
  const get = (item, m) => data[item]?.[m] ?? 0;

  const salesArr = months.map((m) => get("売上高 合計", m));
  const costArr = months.map((m) => get("売上原価 合計", m));
  const grossArr = months.map((m) => get("売上総利益", m));
  const opArr = months.map((m) => get("営業利益", m));
  const expArr = months.map((m) => get("販管費 合計", m));

  const latest = months[months.length - 1];
  const latestSales = get("売上高 合計", latest);
  const latestGross = get("売上総利益", latest);
  const latestOp = get("営業利益", latest);
  const pct = (a, b) => (b ? ((a / b) * 100).toFixed(1) + "%" : "—");

  // テーブル行
  const displayItems = [
    ["【売上高】", "section"],
    ["　製品売上", null],
    ["　その他売上", null],
    ["売上高 合計", "bold"],
    ["【売上原価】", "section"],
    ["　材料費", null],
    ["　労務費", null],
    ["　製造経費", null],
    ["売上原価 合計", "bold"],
    ["売上総利益", "highlight"],
    ["【販売費及び一般管理費】", "section"],
    ["　人件費", null],
    ["　地代家賃", null],
    ["　減価償却費", null],
    ["　光熱費", null],
    ["　通信費", null],
    ["　その他販管費", null],
    ["販管費 合計", "bold"],
    ["営業利益", "highlight"],
    ["【営業外収益】", "section"],
    ["　受取利息", null],
    ["　その他営業外収益", null],
    ["【営業外費用】", "section"],
    ["　支払利息", null],
    ["　その他営業外費用", null],
    ["経常利益", "highlight"],
  ];

  let tableRows = "";
  for (const [item, style] of displayItems) {
    if (style === "section") {
      tableRows += `<tr class="section-header"><td colspan="${months.length + 2}">${item}</td></tr>\n`;
      continue;
    }
    if (!data[item]) continue;
    const total = months.reduce((a, m) => a + get(item, m), 0);
    const cls = style === "bold" ? ' class="subtotal"' : style === "highlight" ? ' class="total-row"' : "";
    const cells = months.map((m) => {
      const v = get(item, m);
      return `<td class="${v < 0 ? "neg" : ""}">${fmt(v)}</td>`;
    }).join("");
    tableRows += `<tr${cls}><td class="item-name">${item}</td>${cells}<td class="${total < 0 ? "neg" : ""}">${fmt(total)}</td></tr>\n`;
  }

  const monthHeaders = months.map((m) => `<th>${m}</th>`).join("");
  const totalCumSales = salesArr.reduce((a, v) => a + v, 0);
  const totalCumOp = opArr.reduce((a, v) => a + v, 0);

  const html = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>月次損益ダッシュボード｜ダイカスト事業部</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Noto Sans JP', sans-serif; background: #f0f4f9; color: #0f172a; }
  header {
    background: linear-gradient(135deg, #0f172a, #1e3a6e, #1e40af);
    padding: 0 32px; height: 64px;
    display: flex; align-items: center; justify-content: space-between;
    position: sticky; top: 0; z-index: 100;
    box-shadow: 0 2px 20px rgba(15,23,42,.3);
  }
  .logo h1 { font-size: 15px; font-weight: 700; color: #fff; }
  .logo span { font-size: 11px; color: rgba(255,255,255,.5); display: block; margin-top: 2px; }
  .header-right { color: #fff; font-size: 13px; font-weight: 600; }
  .header-right small { display: block; font-size: 11px; color: rgba(255,255,255,.5); font-weight: 400; }
  .main { max-width: 1400px; margin: 0 auto; padding: 24px; }
  .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 24px; }
  .kpi-card { background: #fff; border-radius: 12px; padding: 20px; box-shadow: 0 1px 3px rgba(15,23,42,.06); border-left: 4px solid var(--c); }
  .kpi-label { font-size: 12px; color: #64748b; margin-bottom: 8px; }
  .kpi-value { font-size: 24px; font-weight: 700; color: var(--c); }
  .kpi-sub { font-size: 11px; color: #94a3b8; margin-top: 4px; }
  .neg-val { color: #ef4444 !important; }
  .charts { display: grid; grid-template-columns: 2fr 1fr; gap: 16px; margin-bottom: 24px; }
  .chart-card { background: #fff; border-radius: 12px; padding: 20px; box-shadow: 0 1px 3px rgba(15,23,42,.06); }
  .chart-card h3 { font-size: 13px; color: #334155; margin-bottom: 16px; font-weight: 600; }
  .chart-wrap { position: relative; height: 260px; }
  .table-card { background: #fff; border-radius: 12px; padding: 20px; box-shadow: 0 1px 3px rgba(15,23,42,.06); overflow-x: auto; }
  .table-card h3 { font-size: 13px; color: #334155; margin-bottom: 16px; font-weight: 600; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; }
  th { background: #1e40af; color: #fff; padding: 8px 12px; text-align: right; white-space: nowrap; font-weight: 500; }
  th:first-child { text-align: left; min-width: 160px; }
  td { padding: 7px 12px; text-align: right; border-bottom: 1px solid #e4e9f0; color: #334155; }
  td.item-name { text-align: left; color: #475569; }
  td.neg { color: #ef4444; }
  tr.section-header td { background: #f1f5f9; color: #64748b; font-size: 11px; padding: 5px 12px; font-weight: 600; }
  tr.subtotal td { background: #f8fafc; font-weight: 600; color: #1e40af; }
  tr.total-row td { background: #eff6ff; font-weight: 700; font-size: 13px; }
  tr.total-row td.neg { color: #ef4444; }
  tr:hover td { background: #f8fafc; }
  tr.section-header:hover td, tr.subtotal:hover td, tr.total-row:hover td { background: inherit; }
  .updated { font-size: 11px; color: #94a3b8; margin-top: 12px; text-align: right; }
</style>
</head>
<body>
<header>
  <div class="logo">
    <h1>月次損益ダッシュボード｜ダイカスト事業部</h1>
    <span>R8年度 事業部単体損益（税抜・円）</span>
  </div>
  <div class="header-right">
    最新月：${latest}
    <small>単位：万円（表示）</small>
  </div>
</header>
<div class="main">
  <div class="kpi-grid">
    <div class="kpi-card" style="--c:#3b82f6">
      <div class="kpi-label">売上高（${latest}）</div>
      <div class="kpi-value">${fmt(latestSales)}</div>
      <div class="kpi-sub">製品売上 + その他売上</div>
    </div>
    <div class="kpi-card" style="--c:${latestGross >= 0 ? "#10b981" : "#ef4444"}">
      <div class="kpi-label">売上総利益（${latest}）</div>
      <div class="kpi-value ${latestGross < 0 ? "neg-val" : ""}">${fmt(latestGross)}</div>
      <div class="kpi-sub">粗利率 ${pct(latestGross, latestSales)}</div>
    </div>
    <div class="kpi-card" style="--c:${latestOp >= 0 ? "#f59e0b" : "#ef4444"}">
      <div class="kpi-label">営業利益（${latest}）</div>
      <div class="kpi-value ${latestOp < 0 ? "neg-val" : ""}">${fmt(latestOp)}</div>
      <div class="kpi-sub">営業利益率 ${pct(latestOp, latestSales)}</div>
    </div>
    <div class="kpi-card" style="--c:#8b5cf6">
      <div class="kpi-label">累計売上高（${months[0]}〜${latest}）</div>
      <div class="kpi-value">${fmt(totalCumSales)}</div>
      <div class="kpi-sub">${months.length}ヶ月合計</div>
    </div>
    <div class="kpi-card" style="--c:${totalCumOp >= 0 ? "#10b981" : "#ef4444"}">
      <div class="kpi-label">累計営業利益（${months[0]}〜${latest}）</div>
      <div class="kpi-value ${totalCumOp < 0 ? "neg-val" : ""}">${fmt(totalCumOp)}</div>
      <div class="kpi-sub">${months.length}ヶ月合計</div>
    </div>
  </div>
  <div class="charts">
    <div class="chart-card">
      <h3>売上高・原価・売上総利益 月次推移</h3>
      <div class="chart-wrap"><canvas id="chart1"></canvas></div>
    </div>
    <div class="chart-card">
      <h3>営業利益 月次推移</h3>
      <div class="chart-wrap"><canvas id="chart2"></canvas></div>
    </div>
  </div>
  <div class="charts" style="grid-template-columns:1fr 1fr; margin-bottom:24px;">
    <div class="chart-card">
      <h3>販管費 月次推移</h3>
      <div class="chart-wrap"><canvas id="chart3"></canvas></div>
    </div>
    <div class="chart-card">
      <h3>粗利率・営業利益率（%）</h3>
      <div class="chart-wrap"><canvas id="chart4"></canvas></div>
    </div>
  </div>
  <div class="table-card">
    <h3>損益明細（万円表示）</h3>
    <table>
      <thead><tr><th>科目</th>${monthHeaders}<th>合計</th></tr></thead>
      <tbody>${tableRows}</tbody>
    </table>
    <div class="updated">最終更新：pl_agent.js で自動生成</div>
  </div>
</div>
<script>
const Ms    = ${JSON.stringify(months)};
const sales = ${JSON.stringify(salesArr)};
const cost  = ${JSON.stringify(costArr)};
const gross = ${JSON.stringify(grossArr)};
const op    = ${JSON.stringify(opArr)};
const exp   = ${JSON.stringify(expArr)};
const grossPct = sales.map((s,i) => s > 0 ? +(gross[i]/s*100).toFixed(1) : null);
const opPct    = sales.map((s,i) => s > 0 ? +(op[i]/s*100).toFixed(1) : null);
const tt = { backgroundColor:'#fff', borderColor:'#e4e9f0', borderWidth:1, titleColor:'#0f172a', bodyColor:'#64748b', padding:10 };
function fmtTip(v) { return v===0?'—':(v<0?'▲':'')+Math.abs(Math.round(v/1e4)).toLocaleString()+'万'; }
new Chart('chart1',{type:'bar',data:{labels:Ms,datasets:[
  {label:'売上高',data:sales,backgroundColor:'rgba(59,130,246,.2)',borderColor:'#3b82f6',borderWidth:2,borderRadius:4,order:1},
  {label:'売上原価',data:cost,backgroundColor:'rgba(139,92,246,.2)',borderColor:'#8b5cf6',borderWidth:2,borderRadius:4,order:1},
  {label:'売上総利益',type:'line',data:gross,borderColor:'#10b981',backgroundColor:'transparent',tension:.3,pointRadius:6,borderWidth:2,yAxisID:'y2',order:0},
]},options:{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{labels:{color:'#94a3b8',padding:12,font:{size:11}}},tooltip:{...tt,callbacks:{label:ctx=>' '+ctx.dataset.label+': '+fmtTip(ctx.raw)}}},
  scales:{x:{ticks:{color:'#94a3b8'},grid:{color:'#e8ecf0'}},y:{ticks:{color:'#94a3b8',callback:v=>Math.round(v/1e8*10)/10+'億'},grid:{color:'#e8ecf0'}},y2:{position:'right',ticks:{color:'#94a3b8',callback:v=>Math.round(v/1e4).toLocaleString()+'万'},grid:{drawOnChartArea:false}}},
}});
new Chart('chart2',{type:'bar',data:{labels:Ms,datasets:[{label:'営業利益',data:op,backgroundColor:op.map(v=>v>=0?'rgba(16,185,129,.3)':'rgba(239,68,68,.3)'),borderColor:op.map(v=>v>=0?'#10b981':'#ef4444'),borderWidth:2,borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#94a3b8'}},tooltip:{...tt,callbacks:{label:ctx=>' 営業利益: '+fmtTip(ctx.raw)}}},scales:{x:{ticks:{color:'#94a3b8'},grid:{color:'#e8ecf0'}},y:{ticks:{color:'#94a3b8',callback:v=>Math.round(v/1e4).toLocaleString()+'万'},grid:{color:'#e8ecf0'}}}}});
new Chart('chart3',{type:'bar',data:{labels:Ms,datasets:[{label:'販管費',data:exp,backgroundColor:'rgba(251,191,36,.25)',borderColor:'#f59e0b',borderWidth:2,borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#94a3b8'}},tooltip:{...tt,callbacks:{label:ctx=>' 販管費: '+fmtTip(ctx.raw)}}},scales:{x:{ticks:{color:'#94a3b8'},grid:{color:'#e8ecf0'}},y:{ticks:{color:'#94a3b8',callback:v=>Math.round(v/1e4).toLocaleString()+'万'},grid:{color:'#e8ecf0'}}}}});
new Chart('chart4',{type:'line',data:{labels:Ms,datasets:[
  {label:'粗利率(%)',data:grossPct,borderColor:'#10b981',backgroundColor:'rgba(16,185,129,.1)',fill:true,tension:.3,pointRadius:5,spanGaps:true},
  {label:'営業利益率(%)',data:opPct,borderColor:'#f59e0b',backgroundColor:'rgba(245,158,11,.1)',fill:true,tension:.3,pointRadius:5,spanGaps:true},
]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#94a3b8',padding:10}},tooltip:{...tt,callbacks:{label:ctx=>ctx.raw!==null?' '+ctx.dataset.label+': '+ctx.raw+'%':'売上ゼロ'}}},scales:{x:{ticks:{color:'#94a3b8'},grid:{color:'#e8ecf0'}},y:{ticks:{color:'#94a3b8',callback:v=>v+'%'},grid:{color:'#e8ecf0'}}}}});
</script>
</body>
</html>`;

  fs.writeFileSync(HTML_PATH, html, "utf-8");
}

// =====================
// ツール定義
// =====================

const TOOLS = [
  {
    name: "read_pl_data",
    description: "現在の損益CSVデータを読み込んで返します。データ確認・分析に使います。",
    input_schema: { type: "object", properties: {}, required: [] },
  },
  {
    name: "update_pl_from_tsv",
    description:
      "エクセルからコピーしたタブ区切りデータを受け取り、損益CSVとHTMLダッシュボードを更新します。\n" +
      "データ形式（1行目はヘッダー）:\n" +
      "  科目[TAB]4月[TAB]5月[TAB]...[TAB]3月[TAB]合計\n" +
      "  製品売上[TAB]1000000[TAB]...",
    input_schema: {
      type: "object",
      properties: {
        tsv_data: { type: "string", description: "エクセルからコピーしたタブ区切りの損益データ" },
      },
      required: ["tsv_data"],
    },
  },
  {
    name: "summarize_pl",
    description: "損益データの主要科目サマリーを返します。月を指定することもできます。",
    input_schema: {
      type: "object",
      properties: {
        target_month: { type: "string", description: "サマリーを表示する月（例: 4月、11月）。省略すると年間合計。" },
      },
      required: [],
    },
  },
];

// =====================
// ツール実行
// =====================

function executeTool(name, input) {
  switch (name) {
    case "read_pl_data": {
      if (!fs.existsSync(CSV_PATH)) return "損益ファイルが見つかりません。";
      return fs.readFileSync(CSV_PATH, "utf-8");
    }

    case "update_pl_from_tsv": {
      const tsv = input.tsv_data;
      let parsed;
      try {
        parsed = parseTsv(tsv);
      } catch (e) {
        return `データ解析エラー: ${e.message}`;
      }

      const { months, data } = parsed;
      calcAggregates(data, months);

      const csvResult = updateCsv(months, data);
      let csvMsg = typeof csvResult === "number"
        ? `CSV更新: ${csvResult} 科目を更新`
        : csvResult;

      try {
        generateHtml(months, data);
      } catch (e) {
        return `${csvMsg}\nHTML生成失敗: ${e.message}`;
      }

      const latest = months[months.length - 1];
      const sales = data["売上高 合計"]?.[latest] ?? 0;
      const gross = data["売上総利益"]?.[latest] ?? 0;
      const op    = data["営業利益"]?.[latest] ?? 0;
      const opRate = sales ? ((op / sales) * 100).toFixed(1) + "%" : "—";

      return [
        "✅ 更新完了",
        `  ${csvMsg}`,
        `  HTML生成: monthly_dashboard.html を更新`,
        "",
        `【${latest} サマリー】`,
        `  売上高    : ${fmt(sales)}`,
        `  売上総利益: ${fmt(gross)}`,
        `  営業利益  : ${fmt(op)}（利益率 ${opRate}）`,
        "",
        "ブラウザで monthly_dashboard.html を再読み込みすると反映されます。",
      ].join("\n");
    }

    case "summarize_pl": {
      if (!fs.existsSync(CSV_PATH)) return "損益ファイルが見つかりません。";
      const content = fs.readFileSync(CSV_PATH, "utf-8");
      const rows = content.split("\n").map((l) => l.split(","));
      const header = rows[0];
      const months = header.slice(1).filter((h) => h && h !== "合計");

      const dataMap = {};
      for (const row of rows.slice(1)) {
        if (row[0]?.trim()) dataMap[row[0].trim()] = row;
      }

      const keyItems = ["売上高 合計", "売上原価 合計", "売上総利益", "販管費 合計", "営業利益", "経常利益"];
      const targetMonth = input.target_month || "";

      const lines = targetMonth && months.includes(targetMonth)
        ? [`【${targetMonth} 損益サマリー】`]
        : ["【年間合計 損益サマリー】"];

      for (const item of keyItems) {
        if (!dataMap[item]) continue;
        const row = dataMap[item];
        let val = 0;
        if (targetMonth && months.includes(targetMonth)) {
          const idx = header.findIndex((h) => h === targetMonth);
          val = parseInt(row[idx] || "0") || 0;
        } else {
          const totalIdx = header.findIndex((h) => h === "合計");
          val = parseInt(row[totalIdx] || "0") || 0;
        }
        lines.push(`  ${item}: ${fmt(val)}`);
      }
      return lines.join("\n");
    }

    default:
      return `ツール '${name}' が見つかりません。`;
  }
}

// =====================
// エージェントループ
// =====================

async function runAgent(userMessage) {
  const messages = [{ role: "user", content: userMessage }];

  const SYSTEM = `あなたはSNK社ダイカスト事業部の月次損益管理AIアシスタントです。

## 主な役割
- エクセルデータの貼り付けを受け取り、CSVとHTMLダッシュボードを更新する
- 損益データの分析・解説・アドバイス

## データ更新
ユーザーがエクセルデータを貼り付けてきたら:
1. update_pl_from_tsv ツールを呼び出してデータを更新する
2. 更新結果と最新月のサマリーを日本語で分かりやすく報告する
3. 気になる点（赤字・利益率低下など）があれば一言コメントする

## データフォーマット
タブ区切り、1行目はヘッダー:
  科目[TAB]4月[TAB]5月[TAB]...[TAB]3月[TAB]合計

## 回答スタイル
- 日本語・簡潔・わかりやすく
- 金額は万円単位
- ネガティブな数値（赤字など）は明確に指摘する`;

  while (true) {
    const response = await client.messages.create({
      model: MODEL,
      max_tokens: 4096,
      system: SYSTEM,
      tools: TOOLS,
      messages,
    });

    if (response.stop_reason === "end_turn") {
      return response.content.find((b) => b.type === "text")?.text ?? "（回答なし）";
    }

    messages.push({ role: "assistant", content: response.content });

    const toolResults = [];
    for (const block of response.content) {
      if (block.type === "tool_use") {
        const result = executeTool(block.name, block.input);
        toolResults.push({ type: "tool_result", tool_use_id: block.id, content: result });
      }
    }
    messages.push({ role: "user", content: toolResults });
  }
}

// =====================
// メインループ（対話型）
// =====================

async function main() {
  console.log("=".repeat(55));
  console.log("  SNK ダイカスト事業部 月次損益 AI エージェント");
  console.log("=".repeat(55));
  console.log("終了: exit または Ctrl+C\n");
  console.log("【使い方】");
  console.log("  ・エクセルのデータをそのまま貼り付けると自動更新");
  console.log("  ・「11月のサマリーを見せて」などの質問もOK");
  console.log("-".repeat(55) + "\n");

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

  const ask = () => {
    rl.question("あなた: ", async (input) => {
      const text = input.trim();
      if (!text) return ask();
      if (["exit", "quit", "終了"].includes(text.toLowerCase())) {
        console.log("終了します。");
        rl.close();
        return;
      }

      process.stdout.write("エージェント: ");
      try {
        const reply = await runAgent(text);
        console.log(reply);
      } catch (e) {
        console.error("エラー:", e.message);
      }
      console.log();
      ask();
    });
  };

  ask();
}

main();
