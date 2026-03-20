"""
SNK ダイカスト事業部 ダッシュボード更新ツール

使い方:
    python update_dashboard.py

エクセルからコピーしたデータを貼り付けると、以下が自動更新されます:
  ① 毎月の損益/ダイカスト事業部損益_2026.csv
  ② 毎月の損益/monthly_dashboard.html（ダッシュボード）

【エクセルの貼り付け形式】
エクセルで「科目・月別の損益表」をそのままコピーして貼り付けてください。

    科目    4月    5月    6月    7月    ...    3月    合計
    製品売上    1000000    1200000    ...
    その他売上    50000    ...
    ...

※ 1行目はヘッダー（科目, 4月, 5月 ...）
※ 合計列は自動で再計算するので空でもOK
"""

import sys
import csv
import json
import re
from pathlib import Path

BASE_DIR = Path(__file__).parent / "毎月の損益"
CSV_PATH = BASE_DIR / "ダイカスト事業部損益_2026.csv"
HTML_PATH = BASE_DIR / "monthly_dashboard.html"

# 売上・利益の集計に使う科目マッピング
CALC_ITEMS = {
    "売上高 合計":    lambda d: d.get("　製品売上", 0) + d.get("　その他売上", 0),
    "売上原価 合計":  lambda d: d.get("　材料費", 0) + d.get("　労務費", 0) + d.get("　製造経費", 0),
    "売上総利益":     lambda d: d.get("売上高 合計", 0) - d.get("売上原価 合計", 0),
    "販管費 合計":    lambda d: sum(d.get(k, 0) for k in ["　人件費","　地代家賃","　減価償却費","　光熱費","　通信費","　その他販管費"]),
    "営業利益":       lambda d: d.get("売上総利益", 0) - d.get("販管費 合計", 0),
    "経常利益":       lambda d: d.get("営業利益", 0) + d.get("　受取利息", 0) + d.get("　その他営業外収益", 0) - d.get("　支払利息", 0) - d.get("　その他営業外費用", 0),
}


# =====================
# データ解析
# =====================

def parse_tsv(text: str):
    """
    タブ区切りのエクセルデータをパース。
    期待フォーマット:
      1行目: 科目 / 4月 / 5月 / ... / 3月 / 合計
      2行目〜: 科目名 / 値 / 値 / ...
    """
    lines = [l for l in text.strip().splitlines() if l.strip()]
    if not lines:
        raise ValueError("データが空です。")

    rows = [line.split('\t') for line in lines]
    header = rows[0]

    # 月リスト（合計列は除外）
    months = []
    month_indices = []
    for i, h in enumerate(header[1:], start=1):
        h_clean = h.strip()
        if h_clean and h_clean != "合計":
            months.append(h_clean)
            month_indices.append(i)

    if not months:
        raise ValueError("月が見つかりません。ヘッダー行に「4月」「5月」などの月名を入れてください。")

    # データ: {科目: {月: 値}}
    data = {}
    for row in rows[1:]:
        if not row[0].strip():
            continue
        item = row[0]  # 科目名（スペースも含めてそのまま保持）
        values = {}
        for i, month in zip(month_indices, months):
            raw = row[i].strip() if i < len(row) else ""
            cleaned = raw.replace(',', '').replace('¥', '').replace('￥', '').replace('\u00a0', '').replace(' ', '')
            try:
                values[month] = int(float(cleaned)) if cleaned else 0
            except ValueError:
                values[month] = 0
        data[item] = values

    return months, data


def calc_aggregates(data: dict, months: list) -> dict:
    """売上高合計・営業利益などの集計値を計算して data に追加"""
    for item, func in CALC_ITEMS.items():
        values = {}
        for m in months:
            month_data = {k: v.get(m, 0) for k, v in data.items()}
            values[m] = func(month_data)
        data[item] = values
    return data


# =====================
# CSV更新
# =====================

def update_csv(months: list, data: dict) -> int:
    """毎月の損益CSVを更新する"""
    with open(CSV_PATH, encoding="utf-8", newline="") as f:
        reader = csv.reader(f)
        rows = list(reader)

    # ヘッダー確認（月の列インデックスを取得）
    header = rows[0]
    month_col = {}
    for i, h in enumerate(header):
        if h.strip() in months:
            month_col[h.strip()] = i

    # 合計列インデックス
    total_col = header.index("合計") if "合計" in header else None

    # 各行を更新
    updated_items: list[str] = []
    for row in rows[1:]:
        if not row[0].strip():
            continue
        item = row[0]
        if item in data:
            row_total = 0
            for m, col_i in month_col.items():
                val = data[item].get(m, 0)
                row[col_i] = str(val)
                row_total += val
            if total_col:
                row[total_col] = str(row_total)
            updated_items.append(item)

    with open(CSV_PATH, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(rows)

    return len(updated_items)


# =====================
# HTML生成
# =====================

def generate_html(months: list, data: dict) -> None:
    """月次損益ダッシュボードHTMLを生成する"""

    def v(item, m):
        return data.get(item, {}).get(m, 0)

    def fmt(n):
        if n == 0:
            return "—"
        sign = "▲" if n < 0 else ""
        return sign + f"{abs(round(n/10000)):,}万"

    def pct(a, b):
        return f"{a/b*100:.1f}%" if b else "—"

    # JavaScriptデータ
    sales_data   = [v("売上高 合計", m) for m in months]
    cost_data    = [v("売上原価 合計", m) for m in months]
    gross_data   = [v("売上総利益", m) for m in months]
    op_data      = [v("営業利益", m) for m in months]
    exp_data     = [v("販管費 合計", m) for m in months]

    latest = months[-1]
    latest_sales = v("売上高 合計", latest)
    latest_op    = v("営業利益", latest)
    latest_gross = v("売上総利益", latest)

    # テーブル行
    table_rows_html = ""
    display_items = [
        ("【売上高】", None),
        ("　製品売上", None),
        ("　その他売上", None),
        ("売上高 合計", "bold"),
        ("【売上原価】", None),
        ("　材料費", None),
        ("　労務費", None),
        ("　製造経費", None),
        ("売上原価 合計", "bold"),
        ("売上総利益", "highlight"),
        ("【販売費及び一般管理費】", None),
        ("　人件費", None),
        ("　地代家賃", None),
        ("　減価償却費", None),
        ("　光熱費", None),
        ("　通信費", None),
        ("　その他販管費", None),
        ("販管費 合計", "bold"),
        ("営業利益", "highlight"),
        ("【営業外収益】", None),
        ("　受取利息", None),
        ("　その他営業外収益", None),
        ("【営業外費用】", None),
        ("　支払利息", None),
        ("　その他営業外費用", None),
        ("経常利益", "highlight"),
    ]

    for item, style in display_items:
        if item.startswith("【"):
            table_rows_html += f'<tr class="section-header"><td colspan="{len(months)+2}">{item}</td></tr>\n'
            continue
        if item not in data:
            continue
        row_total = sum(v(item, m) for m in months)
        cls = ""
        if style == "bold":
            cls = ' class="subtotal"'
        elif style == "highlight":
            cls = ' class="total-row"'
        cells = "".join(
            f'<td class="{"neg" if v(item, m) < 0 else ""}">{fmt(v(item, m))}</td>'
            for m in months
        )
        total_cls = "neg" if row_total < 0 else ""
        table_rows_html += f'<tr{cls}><td class="item-name">{item}</td>{cells}<td class="{total_cls}">{fmt(row_total)}</td></tr>\n'

    months_js    = json.dumps(months, ensure_ascii=False)
    sales_js     = json.dumps(sales_data)
    cost_js      = json.dumps(cost_data)
    gross_js     = json.dumps(gross_data)
    op_js        = json.dumps(op_data)
    exp_js       = json.dumps(exp_data)
    month_headers = "".join(f"<th>{m}</th>" for m in months)

    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>月次損益ダッシュボード｜ダイカスト事業部</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&display=swap');
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Noto Sans JP', sans-serif; background: #f0f4f9; color: #0f172a; }}
  header {{
    background: linear-gradient(135deg, #0f172a, #1e3a6e, #1e40af);
    padding: 0 32px; height: 64px;
    display: flex; align-items: center; justify-content: space-between;
    position: sticky; top: 0; z-index: 100;
    box-shadow: 0 2px 20px rgba(15,23,42,.3);
  }}
  .logo h1 {{ font-size: 15px; font-weight: 700; color: #fff; }}
  .logo span {{ font-size: 11px; color: rgba(255,255,255,.5); display: block; margin-top: 2px; }}
  .header-right {{ color: #fff; font-size: 13px; font-weight: 600; }}
  .header-right small {{ display: block; font-size: 11px; color: rgba(255,255,255,.5); font-weight: 400; }}
  .main {{ max-width: 1400px; margin: 0 auto; padding: 24px; }}
  .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 24px; }}
  .kpi-card {{
    background: #fff; border-radius: 12px; padding: 20px;
    box-shadow: 0 1px 3px rgba(15,23,42,.06), 0 4px 10px rgba(15,23,42,.04);
    border-left: 4px solid var(--c);
  }}
  .kpi-label {{ font-size: 12px; color: #64748b; margin-bottom: 8px; }}
  .kpi-value {{ font-size: 24px; font-weight: 700; color: var(--c); }}
  .kpi-sub {{ font-size: 11px; color: #94a3b8; margin-top: 4px; }}
  .neg-val {{ color: #ef4444 !important; }}
  .charts {{ display: grid; grid-template-columns: 2fr 1fr; gap: 16px; margin-bottom: 24px; }}
  .chart-card {{
    background: #fff; border-radius: 12px; padding: 20px;
    box-shadow: 0 1px 3px rgba(15,23,42,.06), 0 4px 10px rgba(15,23,42,.04);
  }}
  .chart-card h3 {{ font-size: 13px; color: #334155; margin-bottom: 16px; font-weight: 600; }}
  .chart-wrap {{ position: relative; height: 260px; }}
  .table-card {{
    background: #fff; border-radius: 12px; padding: 20px;
    box-shadow: 0 1px 3px rgba(15,23,42,.06), 0 4px 10px rgba(15,23,42,.04);
    overflow-x: auto;
  }}
  .table-card h3 {{ font-size: 13px; color: #334155; margin-bottom: 16px; font-weight: 600; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
  th {{ background: #1e40af; color: #fff; padding: 8px 12px; text-align: right; white-space: nowrap; font-weight: 500; }}
  th:first-child {{ text-align: left; min-width: 160px; }}
  td {{ padding: 7px 12px; text-align: right; border-bottom: 1px solid #e4e9f0; color: #334155; }}
  td.item-name {{ text-align: left; color: #475569; }}
  td.neg {{ color: #ef4444; }}
  tr.section-header td {{ background: #f1f5f9; color: #64748b; font-size: 11px; padding: 5px 12px; font-weight: 600; }}
  tr.subtotal td {{ background: #f8fafc; font-weight: 600; color: #1e40af; }}
  tr.total-row td {{ background: #eff6ff; font-weight: 700; font-size: 13px; }}
  tr.total-row td.neg {{ color: #ef4444; }}
  tr:hover td {{ background: #f8fafc; }}
  tr.section-header:hover td, tr.subtotal:hover td, tr.total-row:hover td {{ background: inherit; }}
  .updated {{ font-size: 11px; color: #94a3b8; margin-top: 12px; text-align: right; }}
</style>
</head>
<body>
<header>
  <div class="logo">
    <h1>月次損益ダッシュボード｜ダイカスト事業部</h1>
    <span>R8年度 事業部単体損益（税抜・円）</span>
  </div>
  <div class="header-right">
    最新月：{latest}
    <small>単位：万円（表示）</small>
  </div>
</header>

<div class="main">

  <!-- KPI -->
  <div class="kpi-grid">
    <div class="kpi-card" style="--c:#3b82f6">
      <div class="kpi-label">売上高（{latest}）</div>
      <div class="kpi-value">{fmt(latest_sales)}</div>
      <div class="kpi-sub">製品売上 + その他売上</div>
    </div>
    <div class="kpi-card" style="--c:{'#10b981' if latest_gross >= 0 else '#ef4444'}">
      <div class="kpi-label">売上総利益（{latest}）</div>
      <div class="kpi-value {'neg-val' if latest_gross < 0 else ''}">{fmt(latest_gross)}</div>
      <div class="kpi-sub">粗利率 {pct(latest_gross, latest_sales)}</div>
    </div>
    <div class="kpi-card" style="--c:{'#f59e0b' if latest_op >= 0 else '#ef4444'}">
      <div class="kpi-label">営業利益（{latest}）</div>
      <div class="kpi-value {'neg-val' if latest_op < 0 else ''}">{fmt(latest_op)}</div>
      <div class="kpi-sub">営業利益率 {pct(latest_op, latest_sales)}</div>
    </div>
    <div class="kpi-card" style="--c:#8b5cf6">
      <div class="kpi-label">累計売上高（{months[0]}〜{latest}）</div>
      <div class="kpi-value">{fmt(sum(sales_data))}</div>
      <div class="kpi-sub">{len(months)}ヶ月合計</div>
    </div>
    <div class="kpi-card" style="--c:{'#10b981' if sum(op_data) >= 0 else '#ef4444'}">
      <div class="kpi-label">累計営業利益（{months[0]}〜{latest}）</div>
      <div class="kpi-value {'neg-val' if sum(op_data) < 0 else ''}">{fmt(sum(op_data))}</div>
      <div class="kpi-sub">{len(months)}ヶ月合計</div>
    </div>
  </div>

  <!-- チャート -->
  <div class="charts">
    <div class="chart-card">
      <h3>売上高・売上原価・売上総利益 月次推移</h3>
      <div class="chart-wrap"><canvas id="chart1"></canvas></div>
    </div>
    <div class="chart-card">
      <h3>営業利益 月次推移</h3>
      <div class="chart-wrap"><canvas id="chart2"></canvas></div>
    </div>
  </div>
  <div class="charts" style="grid-template-columns: 1fr 1fr; margin-bottom: 24px;">
    <div class="chart-card">
      <h3>販管費 月次推移</h3>
      <div class="chart-wrap"><canvas id="chart3"></canvas></div>
    </div>
    <div class="chart-card">
      <h3>粗利率・営業利益率（%）</h3>
      <div class="chart-wrap"><canvas id="chart4"></canvas></div>
    </div>
  </div>

  <!-- テーブル -->
  <div class="table-card">
    <h3>損益明細（万円表示）</h3>
    <table>
      <thead>
        <tr>
          <th>科目</th>
          {month_headers}
          <th>合計</th>
        </tr>
      </thead>
      <tbody>
        {table_rows_html}
      </tbody>
    </table>
    <div class="updated">最終更新：update_dashboard.py で自動生成</div>
  </div>
</div>

<script>
const Ms    = {months_js};
const sales = {sales_js};
const cost  = {cost_js};
const gross = {gross_js};
const op    = {op_js};
const exp   = {exp_js};

const grossPct = sales.map((s,i) => s > 0 ? +(gross[i]/s*100).toFixed(1) : null);
const opPct    = sales.map((s,i) => s > 0 ? +(op[i]/s*100).toFixed(1) : null);

const tt = {{
  backgroundColor:'#fff', borderColor:'#e4e9f0', borderWidth:1,
  titleColor:'#0f172a', bodyColor:'#64748b', padding:10
}};
function fmtTip(v) {{ return v === 0 ? '—' : (v<0?'▲':'')+Math.abs(Math.round(v/1e4)).toLocaleString()+'万'; }}

// Chart 1: 売上・原価・粗利
new Chart('chart1', {{ type:'bar', data:{{
  labels: Ms,
  datasets:[
    {{ label:'売上高', data:sales, backgroundColor:'rgba(59,130,246,.2)', borderColor:'#3b82f6', borderWidth:2, borderRadius:4, order:1 }},
    {{ label:'売上原価', data:cost, backgroundColor:'rgba(139,92,246,.2)', borderColor:'#8b5cf6', borderWidth:2, borderRadius:4, order:1 }},
    {{ label:'売上総利益', type:'line', data:gross, borderColor:'#10b981', backgroundColor:'transparent', tension:.3, pointRadius:6, borderWidth:2, yAxisID:'y2', order:0 }},
  ]
}}, options:{{
  responsive:true, maintainAspectRatio:false,
  plugins:{{ legend:{{ labels:{{ color:'#94a3b8', padding:12, font:{{size:11}} }} }}, tooltip:{{ ...tt, callbacks:{{ label: ctx => ' '+ctx.dataset.label+': '+fmtTip(ctx.raw) }} }} }},
  scales:{{
    x:{{ ticks:{{color:'#94a3b8'}}, grid:{{color:'#e8ecf0'}} }},
    y:{{ ticks:{{ color:'#94a3b8', callback: v => Math.round(v/1e8*10)/10+'億' }}, grid:{{color:'#e8ecf0'}} }},
    y2:{{ position:'right', ticks:{{ color:'#94a3b8', callback: v => Math.round(v/1e4).toLocaleString()+'万' }}, grid:{{drawOnChartArea:false}} }},
  }}
}}  }});

// Chart 2: 営業利益
new Chart('chart2', {{ type:'bar', data:{{
  labels: Ms,
  datasets:[{{ label:'営業利益', data:op, backgroundColor: op.map(v=>v>=0?'rgba(16,185,129,.3)':'rgba(239,68,68,.3)'), borderColor: op.map(v=>v>=0?'#10b981':'#ef4444'), borderWidth:2, borderRadius:4 }}]
}}, options:{{
  responsive:true, maintainAspectRatio:false,
  plugins:{{ legend:{{ labels:{{ color:'#94a3b8' }} }}, tooltip:{{ ...tt, callbacks:{{ label: ctx => ' 営業利益: '+fmtTip(ctx.raw) }} }} }},
  scales:{{ x:{{ ticks:{{color:'#94a3b8'}}, grid:{{color:'#e8ecf0'}} }}, y:{{ ticks:{{ color:'#94a3b8', callback: v => Math.round(v/1e4).toLocaleString()+'万' }}, grid:{{color:'#e8ecf0'}} }} }}
}}  }});

// Chart 3: 販管費
new Chart('chart3', {{ type:'bar', data:{{
  labels: Ms,
  datasets:[{{ label:'販管費', data:exp, backgroundColor:'rgba(251,191,36,.25)', borderColor:'#f59e0b', borderWidth:2, borderRadius:4 }}]
}}, options:{{
  responsive:true, maintainAspectRatio:false,
  plugins:{{ legend:{{ labels:{{ color:'#94a3b8' }} }}, tooltip:{{ ...tt, callbacks:{{ label: ctx => ' 販管費: '+fmtTip(ctx.raw) }} }} }},
  scales:{{ x:{{ ticks:{{color:'#94a3b8'}}, grid:{{color:'#e8ecf0'}} }}, y:{{ ticks:{{ color:'#94a3b8', callback: v => Math.round(v/1e4).toLocaleString()+'万' }}, grid:{{color:'#e8ecf0'}} }} }}
}}  }});

// Chart 4: 利益率
new Chart('chart4', {{ type:'line', data:{{
  labels: Ms,
  datasets:[
    {{ label:'粗利率(%)', data:grossPct, borderColor:'#10b981', backgroundColor:'rgba(16,185,129,.1)', fill:true, tension:.3, pointRadius:5, spanGaps:true }},
    {{ label:'営業利益率(%)', data:opPct, borderColor:'#f59e0b', backgroundColor:'rgba(245,158,11,.1)', fill:true, tension:.3, pointRadius:5, spanGaps:true }},
  ]
}}, options:{{
  responsive:true, maintainAspectRatio:false,
  plugins:{{ legend:{{ labels:{{ color:'#94a3b8', padding:10 }} }}, tooltip:{{ ...tt, callbacks:{{ label: ctx => ctx.raw !== null ? ' '+ctx.dataset.label+': '+ctx.raw+'%' : '売上ゼロ' }} }} }},
  scales:{{ x:{{ ticks:{{color:'#94a3b8'}}, grid:{{color:'#e8ecf0'}} }}, y:{{ ticks:{{ color:'#94a3b8', callback:v=>v+'%' }}, grid:{{color:'#e8ecf0'}} }} }}
}}  }});
</script>
</body>
</html>"""

    HTML_PATH.write_text(html, encoding="utf-8")


# =====================
# メイン
# =====================

def main():
    print("=" * 55)
    print("  SNK ダイカスト事業部 ダッシュボード更新ツール")
    print("=" * 55)
    print()
    print("エクセルの損益表をコピーして貼り付けてください。")
    print()
    print("  Windows: 貼り付け後 → Enter → Ctrl+Z → Enter")
    print("  Mac:     貼り付け後 → Ctrl+D")
    print()
    print("【期待するフォーマット（タブ区切り）】")
    print("  科目    4月    5月    6月    ...    3月    合計")
    print("  製品売上    1000000    1200000    ...")
    print("  材料費    400000    ...")
    print("-" * 55)

    try:
        text = sys.stdin.read()
    except KeyboardInterrupt:
        print("\n中断しました。")
        sys.exit(0)

    if not text.strip():
        print("❌ データが空です。")
        sys.exit(1)

    try:
        months, data = parse_tsv(text)
    except ValueError as e:
        print(f"❌ 解析エラー: {e}")
        sys.exit(1)

    # 集計値を計算
    data = calc_aggregates(data, months)

    # 確認表示
    print()
    print("【解析結果】")
    print(f"  月: {months}")
    print(f"  科目数: {len(data)} 件")

    # 主要KPIをプレビュー
    latest = months[-1]
    def fmt_preview(item):
        val = data.get(item, {}).get(latest, 0)
        return f"{val:,}" if val else "—"

    print(f"\n  {latest} プレビュー:")
    print(f"    売上高合計  : {fmt_preview('売上高 合計')}")
    print(f"    売上原価合計: {fmt_preview('売上原価 合計')}")
    print(f"    売上総利益  : {fmt_preview('売上総利益')}")
    print(f"    営業利益    : {fmt_preview('営業利益')}")
    print()

    ans = input("この内容で更新しますか？ [y/n]: ").strip().lower()
    if ans not in ("y", "yes", "はい", ""):
        print("キャンセルしました。")
        sys.exit(0)

    # CSV更新
    try:
        updated = update_csv(months, data)
        print(f"✅ CSV更新: {updated} 科目を更新（{CSV_PATH.name}）")
    except Exception as e:
        print(f"⚠️ CSV更新エラー: {e}")

    # HTML生成
    try:
        generate_html(months, data)
        print(f"✅ HTML生成: {HTML_PATH.name}")
    except Exception as e:
        print(f"⚠️ HTML生成エラー: {e}")

    print()
    print(f"  期間: {months[0]} 〜 {months[-1]}（{len(months)}ヶ月）")
    print()
    print("ブラウザで monthly_dashboard.html を開くと確認できます。")


if __name__ == "__main__":
    main()
