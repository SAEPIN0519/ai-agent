// No.4 初期流動不良低減 部署長会議スライド生成スクリプト
const pptxgen = require("pptxgenjs");
const path = require("path");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9";
pptx.author = "技術部 技術課";
pptx.subject = "第57期 方針活動 No.4 初期流動不良低減";

const imgDir = path.join(__dirname, "images");

// カラー定義
const C = {
  navy:    "1A3A5C",
  blue:    "2980B9",
  white:   "FFFFFF",
  orange:  "E67E22",
  red:     "E74C3C",
  green:   "27AE60",
  gray:    "7F8C8D",
  dark:    "2C3E50",
  lightBlue: "EAF6FF",
  lightOrange: "FEF9F0",
  lightGreen: "E8F8F0",
  lightRed: "FDEDEC",
};

// ヘッダー共通関数
function addHeader(slide, num, title) {
  slide.background = { fill: C.white };
  slide.addText(String(num), {
    x: 0.5, y: 0.35, w: 0.35, h: 0.35,
    shape: pptx.ShapeType.ellipse, fill: { color: C.blue },
    fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle",
  });
  slide.addText(title, {
    x: 1.0, y: 0.3, w: 7, h: 0.45,
    fontSize: 20, fontFace: "Yu Gothic UI", color: C.navy, bold: true,
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.5, y: 0.85, w: 9, h: 0.04, fill: { color: C.blue },
  });
}

// ─── スライド1: 表紙 ───
const s1 = pptx.addSlide();
s1.background = { fill: C.navy };
s1.addText("熱処理品ブリスター低減活動", {
  x: 0.8, y: 1.2, w: 8.4, h: 1.3,
  fontSize: 34, fontFace: "Yu Gothic UI", color: C.white, bold: true,
});
s1.addText("第57期 方針活動 No.4 初期流動不良低減", {
  x: 0.8, y: 2.7, w: 8.4, h: 0.5,
  fontSize: 14, fontFace: "Yu Gothic UI", color: C.blue,
});
s1.addShape(pptx.ShapeType.rect, {
  x: 0.8, y: 2.55, w: 3.5, h: 0.03, fill: { color: C.blue },
});
s1.addText("対象：シマノ製品 T6（アッパーブラケット）", {
  x: 0.8, y: 3.3, w: 5, h: 0.5,
  fontSize: 16, fontFace: "Yu Gothic UI", color: "AABBCC",
});
s1.addText("新日本金属工業株式会社　技術部 技術課\n担当：後藤・西川・若山・山北\n2026年2月度 部署長会議", {
  x: 0.8, y: 4.1, w: 5, h: 1.0,
  fontSize: 12, fontFace: "Yu Gothic UI", color: "AABBCC", lineSpacingMultiple: 1.5,
});

// ─── スライド2: 背景と目標 ───
const s2 = pptx.addSlide();
addHeader(s2, 1, "活動の背景と目標");

// 背景ボックス
s2.addShape(pptx.ShapeType.roundRect, {
  x: 0.5, y: 1.1, w: 4.2, h: 1.7, fill: { color: C.lightOrange }, rectRadius: 0.1,
});
s2.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 1.1, w: 0.06, h: 1.7, fill: { color: C.orange },
});
s2.addText("背景", {
  x: 0.7, y: 1.15, w: 3.8, h: 0.28,
  fontSize: 12, fontFace: "Yu Gothic UI", color: C.orange, bold: true,
});
s2.addText("・シマノ製品T6熱処理品で\n  ブリスター（膨れ）不良が発生\n・アッパーブラケットの複数箇所で\n  湯じわ・ブリスターを確認", {
  x: 0.7, y: 1.45, w: 3.8, h: 1.2,
  fontSize: 11, fontFace: "Yu Gothic UI", color: C.dark, lineSpacingMultiple: 1.4,
});

// 目標ボックス
s2.addShape(pptx.ShapeType.roundRect, {
  x: 5.3, y: 1.1, w: 4.2, h: 1.7, fill: { color: C.lightBlue }, rectRadius: 0.1,
});
s2.addShape(pptx.ShapeType.rect, {
  x: 5.3, y: 1.1, w: 0.06, h: 1.7, fill: { color: C.blue },
});
s2.addText("目標", {
  x: 5.5, y: 1.15, w: 3.8, h: 0.28,
  fontSize: 12, fontFace: "Yu Gothic UI", color: C.blue, bold: true,
});
s2.addText("・初期流動不良率 3%以下\n・発生率調査と部位傾向調査を実施\n・素材形状の最適化\n・流動解析による金型方案の最適化", {
  x: 5.5, y: 1.45, w: 3.8, h: 1.2,
  fontSize: 11, fontFace: "Yu Gothic UI", color: C.dark, lineSpacingMultiple: 1.4,
});

// 達成度
s2.addText("10〜1月の達成度推移", {
  x: 0.5, y: 3.0, w: 4, h: 0.3,
  fontSize: 12, fontFace: "Yu Gothic UI", color: C.navy, bold: true,
});
const months = ["10月", "11月", "12月", "1月"];
const achieve1 = ["100%", "100%", "100%", "100%"];
const achieve2 = ["100%", "50%", "—", "100%"];
months.forEach((m, i) => {
  const x = 0.5 + i * 1.15;
  s2.addShape(pptx.ShapeType.roundRect, {
    x: x, y: 3.35, w: 1.0, h: 0.95,
    fill: { color: "F8FAFC" }, line: { color: "D5E4F0", width: 1 }, rectRadius: 0.06,
  });
  s2.addText(m, {
    x: x, y: 3.37, w: 1.0, h: 0.25,
    fontSize: 10, fontFace: "Yu Gothic UI", color: C.navy, bold: true, align: "center",
  });
  s2.addText("防振 " + achieve1[i], {
    x: x, y: 3.62, w: 1.0, h: 0.25,
    fontSize: 9, fontFace: "Yu Gothic UI", color: C.green, align: "center",
  });
  s2.addText("T6 " + achieve2[i], {
    x: x, y: 3.85, w: 1.0, h: 0.25,
    fontSize: 9, fontFace: "Yu Gothic UI",
    color: achieve2[i] === "50%" ? C.orange : (achieve2[i] === "—" ? C.gray : C.green),
    align: "center",
  });
});

// 評価項目画像
s2.addImage({
  path: path.join(imgDir, "image25.png"),
  x: 5.3, y: 2.95, w: 4.2, h: 2.1,
});
s2.addText("製品評価項目（硬度・組織・X線・成分）", {
  x: 5.3, y: 5.05, w: 4.2, h: 0.25,
  fontSize: 9, fontFace: "Yu Gothic UI", color: C.gray, align: "center",
});

// ─── スライド3: 不良部位まとめ（写真付き）───
const s3 = pptx.addSlide();
addHeader(s3, 2, "アッパーブラケット 不良箇所まとめ");

// 上面図（メイン画像）
s3.addImage({
  path: path.join(imgDir, "image33.png"),
  x: 0.3, y: 1.05, w: 4.5, h: 3.8,
});
s3.addText("上面図：不良発生箇所（ピンク部）", {
  x: 0.3, y: 4.85, w: 4.5, h: 0.25,
  fontSize: 9, fontFace: "Yu Gothic UI", color: C.gray, align: "center",
});

// 右側に詳細画像3枚+説明
// 足部アップ
s3.addImage({
  path: path.join(imgDir, "image30.png"),
  x: 5.1, y: 1.05, w: 2.1, h: 1.65,
});
s3.addShape(pptx.ShapeType.roundRect, {
  x: 7.3, y: 1.05, w: 2.3, h: 1.65,
  fill: { color: C.lightOrange }, rectRadius: 0.08,
});
s3.addText("⑧⑨ 天側 足部", {
  x: 7.4, y: 1.1, w: 2.1, h: 0.28,
  fontSize: 11, fontFace: "Yu Gothic UI", color: C.orange, bold: true,
});
s3.addText("・湯じわ発生\n・ブリスター発生\n→ 駒割りによる\n  ガス抜きを検討", {
  x: 7.4, y: 1.4, w: 2.1, h: 1.2,
  fontSize: 10, fontFace: "Yu Gothic UI", color: C.dark, lineSpacingMultiple: 1.3,
});

// 側面図
s3.addImage({
  path: path.join(imgDir, "image31.png"),
  x: 5.1, y: 2.85, w: 2.1, h: 0.95,
});
s3.addShape(pptx.ShapeType.roundRect, {
  x: 7.3, y: 2.85, w: 2.3, h: 0.95,
  fill: { color: C.lightBlue }, rectRadius: 0.08,
});
s3.addText("全体側面図", {
  x: 7.4, y: 2.9, w: 2.1, h: 0.25,
  fontSize: 11, fontFace: "Yu Gothic UI", color: C.blue, bold: true,
});
s3.addText("複数箇所で不良確認\n→ 部位別に対策実施", {
  x: 7.4, y: 3.15, w: 2.1, h: 0.55,
  fontSize: 10, fontFace: "Yu Gothic UI", color: C.dark, lineSpacingMultiple: 1.3,
});

// ゲート部分
s3.addImage({
  path: path.join(imgDir, "image32.png"),
  x: 5.1, y: 3.95, w: 2.1, h: 1.15,
});
s3.addShape(pptx.ShapeType.roundRect, {
  x: 7.3, y: 3.95, w: 2.3, h: 1.15,
  fill: { color: C.lightRed }, rectRadius: 0.08,
});
s3.addText("⑩ ゲート周辺", {
  x: 7.4, y: 4.0, w: 2.1, h: 0.25,
  fontSize: 11, fontFace: "Yu Gothic UI", color: C.red, bold: true,
});
s3.addText("・湯じわ発生\n→ デポをあてて\n  湯流れ改善（暫定）", {
  x: 7.4, y: 4.28, w: 2.1, h: 0.75,
  fontSize: 10, fontFace: "Yu Gothic UI", color: C.dark, lineSpacingMultiple: 1.3,
});

// ─── スライド4: 対策検討案 ───
const s4 = pptx.addSlide();
addHeader(s4, 3, "対策検討案（5項目）");

const measures = [
  { no: "1", text: "⑩箇所の湯じわ → デポをあてて湯流れ改善（暫定処置）", status: "暫定対応", color: C.green },
  { no: "2", text: "⑧⑨ 天側足部の湯じわ・ブリスター → 駒割りによるガス抜きを検討中", status: "検討中", color: C.blue },
  { no: "3", text: "メインゲート幅を広くする改善を実施", status: "実施", color: C.green },
  { no: "4", text: "溶体化温度を下げて硬度変化とブリスター変化のTRYを実施", status: "3月末〜", color: C.orange },
  { no: "5", text: "熱処理網の入れ方を再検討し、トライアルパターンを決めて実施", status: "3月末〜", color: C.orange },
];

measures.forEach((item, i) => {
  const y = 1.05 + i * 0.65;
  const bgMap = { [C.green]: C.lightGreen, [C.blue]: C.lightBlue, [C.orange]: C.lightOrange };
  const bg = bgMap[item.color] || C.lightBlue;

  s4.addShape(pptx.ShapeType.roundRect, {
    x: 0.5, y: y, w: 9, h: 0.55, fill: { color: bg }, rectRadius: 0.08,
  });
  s4.addShape(pptx.ShapeType.rect, {
    x: 0.5, y: y, w: 0.06, h: 0.55, fill: { color: item.color },
  });
  s4.addShape(pptx.ShapeType.ellipse, {
    x: 0.7, y: y + 0.09, w: 0.36, h: 0.36, fill: { color: item.color },
  });
  s4.addText(item.no, {
    x: 0.7, y: y + 0.09, w: 0.36, h: 0.36,
    fontSize: 12, fontFace: "Yu Gothic UI", color: C.white, bold: true, align: "center", valign: "middle",
  });
  s4.addText(item.text, {
    x: 1.2, y: y + 0.02, w: 6.3, h: 0.5,
    fontSize: 11, fontFace: "Yu Gothic UI", color: C.dark, valign: "middle",
  });
  s4.addShape(pptx.ShapeType.roundRect, {
    x: 8.2, y: y + 0.1, w: 1.1, h: 0.34, fill: { color: item.color }, rectRadius: 0.17,
  });
  s4.addText(item.status, {
    x: 8.2, y: y + 0.1, w: 1.1, h: 0.34,
    fontSize: 9, fontFace: "Yu Gothic UI", color: C.white, bold: true, align: "center", valign: "middle",
  });
});

// 改善確認トライ
s4.addShape(pptx.ShapeType.roundRect, {
  x: 0.5, y: 4.45, w: 9, h: 0.45,
  fill: { color: C.red }, rectRadius: 0.08,
});
s4.addText("3/23〜3/25　改善確認トライ実施 → 対策【3】【4】【5】の効果を一括検証", {
  x: 0.5, y: 4.45, w: 9, h: 0.45,
  fontSize: 13, fontFace: "Yu Gothic UI", color: C.white, bold: true, align: "center", valign: "middle",
});

// ─── スライド5: 改善トライ詳細＋担当 ───
const s5 = pptx.addSlide();
addHeader(s5, 4, "改善確認トライ（3/23〜25）検証項目と担当");

// 検証4項目
const verifyItems = [
  { no: "①", text: "硬度分布" },
  { no: "②", text: "溶体化温度\n2パターン比較" },
  { no: "③", text: "ブリスター\n発生有無" },
  { no: "④", text: "変形量の計測" },
];
verifyItems.forEach((item, i) => {
  const x = 0.5 + i * 2.35;
  s5.addShape(pptx.ShapeType.roundRect, {
    x: x, y: 1.05, w: 2.1, h: 0.85,
    fill: { color: C.lightBlue }, line: { color: C.blue, width: 1 }, rectRadius: 0.08,
  });
  s5.addText(item.no, {
    x: x, y: 1.08, w: 2.1, h: 0.25,
    fontSize: 14, fontFace: "Yu Gothic UI", color: C.blue, bold: true, align: "center",
  });
  s5.addText(item.text, {
    x: x + 0.1, y: 1.35, w: 1.9, h: 0.48,
    fontSize: 10, fontFace: "Yu Gothic UI", color: C.dark, align: "center", valign: "middle", lineSpacingMultiple: 1.2,
  });
});

// 担当テーブル
s5.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 2.15, w: 9, h: 0.38, fill: { color: C.navy },
});
s5.addText("タスク", {
  x: 0.5, y: 2.15, w: 4.8, h: 0.38,
  fontSize: 10, fontFace: "Yu Gothic UI", color: C.white, bold: true, align: "center", valign: "middle",
});
s5.addText("担当", {
  x: 5.3, y: 2.15, w: 2.2, h: 0.38,
  fontSize: 10, fontFace: "Yu Gothic UI", color: C.white, bold: true, align: "center", valign: "middle",
});
s5.addText("期限", {
  x: 7.5, y: 2.15, w: 2, h: 0.38,
  fontSize: 10, fontFace: "Yu Gothic UI", color: C.white, bold: true, align: "center", valign: "middle",
});

const tasks = [
  { task: "① 溶体化温度（SNK基準）調査→課員配布", who: "粟野", when: "3/25 済" },
  { task: "② 成分分析（Mg量確認・補充調整）", who: "山田→品管", when: "3/23 AM" },
  { task: "③ UPPER 熱処理前（シム挟みN=30）", who: "若山", when: "3/28" },
  { task: "③ UPPER 熱処理後（対角シム変形確認）", who: "若山", when: "4/1" },
  { task: "④ LOWER NG返却品 段差測定（各10個）", who: "山北", when: "4/4" },
  { task: "④ LOWER X線調査（同ロット）", who: "山北", when: "4/7" },
  { task: "④ LOWER 新規打ち 熱処理後（各15個）", who: "山北", when: "4/11" },
  { task: "⑤ 着座NG仕分け保管・発生数管理", who: "横山→春近", when: "4/15〜" },
];

tasks.forEach((item, i) => {
  const y = 2.53 + i * 0.32;
  const bg = i % 2 === 0 ? "F8FAFC" : C.white;
  const isDone = item.when.includes("済");

  s5.addShape(pptx.ShapeType.rect, {
    x: 0.5, y: y, w: 9, h: 0.32, fill: { color: bg },
  });
  s5.addText(item.task, {
    x: 0.6, y: y, w: 4.6, h: 0.32,
    fontSize: 9, fontFace: "Yu Gothic UI", color: C.dark, valign: "middle",
  });
  s5.addText(item.who, {
    x: 5.3, y: y, w: 2.2, h: 0.32,
    fontSize: 9, fontFace: "Yu Gothic UI", color: C.dark, align: "center", valign: "middle",
  });
  s5.addText(item.when, {
    x: 7.5, y: y, w: 2, h: 0.32,
    fontSize: 9, fontFace: "Yu Gothic UI", color: isDone ? C.green : C.navy,
    bold: isDone, align: "center", valign: "middle",
  });
});

// Mg補充の補足
s5.addShape(pptx.ShapeType.roundRect, {
  x: 0.5, y: 4.65, w: 9, h: 0.5,
  fill: { color: C.lightOrange }, line: { color: C.orange, width: 1 }, rectRadius: 0.08,
});
s5.addText("補足：溶体化温度を下げても硬度が出るよう、事前にMg量を確認・補充してから鋳造する", {
  x: 0.7, y: 4.65, w: 8.6, h: 0.5,
  fontSize: 11, fontFace: "Yu Gothic UI", color: C.dark, valign: "middle",
});

// ─── スライド6: スケジュールまとめ ───
const s6 = pptx.addSlide();
addHeader(s6, 5, "今後のスケジュール");

const timeline = [
  { period: "3/23〜25", title: "改善確認\nトライ", desc: "ゲート幅拡大品で\n熱処理トライ実施\n溶体化温度2パターン", color: C.red, active: true },
  { period: "3/25〜4/1", title: "UPPER\n調査", desc: "熱処理前後の\nシム挟み・変形確認\nN=30で検証", color: C.blue, active: false },
  { period: "4/4〜4/11", title: "LOWER\n調査", desc: "NG返却品 段差測定\nX線調査\n新規打ち確認", color: C.blue, active: false },
  { period: "4/15〜", title: "着座NG\n調査", desc: "春近さんへ依頼\nNG仕分け保管\n発生数の管理", color: C.blue, active: false },
];

timeline.forEach((item, i) => {
  const x = 0.5 + i * 2.35;
  const bgColor = item.active ? C.lightRed : C.lightBlue;
  const borderColor = item.active ? C.red : "D5E4F0";

  s6.addShape(pptx.ShapeType.roundRect, {
    x: x, y: 1.05, w: 2.1, h: 2.4,
    fill: { color: bgColor }, line: { color: borderColor, width: 2 }, rectRadius: 0.1,
  });
  s6.addShape(pptx.ShapeType.roundRect, {
    x: x + 0.3, y: 1.18, w: 1.5, h: 0.28,
    fill: { color: item.color }, rectRadius: 0.14,
  });
  s6.addText(item.period, {
    x: x + 0.3, y: 1.18, w: 1.5, h: 0.28,
    fontSize: 9, fontFace: "Yu Gothic UI", color: C.white, bold: true, align: "center", valign: "middle",
  });
  s6.addText(item.title, {
    x: x + 0.1, y: 1.55, w: 1.9, h: 0.5,
    fontSize: 13, fontFace: "Yu Gothic UI", color: C.navy, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.1,
  });
  s6.addText(item.desc, {
    x: x + 0.1, y: 2.1, w: 1.9, h: 1.0,
    fontSize: 10, fontFace: "Yu Gothic UI", color: C.dark, align: "center", lineSpacingMultiple: 1.3,
  });
});

// ポイント
s6.addShape(pptx.ShapeType.roundRect, {
  x: 0.5, y: 3.65, w: 9, h: 1.4,
  fill: { color: C.lightGreen }, line: { color: C.green, width: 2 }, rectRadius: 0.1,
});
s6.addText("ポイント", {
  x: 0.7, y: 3.7, w: 2, h: 0.3,
  fontSize: 12, fontFace: "Yu Gothic UI", color: C.green, bold: true,
});
s6.addText(
  "・溶体化温度を下げても硬度が出るよう、事前にMg量を確認・補充してから鋳造\n" +
  "・熱処理前後の変形量・ブリスター発生有無を定量的に比較し、最適条件を検証\n" +
  "・UPPER / LOWER それぞれで調査を進め、4月中に全体の方向性を確定",
  {
    x: 0.7, y: 4.0, w: 8.6, h: 0.9,
    fontSize: 11, fontFace: "Yu Gothic UI", color: C.dark, lineSpacingMultiple: 1.5,
  }
);

// ─── 保存 ───
const outPath = path.join(__dirname, "初期流動不良低減_部署長会議_v2.pptx");
pptx.writeFile({ fileName: outPath }).then(() => {
  console.log("生成完了: " + outPath);
}).catch(err => {
  console.error("エラー:", err);
});
