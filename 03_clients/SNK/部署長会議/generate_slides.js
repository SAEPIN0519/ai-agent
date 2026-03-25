// 図面管理システム 部署長会議スライド生成スクリプト
const pptxgen = require("pptxgenjs");
const path = require("path");

const pptx = new pptxgen();

// マスター設定
pptx.layout = "LAYOUT_16x9";
pptx.author = "技術部 技術課";
pptx.subject = "第57期 方針活動 No.5 AI類似図面検索システム導入検討";

// カラー定義
const C = {
  navy:    "1A3A5C",
  blue:    "2980B9",
  white:   "FFFFFF",
  light:   "F5F7FA",
  orange:  "E67E22",
  red:     "E74C3C",
  green:   "27AE60",
  gray:    "7F8C8D",
  dark:    "2C3E50",
  lightBlue: "EAF6FF",
  lightOrange: "FEF9F0",
  lightGreen: "E8F8F0",
};

// ─── スライド1: 表紙 ───
const s1 = pptx.addSlide();
s1.background = { fill: C.navy };

s1.addText("AI類似図面検索システム\n導入検討", {
  x: 0.8, y: 1.2, w: 8.4, h: 2.2,
  fontSize: 36, fontFace: "Yu Gothic UI",
  color: C.white, bold: true, lineSpacingMultiple: 1.3,
});
s1.addText("第57期 方針活動 No.5 技術部環境整備 ― DX推進計画③", {
  x: 0.8, y: 3.5, w: 8.4, h: 0.5,
  fontSize: 14, fontFace: "Yu Gothic UI", color: C.blue,
});
s1.addShape(pptx.ShapeType.rect, {
  x: 0.8, y: 3.3, w: 3.5, h: 0.03, fill: { color: C.blue },
});
s1.addText("新日本金属工業株式会社　技術部 技術課\n担当：粟野・津田\n2026年2月度 部署長会議", {
  x: 0.8, y: 4.3, w: 5, h: 1.2,
  fontSize: 13, fontFace: "Yu Gothic UI", color: "AABBCC", lineSpacingMultiple: 1.5,
});

// ─── スライド2: 現状の課題 ───
const s2 = pptx.addSlide();
s2.background = { fill: C.white };

s2.addText([
  { text: "1", options: { fontSize: 14, bold: true, color: C.white } },
], {
  x: 0.5, y: 0.35, w: 0.35, h: 0.35,
  shape: pptx.ShapeType.ellipse, fill: { color: C.blue }, align: "center", valign: "middle",
});
s2.addText("現状の課題", {
  x: 1.0, y: 0.3, w: 4, h: 0.45,
  fontSize: 22, fontFace: "Yu Gothic UI", color: C.navy, bold: true,
});
s2.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 0.85, w: 9, h: 0.04, fill: { color: C.blue },
});

const issues = [
  { label: "課題①", text: "製品種類の増加により図面数が増大し、\n目的の図面にたどり着けない" },
  { label: "課題②", text: "類似製品の検索に時間がかかる\n（手作業で過去図面を1枚ずつ確認）" },
  { label: "課題③", text: "過去の対策品・トラブル履歴の\n横展開が遅い" },
  { label: "課題④", text: "検索ノウハウが設計者の経験値に依存し、\n無駄な再設計が発生" },
];

issues.forEach((item, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.5 + col * 4.7;
  const y = 1.2 + row * 1.6;

  s2.addShape(pptx.ShapeType.rect, {
    x: x, y: y, w: 4.3, h: 1.3,
    fill: { color: C.lightOrange },
    rectRadius: 0.1,
  });
  s2.addShape(pptx.ShapeType.rect, {
    x: x, y: y, w: 0.06, h: 1.3, fill: { color: C.orange },
  });
  s2.addText(item.label, {
    x: x + 0.2, y: y + 0.1, w: 3.8, h: 0.3,
    fontSize: 11, fontFace: "Yu Gothic UI", color: C.orange, bold: true,
  });
  s2.addText(item.text, {
    x: x + 0.2, y: y + 0.4, w: 3.8, h: 0.8,
    fontSize: 13, fontFace: "Yu Gothic UI", color: C.dark, lineSpacingMultiple: 1.3,
  });
});

// 本質的課題
s2.addShape(pptx.ShapeType.roundRect, {
  x: 1.5, y: 4.6, w: 7, h: 0.7,
  fill: { color: C.red }, rectRadius: 0.1,
});
s2.addText('本質的課題：「過去の設計資産が活かせていない」', {
  x: 1.5, y: 4.6, w: 7, h: 0.7,
  fontSize: 18, fontFace: "Yu Gothic UI", color: C.white, bold: true, align: "center", valign: "middle",
});

// ─── スライド3: 目指す姿 ───
const s3 = pptx.addSlide();
s3.background = { fill: C.white };

s3.addText([
  { text: "2", options: { fontSize: 14, bold: true, color: C.white } },
], {
  x: 0.5, y: 0.35, w: 0.35, h: 0.35,
  shape: pptx.ShapeType.ellipse, fill: { color: C.blue }, align: "center", valign: "middle",
});
s3.addText("目指す姿（To-Be）", {
  x: 1.0, y: 0.3, w: 5, h: 0.45,
  fontSize: 22, fontFace: "Yu Gothic UI", color: C.navy, bold: true,
});
s3.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 0.85, w: 9, h: 0.04, fill: { color: C.blue },
});

const tobes = [
  { title: "画像アップロードで即検索", desc: "図面画像をアップロードするだけで\n類似図面を自動抽出" },
  { title: "OCR＋形状のハイブリッド検索", desc: "図面の寸法・注記文字をすべて\nOCR読取り。注記の文言検索も可能" },
  { title: "トラブル履歴と紐付け", desc: "過去のトラブル対策履歴と紐付け、\n再発防止を即時化" },
  { title: "設計横展開の即時化", desc: "無駄な再設計をゼロへ\n過去資産を最大限に活用" },
];

tobes.forEach((item, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.5 + col * 4.7;
  const y = 1.2 + row * 1.8;

  s3.addShape(pptx.ShapeType.rect, {
    x: x, y: y, w: 4.3, h: 1.5,
    fill: { color: C.lightBlue },
    rectRadius: 0.1,
  });
  s3.addShape(pptx.ShapeType.rect, {
    x: x, y: y, w: 0.06, h: 1.5, fill: { color: C.blue },
  });
  s3.addText(item.title, {
    x: x + 0.2, y: y + 0.15, w: 3.8, h: 0.35,
    fontSize: 14, fontFace: "Yu Gothic UI", color: C.blue, bold: true,
  });
  s3.addText(item.desc, {
    x: x + 0.2, y: y + 0.55, w: 3.8, h: 0.8,
    fontSize: 13, fontFace: "Yu Gothic UI", color: C.dark, lineSpacingMultiple: 1.4,
  });
});

// ─── スライド4: 解決できる問題 ───
const s4 = pptx.addSlide();
s4.background = { fill: C.white };

s4.addText([
  { text: "3", options: { fontSize: 14, bold: true, color: C.white } },
], {
  x: 0.5, y: 0.35, w: 0.35, h: 0.35,
  shape: pptx.ShapeType.ellipse, fill: { color: C.blue }, align: "center", valign: "middle",
});
s4.addText("システム機能と解決できる問題", {
  x: 1.0, y: 0.3, w: 6, h: 0.45,
  fontSize: 22, fontFace: "Yu Gothic UI", color: C.navy, bold: true,
});
s4.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 0.85, w: 9, h: 0.04, fill: { color: C.blue },
});

// 機能タグ
const funcs = ["図面をAI画像解析", "形状類似度を自動判定", "OCRで寸法・注記を全文検索", "図面ファイル出力", "付帯情報入力"];
funcs.forEach((f, i) => {
  s4.addShape(pptx.ShapeType.roundRect, {
    x: 0.5 + i * 1.85, y: 1.15, w: 1.7, h: 0.45,
    fill: { color: C.blue }, rectRadius: 0.2,
  });
  s4.addText(f, {
    x: 0.5 + i * 1.85, y: 1.15, w: 1.7, h: 0.45,
    fontSize: 9, fontFace: "Yu Gothic UI", color: C.white, bold: true, align: "center", valign: "middle",
  });
});

// 解決テーブル
const solves = [
  { before: "類似図面の検索が困難", after: "形状AIで瞬時に検索" },
  { before: "横展開の遅延", after: "類似事例を即時抽出" },
  { before: "設計者の経験値に依存", after: "データベース化で属人性を排除" },
  { before: "無駄な再設計が発生", after: "過去図面の流用を促進" },
];

// テーブルヘッダー
s4.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 1.9, w: 9, h: 0.5, fill: { color: C.navy },
});
s4.addText("現状の課題", {
  x: 0.5, y: 1.9, w: 3.8, h: 0.5,
  fontSize: 13, fontFace: "Yu Gothic UI", color: C.white, bold: true, align: "center", valign: "middle",
});
s4.addText("", {
  x: 4.3, y: 1.9, w: 1.4, h: 0.5,
  fontSize: 13, align: "center", valign: "middle",
});
s4.addText("導入後の解決", {
  x: 5.7, y: 1.9, w: 3.8, h: 0.5,
  fontSize: 13, fontFace: "Yu Gothic UI", color: C.white, bold: true, align: "center", valign: "middle",
});

solves.forEach((item, i) => {
  const y = 2.4 + i * 0.6;
  const bg = i % 2 === 0 ? "F8FAFC" : C.white;

  s4.addShape(pptx.ShapeType.rect, {
    x: 0.5, y: y, w: 9, h: 0.55, fill: { color: bg },
  });
  s4.addText(item.before, {
    x: 0.7, y: y, w: 3.4, h: 0.55,
    fontSize: 13, fontFace: "Yu Gothic UI", color: C.dark, valign: "middle",
  });
  s4.addText("→", {
    x: 4.3, y: y, w: 1.4, h: 0.55,
    fontSize: 20, fontFace: "Yu Gothic UI", color: C.blue, bold: true, align: "center", valign: "middle",
  });
  s4.addText(item.after, {
    x: 5.9, y: y, w: 3.4, h: 0.55,
    fontSize: 13, fontFace: "Yu Gothic UI", color: C.navy, bold: true, valign: "middle",
  });
});

// ─── スライド5: DX実行ステップ ───
const s5 = pptx.addSlide();
s5.background = { fill: C.white };

s5.addText([
  { text: "4", options: { fontSize: 14, bold: true, color: C.white } },
], {
  x: 0.5, y: 0.35, w: 0.35, h: 0.35,
  shape: pptx.ShapeType.ellipse, fill: { color: C.blue }, align: "center", valign: "middle",
});
s5.addText("DX実行ステップ", {
  x: 1.0, y: 0.3, w: 5, h: 0.45,
  fontSize: 22, fontFace: "Yu Gothic UI", color: C.navy, bold: true,
});
s5.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 0.85, w: 9, h: 0.04, fill: { color: C.blue },
});

const steps = [
  { num: "STEP 1", title: "デモ実施済", period: "3/19 完了", items: "・システム部 後藤本部長\n・技術 橋爪部長\n・後藤にて\n  デモンストレーション\n  トライアル実施", active: true },
  { num: "STEP 2", title: "稟議", period: "4月中", items: "・導入稟議の提出\n・投資対効果の算出\n・承認取得", active: false },
  { num: "STEP 3", title: "環境準備", period: "5月中", items: "・サーバー準備\n・システム会社にて\n  全図面データを\n  AIに学習・読取り", active: false },
  { num: "STEP 4", title: "実装", period: "6月〜7月", items: "・本番環境への実装\n・全図面の登録\n・運用開始", active: false },
];

steps.forEach((step, i) => {
  const x = 0.5 + i * 2.35;
  const borderColor = step.active ? C.blue : "D5E4F0";
  const bgColor = step.active ? C.lightBlue : C.white;

  s5.addShape(pptx.ShapeType.roundRect, {
    x: x, y: 1.15, w: 2.1, h: 3.0,
    fill: { color: bgColor },
    line: { color: borderColor, width: 2 },
    rectRadius: 0.12,
  });

  // ステップ番号バッジ
  const badgeColor = step.active ? C.red : C.blue;
  s5.addShape(pptx.ShapeType.roundRect, {
    x: x + 0.45, y: 1.3, w: 1.2, h: 0.3,
    fill: { color: badgeColor }, rectRadius: 0.15,
  });
  s5.addText(step.active ? step.num + "（現在）" : step.num, {
    x: x + 0.45, y: 1.3, w: 1.2, h: 0.3,
    fontSize: 9, fontFace: "Yu Gothic UI", color: C.white, bold: true, align: "center", valign: "middle",
  });

  s5.addText(step.title, {
    x: x + 0.15, y: 1.7, w: 1.8, h: 0.4,
    fontSize: 15, fontFace: "Yu Gothic UI", color: C.navy, bold: true, align: "center", valign: "middle",
  });
  s5.addText(step.items, {
    x: x + 0.15, y: 2.15, w: 1.8, h: 1.4,
    fontSize: 10, fontFace: "Yu Gothic UI", color: C.dark, lineSpacingMultiple: 1.4,
  });
  s5.addText(step.period, {
    x: x + 0.15, y: 3.6, w: 1.8, h: 0.35,
    fontSize: 10, fontFace: "Yu Gothic UI", color: C.blue, bold: true, align: "center", valign: "middle",
  });
});

// 3/19トライアルボックス
s5.addShape(pptx.ShapeType.roundRect, {
  x: 0.5, y: 4.35, w: 9, h: 0.8,
  fill: { color: C.lightGreen },
  line: { color: C.green, width: 2 },
  rectRadius: 0.1,
});
s5.addText("3/19 デモンストレーショントライアル実施済", {
  x: 0.7, y: 4.35, w: 5, h: 0.35,
  fontSize: 12, fontFace: "Yu Gothic UI", color: C.green, bold: true,
});
s5.addText("✔ システム部 後藤本部長・技術 橋爪部長・後藤にてデモ実施　　✔ 4月中に稟議 → 5月サーバー準備・AI学習 → 6〜7月実装予定", {
  x: 0.7, y: 4.7, w: 8.5, h: 0.35,
  fontSize: 11, fontFace: "Yu Gothic UI", color: C.dark,
});

// ─── スライド6: 費用対効果 ───
const s6 = pptx.addSlide();
s6.background = { fill: C.white };

s6.addText([
  { text: "5", options: { fontSize: 14, bold: true, color: C.white } },
], {
  x: 0.5, y: 0.35, w: 0.35, h: 0.35,
  shape: pptx.ShapeType.ellipse, fill: { color: C.blue }, align: "center", valign: "middle",
});
s6.addText("費用対効果（想定）", {
  x: 1.0, y: 0.3, w: 5, h: 0.45,
  fontSize: 22, fontFace: "Yu Gothic UI", color: C.navy, bold: true,
});
s6.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 0.85, w: 9, h: 0.04, fill: { color: C.blue },
});

// 現状コストボックス
s6.addShape(pptx.ShapeType.roundRect, {
  x: 0.5, y: 1.3, w: 3.5, h: 2.8,
  fill: { color: "F8FAFC" },
  line: { color: "D5E4F0", width: 2 },
  rectRadius: 0.12,
});
s6.addText("現状の検索コスト（年間）", {
  x: 0.5, y: 1.4, w: 3.5, h: 0.4,
  fontSize: 11, fontFace: "Yu Gothic UI", color: C.gray, align: "center",
});
s6.addText("約137万円", {
  x: 0.5, y: 1.9, w: 3.5, h: 0.7,
  fontSize: 36, fontFace: "Yu Gothic UI", color: C.navy, bold: true, align: "center",
});
s6.addText("設計者13名 × 月5時間 × 12ヶ月\n= 年間780時間 × 時給1,750円", {
  x: 0.5, y: 2.7, w: 3.5, h: 0.8,
  fontSize: 11, fontFace: "Yu Gothic UI", color: C.gray, align: "center", lineSpacingMultiple: 1.5,
});

// 矢印
s6.addText("→", {
  x: 4.0, y: 2.0, w: 2, h: 0.8,
  fontSize: 40, fontFace: "Yu Gothic UI", color: C.blue, bold: true, align: "center", valign: "middle",
});

// 効果ボックス
s6.addShape(pptx.ShapeType.roundRect, {
  x: 6.0, y: 1.3, w: 3.5, h: 2.8,
  fill: { color: C.blue },
  rectRadius: 0.12,
});
s6.addText("70%削減で見込める年間効果", {
  x: 6.0, y: 1.4, w: 3.5, h: 0.4,
  fontSize: 11, fontFace: "Yu Gothic UI", color: "AACCEE", align: "center",
});
s6.addText("約96万円", {
  x: 6.0, y: 1.9, w: 3.5, h: 0.7,
  fontSize: 36, fontFace: "Yu Gothic UI", color: C.white, bold: true, align: "center",
});
s6.addText("検索時間の半減による人件費削減\n＋品質向上・横展開加速の副次効果", {
  x: 6.0, y: 2.7, w: 3.5, h: 0.8,
  fontSize: 11, fontFace: "Yu Gothic UI", color: "AACCEE", align: "center", lineSpacingMultiple: 1.5,
});

// フッターまとめ
s6.addShape(pptx.ShapeType.roundRect, {
  x: 0.5, y: 4.4, w: 9, h: 0.6,
  fill: { color: C.lightBlue },
  rectRadius: 0.08,
});
s6.addText("3/19 デモ実施済 → 4月中に稟議提出 → 5月サーバー準備・AI学習 → 6〜7月 実装予定", {
  x: 0.5, y: 4.4, w: 9, h: 0.6,
  fontSize: 14, fontFace: "Yu Gothic UI", color: C.navy, bold: true, align: "center", valign: "middle",
});

// ─── 保存 ───
const outPath = path.join(__dirname, "図面管理システム_部署長会議_v6.pptx");
pptx.writeFile({ fileName: outPath }).then(() => {
  console.log("生成完了: " + outPath);
}).catch(err => {
  console.error("エラー:", err);
});
