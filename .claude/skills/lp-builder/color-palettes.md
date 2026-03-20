# カラーパレット集（CSS変数形式）

## 目次
- [モダン・テック](#modern)
- [温かみ・ナチュラル](#warm)
- [クール・プレミアム](#cool)
- [エレガント・ラグジュアリー](#elegant)
- [ポップ・クリエイティブ](#pop)
- [ダーク・シック](#dark)

---

## モダン・テック {#modern}

**用途：** SaaS、テクノロジー、スタートアップ、BtoB、AI系

```css
:root {
  --primary: #3D5AFE;       /* インディゴブルー */
  --primary-dark: #1939B7;
  --accent: #00BCD4;        /* シアン */
  --bg: #F8FAFF;
  --bg-dark: #EEF2FF;
  --white: #FFFFFF;
  --text: #1A1A2E;
  --text-muted: #5C6BC0;
  --border: rgba(61, 90, 254, 0.12);
  --shadow: 0 4px 24px rgba(61, 90, 254, 0.1);
  --radius: 16px;
  --radius-sm: 8px;
}
/* Hero背景グラデーション */
/* background: linear-gradient(135deg, #1A1A2E 0%, #16213E 50%, #3D5AFE 100%); */
```

---

## 温かみ・ナチュラル {#warm}

**用途：** 個人ブランド、ライフスタイル、ハンドメイド、教育、コーチング

```css
:root {
  --primary: #E07B39;       /* テラコッタオレンジ */
  --primary-dark: #B85C1A;
  --accent: #6D9B4A;        /* アースグリーン */
  --bg: #FDF8F3;
  --bg-dark: #F5EDE3;
  --white: #FFFFFF;
  --text: #2C1810;
  --text-muted: #7D5A50;
  --border: rgba(224, 123, 57, 0.15);
  --shadow: 0 4px 24px rgba(224, 123, 57, 0.1);
  --radius: 20px;
  --radius-sm: 10px;
}
/* Hero背景グラデーション */
/* background: linear-gradient(135deg, #2C1810 0%, #5D3A1A 50%, #E07B39 100%); */
```

---

## クール・プレミアム {#cool}

**用途：** コンサル、士業、金融、医療、信頼感が重要なBtoB

```css
:root {
  --primary: #1565C0;       /* ロイヤルブルー */
  --primary-dark: #0D47A1;
  --accent: #0097A7;        /* ティール */
  --bg: #F5F7FA;
  --bg-dark: #E8EDF5;
  --white: #FFFFFF;
  --text: #0D1B2A;
  --text-muted: #546E7A;
  --border: rgba(21, 101, 192, 0.12);
  --shadow: 0 4px 24px rgba(21, 101, 192, 0.1);
  --radius: 12px;
  --radius-sm: 6px;
}
/* Hero背景グラデーション */
/* background: linear-gradient(135deg, #0D1B2A 0%, #1565C0 100%); */
```

---

## エレガント・ラグジュアリー {#elegant}

**用途：** ブライダル、ファッション、美容、高級感が必要なBtoC

```css
:root {
  --primary: #8D6748;       /* ウォームブラウン */
  --primary-dark: #6D4C41;
  --accent: #C9A96E;        /* ゴールド */
  --bg: #FAF7F4;
  --bg-dark: #F0EBE3;
  --white: #FFFFFF;
  --text: #1C1208;
  --text-muted: #8D6748;
  --border: rgba(201, 169, 110, 0.2);
  --shadow: 0 4px 24px rgba(201, 169, 110, 0.15);
  --radius: 4px;            /* エレガントは角丸を小さく */
  --radius-sm: 2px;
}
/* Hero背景グラデーション */
/* background: linear-gradient(135deg, #1C1208 0%, #3D2314 50%, #8D6748 100%); */
```

---

## ポップ・クリエイティブ {#pop}

**用途：** イベント、若年層向けサービス、エンタメ、クリエイター

```css
:root {
  --primary: #7C4DFF;       /* バイオレット */
  --primary-dark: #6200EA;
  --accent: #FF4081;        /* ピンク */
  --bg: #F9F5FF;
  --bg-dark: #EDE7F6;
  --white: #FFFFFF;
  --text: #1A0533;
  --text-muted: #7E57C2;
  --border: rgba(124, 77, 255, 0.15);
  --shadow: 0 4px 24px rgba(124, 77, 255, 0.15);
  --radius: 24px;           /* ポップは角丸を大きく */
  --radius-sm: 12px;
}
/* Hero背景グラデーション */
/* background: linear-gradient(135deg, #1A0533 0%, #4A148C 50%, #7C4DFF 100%); */
```

---

## ダーク・シック {#dark}

**用途：** ポートフォリオ（クリエイター系）、音楽、映像、ゲーム

```css
:root {
  --primary: #76FF03;       /* エレクトリックグリーン */
  --primary-dark: #64DD17;
  --accent: #FF6D00;        /* オレンジ */
  --bg: #0A0A0A;
  --bg-dark: #121212;
  --bg-card: #1E1E1E;
  --white: #F5F5F5;
  --text: #EFEFEF;
  --text-muted: #9E9E9E;
  --border: rgba(118, 255, 3, 0.15);
  --shadow: 0 4px 24px rgba(0, 0, 0, 0.5);
  --radius: 12px;
  --radius-sm: 6px;
}
/* ダークテーマのbody設定 */
/* body { background: var(--bg); color: var(--white); } */
/* Hero背景グラデーション */
/* background: linear-gradient(135deg, #0A0A0A 0%, #121212 60%, #1A2A0A 100%); */
```

---

## アクセントカラーとCTAボタンのガイドライン

| 用途 | 色の選択 |
|------|---------|
| CTAボタン（最重要） | `var(--accent)` を使う |
| 見出し強調 | `var(--primary)` |
| テキストリンク | `var(--primary)` with underline |
| 成功・ポジティブ | `#4CAF50` |
| 警告 | `#FF9800` |
| エラー | `#F44336` |

```css
/* アクセントカラーのグロー効果（CTAボタン等） */
.btn-cta {
  background: var(--accent);
  box-shadow: 0 4px 24px color-mix(in srgb, var(--accent) 40%, transparent);
}
.btn-cta:hover {
  box-shadow: 0 8px 32px color-mix(in srgb, var(--accent) 55%, transparent);
}
```

---

## グラデーション組み合わせ例

```css
/* パステルグラデーション（温かみ系） */
background: linear-gradient(135deg, #FFF9F0 0%, #FFF0F5 100%);

/* ガラス（glass morphism）効果 */
.glass-card {
  background: rgba(255, 255, 255, 0.1);
  backdrop-filter: blur(12px);
  border: 1px solid rgba(255, 255, 255, 0.2);
}

/* セクション背景の交互切り替え */
/* 奇数セクション: var(--bg) / 偶数セクション: var(--white) */
```
