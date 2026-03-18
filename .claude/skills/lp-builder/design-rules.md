# コンマリデザイン原則

## 絶対ルール

- **絵文字は一切使わない**（HTML・テキスト・コメントを問わず禁止）
- アイコンは Google Material Icons のみ使用
- 情報商材っぽさ・意図しないダサさを避ける
- 装飾より情報の読みやすさを優先

```html
<link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
<span class="material-icons">arrow_forward</span>
```

---

## 四大原則

| 原則 | 実装方法 |
|------|---------|
| **近接** | 関連要素を `gap: 8–16px`、セクション間は `padding: 80–120px` |
| **整列** | 要素の端・中心を揃える。テキストは左揃えか中央揃えに統一 |
| **反復** | CSS変数でカラー・角丸・シャドウを定義して繰り返す |
| **対比** | 見出しを本文の約1.6倍サイズに（ジャンプ率）。重要語はアクセントカラーで強調 |

---

## カラー構造（CSS変数）

```css
:root {
  --primary: #3D5AFE;      /* メインカラー */
  --primary-dark: #1939B7;
  --accent: #FF6D00;       /* アクセント（CTAボタン等） */
  --bg: #F9F5F0;           /* 背景 */
  --white: #FFFFFF;
  --text: #1A1A2E;         /* 本文 */
  --text-muted: #6B7280;   /* サブテキスト */
  --border: rgba(0,0,0,0.08);
  --shadow: 0 4px 24px rgba(0,0,0,0.08);
  --radius: 16px;
  --radius-sm: 8px;
}
```

配色面積比：メイン62% : アクセント38%（黄金比）

---

## タイポグラフィ

```css
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;700&family=Inter:wght@300;400;600;700&display=swap');

body { font-family: 'Noto Sans JP', 'Inter', sans-serif; }

h1 { font-size: clamp(2rem, 5vw, 3.5rem); font-weight: 700; line-height: 1.2; }
h2 { font-size: clamp(1.5rem, 3vw, 2.25rem); font-weight: 700; }
h3 { font-size: clamp(1.1rem, 2vw, 1.375rem); font-weight: 600; }
p  { font-size: 1rem; line-height: 1.8; }
```

---

## レイアウト

```css
.section-inner {
  max-width: 1100px;
  margin: 0 auto;
  padding: 0 24px;
}

section { padding: 80px 0; }

@media (max-width: 768px) {
  section { padding: 48px 0; }
  .grid-3 { grid-template-columns: 1fr; }
}
```

---

## スクロールアニメーション（Intersection Observer）

```css
.fi {
  opacity: 0;
  transform: translateY(20px);
  transition: opacity 0.6s ease, transform 0.6s ease;
}
.fi.visible { opacity: 1; transform: translateY(0); }
.fi.delay-1 { transition-delay: 0.1s; }
.fi.delay-2 { transition-delay: 0.2s; }
.fi.delay-3 { transition-delay: 0.3s; }
```

```js
document.addEventListener('DOMContentLoaded', () => {
  const items = Array.from(document.querySelectorAll('.fi'));
  const show = el => el.classList.add('visible');
  if (!('IntersectionObserver' in window)) { items.forEach(show); return; }
  const io = new IntersectionObserver((entries, observer) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) { show(entry.target); observer.unobserve(entry.target); }
    });
  }, { threshold: 0.06 });
  items.forEach(el => {
    const rect = el.getBoundingClientRect();
    if (rect.top < window.innerHeight * 0.92) show(el);
    else io.observe(el);
  });
});
```

---

## ナビゲーション

```html
<nav id="mainNav">
  <a href="#" class="nav-logo">Brand<span>Name</span></a>
  <ul class="nav-links">
    <li><a href="#about">About</a></li>
    <li><a href="#features">Features</a></li>
    <li><a href="#contact">Contact</a></li>
  </ul>
  <a href="#contact" class="nav-cta">お問い合わせ</a>
</nav>
```

```css
#mainNav {
  position: fixed; top: 0; left: 0; right: 0; z-index: 100;
  display: flex; align-items: center; justify-content: space-between;
  padding: 20px 40px;
  transition: background 0.3s, box-shadow 0.3s;
}
#mainNav.scrolled {
  background: rgba(255,255,255,0.95);
  backdrop-filter: blur(12px);
  box-shadow: 0 2px 20px rgba(0,0,0,0.08);
}
```

```js
const mainNav = document.getElementById('mainNav');
window.addEventListener('scroll', () => {
  mainNav.classList.toggle('scrolled', window.scrollY > 40);
}, { passive: true });
```

---

## スクロールプログレスバー

```html
<div class="scroll-indicator" id="scrollBar"></div>
```

```css
.scroll-indicator {
  position: fixed; top: 0; left: 0;
  height: 3px; background: var(--primary);
  z-index: 200; width: 0%; transition: width 0.1s;
}
```

```js
const scrollBar = document.getElementById('scrollBar');
window.addEventListener('scroll', () => {
  const scrolled = (window.scrollY / (document.documentElement.scrollHeight - window.innerHeight)) * 100;
  scrollBar.style.width = scrolled + '%';
}, { passive: true });
```

---

## ウェーブ区切り

```html
<div class="wave" style="background: var(--white);">
  <svg viewBox="0 0 1440 90" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="none">
    <path d="M0,45 C180,90 360,5 540,45 C720,85 900,10 1080,45 C1200,68 1360,28 1440,42 L1440,90 L0,90 Z" fill="#F9F5F0"/>
  </svg>
</div>
```

```css
.wave { line-height: 0; }
.wave svg { display: block; width: 100%; height: 90px; }
```

---

## グラデーション Hero

```css
.hero {
  min-height: 100vh;
  display: flex; align-items: center;
  background: linear-gradient(135deg, #1A1A2E 0%, #16213E 50%, #0F3460 100%);
  position: relative; overflow: hidden;
}
/* 背景blob */
.hero-blob {
  position: absolute; border-radius: 50%;
  filter: blur(80px); opacity: 0.15;
  animation: float 8s ease-in-out infinite;
}
@keyframes float {
  0%, 100% { transform: translateY(0) scale(1); }
  50% { transform: translateY(-20px) scale(1.05); }
}
```

---

## CTAボタン

```css
.btn-primary {
  display: inline-flex; align-items: center; gap: 8px;
  padding: 16px 36px; border-radius: 50px;
  background: var(--accent); color: #fff;
  font-size: 1rem; font-weight: 700; text-decoration: none;
  transition: transform 0.2s, box-shadow 0.2s;
  box-shadow: 0 4px 24px rgba(255,109,0,0.3);
}
.btn-primary:hover {
  transform: translateY(-2px);
  box-shadow: 0 8px 32px rgba(255,109,0,0.45);
}
```

---

## 読みやすさチェック（HTML生成後に必ず実行）

- [ ] カード・ボックス内テキストが不自然な位置で改行されていないか
- [ ] 狭いカードに長い日本語を押し込んでいないか（2列 or 横並びを検討）
- [ ] レスポンシブで崩れていないか（グリッド列数 × テキスト量のバランス）
- [ ] プレースホルダーリンク（`#xxx`）が残っていないか
- [ ] 絵文字が混入していないか
