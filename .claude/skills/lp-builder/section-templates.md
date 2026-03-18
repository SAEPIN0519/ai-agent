# ページ別セクション構成テンプレート

## 目次
- [LP（ランディングページ）](#lp)
- [ポートフォリオサイト](#portfolio)
- [サービス紹介](#service)
- [イベントページ](#event)
- [採用ページ](#recruitment)

---

## LP（ランディングページ）{#lp}

**用途：** 商品・サービスの単品販売、申し込み、登録など1つのCTAに絞るページ

| # | セクション名 | 役割 |
|---|------------|------|
| 1 | Hero | キャッチコピー + サブコピー + CTA + ヒーロービジュアル |
| 2 | Problem | ターゲットが抱える悩み・課題を言語化（共感を得る） |
| 3 | Solution | このサービスがその課題を解決する説明 |
| 4 | Features | 特徴・機能を3〜6個のカードで提示 |
| 5 | Results | 実績・数字・ビフォーアフター（信頼構築） |
| 6 | Testimonials | 利用者の声（3枚のカード形式） |
| 7 | FAQ | よくある質問・不安解消（4〜6問） |
| 8 | Pricing | 料金プラン（なければCTA強化セクションに変更） |
| 9 | CTA | 最終行動喚起 + 申し込みフォームまたはボタン |

```
Hero → Problem → Solution → Features → Results → Testimonials → FAQ → Pricing → CTA
```

---

## ポートフォリオサイト {#portfolio}

**用途：** 個人・フリーランスの実績・スキル・人柄を伝える

| # | セクション名 | 役割 |
|---|------------|------|
| 1 | Hero | 名前 + 肩書き + キャッチコピー + アバター/ビジュアル |
| 2 | About | 自己紹介・ストーリー・価値観（人柄を伝える） |
| 3 | Skills | 得意領域・スキルセット（アイコン付きカード） |
| 4 | Works | 代表実績3〜6件（サムネイル + 概要 + リンク） |
| 5 | Process | 仕事の進め方・フロー（信頼感を高める） |
| 6 | Testimonials | クライアント・受講者の声 |
| 7 | Speaking | 登壇・メディア掲載実績（なければ省略） |
| 8 | Contact | 問い合わせフォームまたはSNSリンク |

```
Hero → About → Skills → Works → Process → Testimonials → Contact
```

---

## サービス紹介 {#service}

**用途：** BtoBのサービス・ツール・SaaSなど、複数の機能を持つサービスの紹介

| # | セクション名 | 役割 |
|---|------------|------|
| 1 | Hero | サービス名 + キャッチコピー + ダッシュボード/スクリーンショット |
| 2 | Problem | 既存の課題・不便さ（共感） |
| 3 | Features | 主要機能3〜6個（アイコン付きカード） |
| 4 | How It Works | 使い方フロー（3ステップ） |
| 5 | Use Cases | ユースケース・活用シーン |
| 6 | Results | 導入実績・数字・ケーススタディ |
| 7 | Testimonials | 利用企業・ユーザーの声 |
| 8 | Pricing | 料金プラン（Free / Pro / Enterprise等） |
| 9 | CTA | 無料トライアル・問い合わせ |

```
Hero → Problem → Features → How It Works → Use Cases → Results → Testimonials → Pricing → CTA
```

---

## イベントページ {#event}

**用途：** セミナー・ワークショップ・勉強会・イベントの告知・申し込み

| # | セクション名 | 役割 |
|---|------------|------|
| 1 | Hero | イベントタイトル + 日時・場所 + 申し込みCTA |
| 2 | About | イベントの概要・テーマ・参加するとどうなるか |
| 3 | Speakers | 登壇者紹介（写真 + 名前 + 肩書き + プロフィール） |
| 4 | Schedule | タイムテーブル（時間 + セッション名） |
| 5 | Target | こんな人に来てほしい（箇条書き3〜5個） |
| 6 | Testimonials | 過去参加者の声（初回なら省略可） |
| 7 | FAQ | 参加方法・場所・持ち物等のQ&A |
| 8 | CTA | 申し込みフォームへのリンク |

```
Hero → About → Speakers → Schedule → Target → FAQ → CTA
```

---

## 採用ページ {#recruitment}

**用途：** 求人・採用活動のためのページ

| # | セクション名 | 役割 |
|---|------------|------|
| 1 | Hero | 採用メッセージ + ビジュアル（チーム写真等） |
| 2 | Mission | 会社のミッション・ビジョン・バリュー |
| 3 | Culture | カルチャー・働き方・チームの雰囲気 |
| 4 | Benefits | 福利厚生・制度・待遇 |
| 5 | Members | メンバー紹介（写真 + 名前 + 役割 + 一言） |
| 6 | Jobs | 募集職種一覧（役割 + 雇用形態 + 待遇） |
| 7 | Process | 選考フロー（3〜4ステップ） |
| 8 | CTA | 応募ボタン + 説明会情報 |

```
Hero → Mission → Culture → Benefits → Members → Jobs → Process → CTA
```

---

## セクション共通パーツ

### eyebrow（上部の小見出し）
```html
<p class="section-eyebrow">Features</p>
<h2 class="section-title">主な<span class="accent">機能</span></h2>
<p class="section-lead">サービスの主要機能をご紹介します。</p>
```

### 3カラムカードグリッド
```html
<div class="grid-3">
  <div class="card">
    <span class="material-icons card-icon">star</span>
    <h3>機能名</h3>
    <p>説明文</p>
  </div>
</div>
```
```css
.grid-3 { display: grid; grid-template-columns: repeat(3, 1fr); gap: 24px; }
@media (max-width: 768px) { .grid-3 { grid-template-columns: 1fr; } }
```

### 3ステップフロー
```html
<div class="steps">
  <div class="step">
    <div class="step-num">01</div>
    <h3>ステップ名</h3>
    <p>説明</p>
  </div>
</div>
```

### 統計・数字セクション
```html
<div class="stats-grid">
  <div class="stat">
    <div class="stat-num">300<span>+</span></div>
    <div class="stat-label">利用者数</div>
  </div>
</div>
```
