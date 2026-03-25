/**
 * サムネイル動画録画スクリプト
 * サムネイル.htmlのアニメーションをMP4動画として録画する
 *
 * 使い方: node record-video.js
 * 出力: サムネイル動画.webm
 */

const { chromium } = require('playwright');
const path = require('path');

(async () => {
  const htmlPath = path.resolve(__dirname, 'サムネイル.html');
  const outputDir = path.resolve(__dirname, 'video-output');

  console.log('ブラウザを起動中...');

  const browser = await chromium.launch();
  const context = await browser.newContext({
    // 1080x1080のInstagram正方形サイズで録画
    viewport: { width: 1080, height: 1080 },
    recordVideo: {
      dir: outputDir,
      size: { width: 1080, height: 1080 }
    }
  });

  const page = await context.newPage();

  console.log('サムネイルを読み込み中...');
  await page.goto(`file:///${htmlPath.replace(/\\/g, '/')}`, {
    waitUntil: 'networkidle'
  });

  // bodyの余白を消してサムネイルだけ表示
  await page.evaluate(() => {
    document.body.style.margin = '0';
    document.body.style.padding = '0';
    document.body.style.display = 'block';
    document.body.style.minHeight = 'auto';
    document.body.style.background = 'none';
    const thumb = document.querySelector('.thumbnail');
    if (thumb) {
      thumb.style.position = 'fixed';
      thumb.style.top = '0';
      thumb.style.left = '0';
    }
  });

  // アニメーションを15秒間録画（花びらの1サイクル分）
  const duration = 15000;
  console.log(`${duration / 1000}秒間アニメーションを録画中...`);
  await page.waitForTimeout(duration);

  // 録画を停止・保存
  const videoPath = await page.video().path();
  await context.close();
  await browser.close();

  console.log(`\n録画完了！`);
  console.log(`保存先: ${videoPath}`);
  console.log(`\nInstagramに投稿するには:`);
  console.log(`1. このファイルをスマホに送る（LINE/Google Drive/AirDrop）`);
  console.log(`2. Instagram → + → リール → 動画を選択 → 投稿`);
  console.log(`\n※ .webm形式のまま投稿できない場合は、スマホの動画編集アプリで.mp4に変換してね`);
})();
