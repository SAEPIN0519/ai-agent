const { chromium } = require('playwright');
const path = require('path');

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1080, height: 1080 } });
  const htmlPath = 'file:///' + path.resolve('サムネイル.html').replace(/\\/g, '/');
  await page.goto(htmlPath, { waitUntil: 'networkidle' });
  await page.evaluate(() => {
    document.body.style.margin = '0';
    document.body.style.display = 'block';
    document.body.style.minHeight = 'auto';
    document.body.style.background = 'none';
    const t = document.querySelector('.thumbnail');
    if (t) { t.style.position = 'fixed'; t.style.top = '0'; t.style.left = '0'; }
  });
  await page.screenshot({ path: 'サムネイル_preview.jpg', type: 'jpeg', quality: 85 });
  await browser.close();
  console.log('プレビュー画像生成完了');
})();
