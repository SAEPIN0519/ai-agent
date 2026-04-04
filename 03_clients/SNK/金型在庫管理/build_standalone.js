/**
 * 金型在庫管理HTMLにマスターデータを埋め込んでstandalone版を生成
 */
const fs = require('fs');
const path = require('path');

const htmlPath = path.join(__dirname, '金型在庫管理.html');
const jsPath = path.join(__dirname, 'mold_master.js');
const outPath = path.join(__dirname, '金型在庫管理_standalone.html');

let html = fs.readFileSync(htmlPath, 'utf8');
const masterData = fs.readFileSync(jsPath, 'utf8');

html = html.replace(
    '<script id="mold-master-embed">/* マスターデータは末尾に埋め込み */</script>',
    '<script>' + masterData + '</script>'
);

fs.writeFileSync(outPath, html);
const size = (fs.statSync(outPath).size / 1024).toFixed(1);
console.log(`  OK (${size}KB)`);
