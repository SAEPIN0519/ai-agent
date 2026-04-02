/**
 * SPFx パッケージ (.sppkg) を直接生成するスクリプト
 * HTMLアプリをSharePoint Webパーツとして配置するためのパッケージ
 */
const fs = require('fs');
const path = require('path');
const archiver = require('archiver');

const SOLUTION_ID = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890';
const WEBPART_ID = 'f1e2d3c4-b5a6-7890-fedc-ba9876543210';
const PACKAGE_NAME = 'mold-inventory-management';
const WEBPART_TITLE = '金型在庫管理';

// HTMLアプリの内容を読み込み
const htmlContent = fs.readFileSync(
    path.join(__dirname, '..', '金型在庫管理_standalone.html'), 'utf8'
);

// HTMLをエスケープしてJSの中に埋め込む
const escapedHtml = htmlContent
    .replace(/\\/g, '\\\\')
    .replace(/`/g, '\\`')
    .replace(/\$/g, '\\$');

// WebパーツのJavaScriptバンドル
const webpartBundle = `
"use strict";
(function() {
    var __extends = (this && this.__extends) || (function () {
        var extendStatics = function (d, b) {
            extendStatics = Object.setPrototypeOf || ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) || function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
            return extendStatics(d, b);
        };
        return function (d, b) { extendStatics(d, b); function __() { this.constructor = d; } d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __()); };
    })();

    // Webパーツクラスのレンダリング
    window.__renderMoldInventory = function(domElement) {
        // iframe を使ってHTMLアプリを表示
        var iframe = document.createElement('iframe');
        iframe.style.width = '100%';
        iframe.style.height = '100vh';
        iframe.style.border = 'none';
        iframe.style.overflow = 'auto';

        domElement.appendChild(iframe);

        var doc = iframe.contentDocument || iframe.contentWindow.document;
        doc.open();
        doc.write(\`${escapedHtml}\`);
        doc.close();
    };
})();
`;

// Feature XML
const featureXml = `<?xml version="1.0" encoding="utf-8"?>
<Feature xmlns="http://schemas.microsoft.com/sharepoint/"
    Id="${SOLUTION_ID}"
    Title="${WEBPART_TITLE}"
    Description="金型在庫管理システム - ラック配置図から入出庫管理"
    Version="1.0.0.0">
</Feature>`;

// Webパーツマニフェスト
const manifest = {
    "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json",
    "id": WEBPART_ID,
    "alias": "MoldInventoryWebPart",
    "componentType": "WebPart",
    "version": "1.0.0",
    "manifestVersion": 2,
    "requiresCustomScript": true,
    "supportedHosts": ["SharePointWebPart", "SharePointFullPage"],
    "preconfiguredEntries": [{
        "groupId": "5c03119e-3074-46fd-976b-c60198311f70",
        "group": { "default": "金型管理" },
        "title": { "default": WEBPART_TITLE },
        "description": { "default": "金型在庫管理システム" },
        "officeFabricIconFontName": "BoxMultiplySolid",
        "properties": {}
    }]
};

// Solution マニフェスト（AppManifest.xml 相当）
const appManifestXml = `<?xml version="1.0" encoding="utf-8"?>
<App xmlns="http://schemas.microsoft.com/sharepoint/2012/app/manifest"
     Name="${PACKAGE_NAME}"
     ProductID="{${SOLUTION_ID}}"
     Version="1.0.0.0"
     SharePointMinVersion="16.0.0.0">
  <Properties>
    <Title>${WEBPART_TITLE}</Title>
    <StartPage>~appWebUrl</StartPage>
  </Properties>
  <AppPrincipal>
    <Internal/>
  </AppPrincipal>
</App>`;

// package-solution.json 相当の情報
const solutionConfig = {
    "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json",
    "solution": {
        "name": PACKAGE_NAME,
        "id": SOLUTION_ID,
        "version": "1.0.0.0",
        "includeClientSideAssets": true,
        "skipFeatureDeployment": true,
        "isDomainIsolated": false,
        "developer": {
            "name": "SAEPIN",
            "websiteUrl": "",
            "privacyUrl": "",
            "termsOfUseUrl": "",
            "mpnId": "Undefined-1.0.0"
        },
        "metadata": {
            "shortDescription": { "default": "金型在庫管理システム" },
            "longDescription": { "default": "ラック配置図から入出庫管理を行うWebパーツ" },
            "screenshotPaths": [],
            "videoUrl": "",
            "categories": []
        },
        "features": [{
            "title": WEBPART_TITLE,
            "description": "金型在庫管理Webパーツ",
            "id": SOLUTION_ID,
            "version": "1.0.0.0",
            "componentIds": [WEBPART_ID]
        }]
    },
    "paths": { "zippedPackage": "solution/" + PACKAGE_NAME + ".sppkg" }
};

// ClientSideAssets マニフェスト
const clientSideManifest = {
    "id": WEBPART_ID,
    "alias": "MoldInventoryWebPart",
    "componentType": "WebPart",
    "version": "1.0.0",
    "manifestVersion": 2,
    "loaderConfig": {
        "internalModuleBaseUrls": [""],
        "entryModuleId": "mold-inventory-webpart",
        "scriptResources": {
            "mold-inventory-webpart": {
                "type": "localizedPath",
                "paths": {
                    "default": { "path": "mold-inventory-webpart.js", "integrity": null }
                },
                "defaultPath": { "path": "mold-inventory-webpart.js", "integrity": null }
            }
        }
    },
    "requiresCustomScript": true
};

async function buildPackage() {
    const outDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

    const sppkgPath = path.join(outDir, PACKAGE_NAME + '.sppkg');
    const output = fs.createWriteStream(sppkgPath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    archive.pipe(output);

    // Feature XML
    archive.append(featureXml, { name: `${SOLUTION_ID}/Feature.xml` });

    // Webパーツマニフェスト
    archive.append(JSON.stringify(manifest, null, 2), {
        name: `${SOLUTION_ID}/${WEBPART_ID}.manifest.json`
    });

    // JSバンドル
    archive.append(webpartBundle, {
        name: `${SOLUTION_ID}/mold-inventory-webpart.js`
    });

    // AppManifest
    archive.append(appManifestXml, { name: 'AppManifest.xml' });

    // [Content_Types].xml
    const contentTypes = `<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml" />
  <Default Extension="json" ContentType="application/json" />
  <Default Extension="js" ContentType="application/javascript" />
</Types>`;
    archive.append(contentTypes, { name: '[Content_Types].xml' });

    await archive.finalize();

    return new Promise((resolve, reject) => {
        output.on('close', () => {
            console.log(`✓ ${sppkgPath} (${(archive.pointer() / 1024).toFixed(1)} KB)`);
            resolve();
        });
        output.on('error', reject);
    });
}

buildPackage().then(() => {
    console.log('\n完了！');
    console.log('このファイルをSharePointのアプリカタログにアップロードしてね。');
}).catch(console.error);
