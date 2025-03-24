const fs = require('fs');
const { google } = require('googleapis');
const { parse } = require('svgson');
const colornames = require('colornames'); // カラーネーム変換

// Google API 設定
const SCOPES = ['https://www.googleapis.com/auth/presentations'];
const CREDENTIALS_PATH = 'credentials.json'; // GCP で取得した認証情報
const svgFilePath = "1.svg"; // 変換対象の SVG ファイル

// Google API の認証
async function authorize() {
    const content = fs.readFileSync(CREDENTIALS_PATH);
    const credentials = JSON.parse(content);
    const { client_secret, client_id, redirect_uris } = credentials.installed;

    const auth = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
    const tokenPath = 'token.json';

    if (fs.existsSync(tokenPath)) {
        auth.setCredentials(JSON.parse(fs.readFileSync(tokenPath)));
        return auth;
    }

    throw new Error('OAuth トークンが必要です。Google Cloud Console で OAuth 認証を完了し、token.json を保存してください。');
}

// Google スライドの新規作成
async function createPresentation(auth) {
    const slides = google.slides({ version: 'v1', auth });

    const presentation = await slides.presentations.create({
        requestBody: { title: 'SVG to Google Slides' },
    });

    console.log(`📄 プレゼンテーション作成: ${presentation.data.presentationId}`);
    return presentation.data.presentationId;
}

// SVG を Google スライドへ変換
async function addSvgToSlides(auth, presentationId, svgFilePath) {
    const slides = google.slides({ version: 'v1', auth });

    const svgData = fs.readFileSync(svgFilePath, 'utf8');
    const svgJson = await parse(svgData);

    // SVGの `viewBox` からスケール変換を決定
    const viewBox = svgJson.attributes.viewBox ? svgJson.attributes.viewBox.split(" ") : [0, 0, 1600, 900];
    const [vbX, vbY, vbWidth, vbHeight] = viewBox.map(Number);

    const pptWidth = 10 * 72; // Google スライドの幅 (pt)
    const pptHeight = 5.625 * 72; // Google スライドの高さ (pt)
    const scaleX = pptWidth / vbWidth;
    const scaleY = pptHeight / vbHeight;
    const scale = Math.min(scaleX, scaleY); // 比率を統一

    const requests = [];

    for (const element of svgJson.children) {
        if (element.name === 'text') {
            let x = (parseFloat(element.attributes.x || 0) - vbX) * scale;
            let y = (parseFloat(element.attributes.y || 0) - vbY) * scale;
            let fontSize = parseFloat(element.attributes["font-size"] || 14) * 0.75 * scale;

            requests.push({
                createShape: {
                    objectId: `text_${Date.now()}`,
                    shapeType: 'TEXT_BOX',
                    elementProperties: {
                        pageObjectId: 'p', // スライド ID
                        size: {
                            height: { magnitude: fontSize * 1.5, unit: 'PT' },
                            width: { magnitude: fontSize * element.children.length * 0.6, unit: 'PT' },
                        },
                        transform: {
                            scaleX: 1,
                            scaleY: 1,
                            translateX: x,
                            translateY: y,
                            unit: 'PT',
                        },
                    },
                },
            });
        }

        if (element.name === 'rect') {
            let x = (parseFloat(element.attributes.x || 0) - vbX) * scale;
            let y = (parseFloat(element.attributes.y || 0) - vbY) * scale;
            let width = parseFloat(element.attributes.width || 100) * scale;
            let height = parseFloat(element.attributes.height || 30) * scale;
            let fillColor = convertColor(element.attributes.fill || "#FFFFFF");

            requests.push({
                createShape: {
                    objectId: `rect_${Date.now()}`,
                    shapeType: 'RECTANGLE',
                    elementProperties: {
                        pageObjectId: 'p',
                        size: { width: { magnitude: width, unit: 'PT' }, height: { magnitude: height, unit: 'PT' } },
                        transform: {
                            scaleX: 1,
                            scaleY: 1,
                            translateX: x,
                            translateY: y,
                            unit: 'PT',
                        },
                    },
                    shapeProperties: {
                        shapeBackgroundFill: {
                            solidFill: {
                                color: { rgbColor: { red: fillColor[0] / 255, green: fillColor[1] / 255, blue: fillColor[2] / 255 } },
                            },
                        },
                    },
                },
            });
        }

        if (element.name === 'circle') {
            let cx = (parseFloat(element.attributes.cx || 0) - vbX) * scale;
            let cy = (parseFloat(element.attributes.cy || 0) - vbY) * scale;
            let r = parseFloat(element.attributes.r || 10) * scale;
            let fillColor = convertColor(element.attributes.fill || "#000000");

            requests.push({
                createShape: {
                    objectId: `circle_${Date.now()}`,
                    shapeType: 'ELLIPSE',
                    elementProperties: {
                        pageObjectId: 'p',
                        size: { width: { magnitude: r * 2, unit: 'PT' }, height: { magnitude: r * 2, unit: 'PT' } },
                        transform: {
                            scaleX: 1,
                            scaleY: 1,
                            translateX: cx - r,
                            translateY: cy - r,
                            unit: 'PT',
                        },
                    },
                    shapeProperties: {
                        shapeBackgroundFill: {
                            solidFill: {
                                color: { rgbColor: { red: fillColor[0] / 255, green: fillColor[1] / 255, blue: fillColor[2] / 255 } },
                            },
                        },
                    },
                },
            });
        }
    }

    await slides.presentations.batchUpdate({
        presentationId,
        requestBody: { requests },
    });

    console.log('✅ SVGのデータをGoogleスライドに追加しました');
}

// カラー変換関数
function convertColor(color) {
    if (!color) return [0, 0, 0]; // デフォルト黒
    if (colornames(color)) {
        const hex = colornames(color).replace("#", "");
        return [parseInt(hex.substr(0, 2), 16), parseInt(hex.substr(2, 2), 16), parseInt(hex.substr(4, 2), 16)];
    }
    return [37, 99, 235]; // デフォルトは青
}

// メイン処理
(async () => {
    try {
        const auth = await authorize();
        const presentationId = await createPresentation(auth);
        await addSvgToSlides(auth, presentationId, '1.svg');
    } catch (error) {
        console.error('❌ エラー:', error);
    }
})();
