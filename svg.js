const fs = require('fs');
const pptxgen = require('pptxgenjs');
const { parse } = require('svgson');
const colornames = require('colornames'); // カラーネーム変換

// SVGファイルのパス
const svgFilePath = "1.svg";

// PowerPointのスライドサイズ（デフォルト: 10 × 5.625 インチ）
const pptWidth = 10;
const pptHeight = 5.625;

// デバッグ用ログ
const DEBUG = false;

// SVGファイルの読み込み
fs.readFile(svgFilePath, 'utf8', (err, data) => {
    if (err) {
        console.error("SVGファイルの読み込みエラー:", err);
        return;
    }

    // SVGの解析
    parse(data).then(svgJson => {
        if (DEBUG) console.log("SVG要素:", JSON.stringify(svgJson, null, 2));

        const pptx = new pptxgen();
        let slide = pptx.addSlide();
        let rects = [];

        const viewBox = svgJson.attributes.viewBox ? svgJson.attributes.viewBox.split(" ") : [0, 0, 1600, 900];
        const [vbX, vbY, vbWidth, vbHeight] = viewBox.map(Number);

        // スケール変換係数 (SVG → PPT)
        const scaleX = pptWidth / vbWidth;
        const scaleY = pptHeight / vbHeight;
        const scale = Math.min(scaleX, scaleY); // 等比スケール

        // 色変換関数（RGBA & colornames 対応）
        function convertColor(color) {
            if (!color) return "000000"; // デフォルトの黒
            if (colornames(color)) return colornames(color).replace("#", "").toUpperCase();
            if (color.startsWith("rgba")) {
                let rgba = color.match(/([\d.]+)/g);
                if (rgba && rgba.length === 4) {
                    let hex = (
                        ((1 << 24) + (parseInt(rgba[0]) << 16) + (parseInt(rgba[1]) << 8) + parseInt(rgba[2]))
                        .toString(16)
                        .slice(1)
                    );
                    return hex.toUpperCase();
                }
            }

            if (color.startsWith("url(")) return "2563EB"; // グラデーションは青 (#2563EB) に統一
            return color.replace("#", "").toUpperCase(); // PPTX用のRGBフォーマット
        }

        // SVG要素をPPTXに変換する関数
        function processElement(element, parentTransform = { x: 0, y: 0 }) {
            const { name, attributes, children } = element;
            let x = (parseFloat(attributes.x || 0) - vbX + parentTransform.x) * scale;
            let y = (parseFloat(attributes.y || 0) - vbY + parentTransform.y) * scale;
            let width = parseFloat(attributes.width || 100) * scale;
            let height = parseFloat(attributes.height || 30) * scale;
            let rx = parseFloat(attributes.rx || 0) * scale;
            let ry = parseFloat(attributes.ry || 0) * scale;
            let fill = convertColor(attributes.fill || "#FFFFFF");
            let opacity = parseFloat(attributes.opacity || attributes["fill-opacity"] || 1);
        
            let fontSize = attributes["font-size"] ? parseFloat(attributes["font-size"].replace("px", "")) : 14;
            let textColor = convertColor(attributes.fill || "#000000");
            
            
            // **🔹 rect の処理**
            if (name === "rect") {
                let rectOptions = {
                    x: x,
                    y: y,
                    w: width,
                    h: height,
                    fill: { color: fill, transparency: 100 - (opacity * 100) }
                };

                // **スケール変換を適用**
                let rxScaled = rx * scale;
                let ryScaled = ry * scale;

                // **過度な丸みを抑制**
                let maxRadius = Math.min(width * 0.2, height * 0.2); // 最大でも幅・高さの20%まで
                let adjustedRx = Math.min(rxScaled, maxRadius);
                let adjustedRy = Math.min(ryScaled, maxRadius);
                let finalRadius = Math.max(adjustedRx, adjustedRy); // より大きい方を採用

                // **角丸の処理**
                if (finalRadius > 0) {
                    rectOptions.radius = finalRadius; // 適切な丸みを適用
                    slide.addShape(pptx.ShapeType.roundRect, rectOptions);
                } else {
                    slide.addShape(pptx.ShapeType.rect, rectOptions);
                }
            }

            // テキストの処理
            if (name === "text") {
                let textRuns = [];
                let totalText = ""; // すべてのテキストを統合
            
                children.forEach(child => {
                    if (child.name === "tspan") {
                        let tspanColor = convertColor(child.attributes.fill || textColor);
                        let tspanText = child.children.map(c => c.value || "").join("").trim();
                        let isBold = child.attributes["font-weight"] === "bold"; // 太字判定
            
                        totalText += tspanText; // 文字数を合計
                        if (tspanText) {
                            textRuns.push({
                                text: tspanText,
                                options: {
                                    color: tspanColor,
                                    bold: isBold // 太字を適用
                                }
                            });
                        }
                    } else {
                        let normalText = child.value || "";
                        let isBold = attributes["font-weight"] === "bold"; // `text` 自体が太字か判定
            
                        totalText += normalText; // 文字数を合計
                        if (normalText.trim()) {
                            textRuns.push({
                                text: normalText,
                                options: {
                                    color: textColor,
                                    bold: isBold // 太字を適用
                                }
                            });
                        }
                    }
                });
            
                if (textRuns.length > 0) {
                    let fontSizePx = fontSize;
                    let fontSizePt = fontSizePx * 0.75 / 1.5; // px → pt変換 ＋ スケールの影響を調整
            
                    const ptToCm = 0.0352778; // 1 pt = 0.0352778 cm
                    let textBoxHeight = fontSizePt * ptToCm; // 高さを調整
                    let textBoxWidth = fontSizePt * totalText.length * ptToCm * 0.7; // 幅を `totalText.length` に基づいて設定
            
                    // **🔹 直前の `rect` を取得し、その範囲に収める **
                    let lastRect = null;
                    for (let i = rects.length - 1; i >= 0; i--) {
                        if (y >= rects[i].y && y <= rects[i].y + rects[i].h) {
                            lastRect = rects[i];
                            break;
                        }
                    }
            
                    if (lastRect) {
                        let rectStartX = lastRect.x;
                        let rectEndX = lastRect.x + lastRect.w;
                        let textEndX = x + textBoxWidth;
            
                        if (textEndX > rectEndX) {
                            textBoxWidth = rectEndX - x; // `text` が `rect` を超えないように調整
                        }
                    }

                    // **🔹 `text-anchor` の影響を考慮したX座標調整**
                    let textAlign = "left"; // デフォルトは左揃え
                    let xOffset = 0;

                    if (attributes["text-anchor"] === "middle") {
                        textAlign = "center";
                        xOffset = -textBoxWidth / 2;
                    } else if (attributes["text-anchor"] === "end") {
                        textAlign = "right";
                        xOffset = -textBoxWidth;
                    }

                    let correctedX = (x + xOffset);
                    let correctedY = y - (textBoxHeight / 2);

                    // **🔹 PowerPointの `align` オプションも適用**
                    slide.addText(textRuns, {
                        x: correctedX,
                        y: correctedY,
                        fontSize: fontSizePt, // 適切にスケールされたフォントサイズ
                        w: textBoxWidth, // cm単位の幅
                        h: textBoxHeight,  // cm単位の高さ
                        autoFit: true, // 自動調整を有効にする
                        align: textAlign // `left` / `center` / `right` を適用
                    });
                }
            } 

            // **🔹 circle の処理**
            else if (name === "circle") {
                let cx = parseFloat(attributes.cx || 0) - vbX;
                let cy = parseFloat(attributes.cy || 0) - vbY;
                let r = parseFloat(attributes.r || 10);

                slide.addShape(pptx.ShapeType.ellipse, {
                    x: (cx + parentTransform.x) * scale, 
                    y: (cy + parentTransform.y) * scale, 
                    w: r * 2 * scale, 
                    h: r * 2 * scale,
                    fill: { color: fill, transparency: 100 - opacity * 100 }
                });
            } 
            
            // **🔹 polygon の処理**
            else if (name === "polygon") {
                let points = attributes.points.split(" ").map(p => {
                    let [px, py] = p.split(",").map(Number);
                    return { x: px * scale, y: py * scale };
                });

                slide.addShape(pptx.ShapeType.triangle, {
                    x, y, w: width, h: height, fill: { color: fill }
                });
            }

            // **🔹 グループ <g> の処理**
            else if (name === "g") {
                let groupTransform = { x, y };
                children.forEach(child => processElement(child, groupTransform));
                return;
            }

            if (children && children.length > 0) {
                children.forEach(child => processElement(child, parentTransform));
            }
        }

        svgJson.children.forEach(element => processElement(element));

        pptx.writeFile({ fileName: "presentation.pptx" }).then(() => {
            console.log("PPTXファイルが作成されました");
        });
    });
});
