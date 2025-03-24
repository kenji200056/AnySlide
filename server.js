const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const { convertSvgToPptx } = require('./converter');

const app = express();
const upload = multer({ dest: 'uploads/' });

// 静的ファイルの提供（index.htmlのあるフォルダ）
app.use(express.static('public'));

// ファイルアップロードと変換処理
app.post('/upload', upload.single('svgFile'), async (req, res) => {
    try {
        if (!req.file) {
            console.error("❌ SVGファイルがアップロードされていません");
            return res.status(400).send("❌ SVGファイルがアップロードされていません");
        }

        const svgPath = req.file.path; // multerが保存したアップロードファイルのパス
        const outputPptxPath = `uploads/${Date.now()}.pptx`; // 出力PPTXファイル名

        console.log("✅ アップロード完了 - SVGファイル:", svgPath);

        // SVG → PPTX 変換
        const pptxFileName = await convertSvgToPptx(svgPath, outputPptxPath);
        console.log("✅ 変換完了 - PPTXファイル:", pptxFileName);

        // 変換が成功した場合のみダウンロード
        if (pptxFileName && fs.existsSync(pptxFileName)) {
            res.download(pptxFileName, 'converted.pptx', (err) => {
                if (err) {
                    console.error("❌ ファイルダウンロードエラー:", err);
                }
                // ダウンロード完了後に削除
                fs.unlinkSync(svgPath);
                fs.unlinkSync(pptxFileName);
            });
        } else {
            console.error("❌ PPTXファイルの作成に失敗しました");
            res.status(500).send("❌ PPTXファイルの作成に失敗しました");
        }
    } catch (error) {
        console.error("❌ 変換エラー:", error);
        res.status(500).send("❌ 変換に失敗しました");
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`✅ サーバー起動: http://localhost:${PORT}`);
});
