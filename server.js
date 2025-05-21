const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const app = express();
const port = process.env.PORT || 3000;

// 中间件
app.use(express.json()); // 用于解析 application/json
app.use(express.urlencoded({ extended: true })); // 用于解析 application/x-www-form-urlencoded
app.use(express.static(path.join(__dirname, 'public'))); // 提供 public 文件夹中的静态文件

// 确保上传和输出目录存在
const uploadsDir = path.join(__dirname, 'uploads');
const filledDocumentsDir = path.join(__dirname, 'filled_documents');
if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
}
if (!fs.existsSync(filledDocumentsDir)) {
    fs.mkdirSync(filledDocumentsDir, { recursive: true });
}

// 配置 multer 用于文件上传
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, uploadsDir); // 上传文件保存到 'uploads/' 目录
    },
    filename: function (req, file, cb) {
        cb(null, Date.now() + '-' + file.originalname);
    }
});
const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        if (file.fieldname === "wordFile") {
            if (!file.originalname.match(/\.(docx)$/)) {
                return cb(new Error('请只上传 .docx 格式的 Word 文件!'), false);
            }
        } else if (file.fieldname === "jsonFile") {
            if (!file.originalname.match(/\.(json)$/)) {
                return cb(new Error('请只上传 .json 格式的 JSON 文件!'), false);
            }
        }
        cb(null, true);
    }
});

// API 端点：填充 Word 文档
app.post('/api/fill-document', upload.fields([{ name: 'wordFile', maxCount: 1 }, { name: 'jsonFile', maxCount: 1 }]), async (req, res) => {
    let wordFilePath = null;
    let jsonFilePath = null;

    try {
        const wordFile = req.files && req.files['wordFile'] ? req.files['wordFile'][0] : null;
        const uploadedJsonFile = req.files && req.files['jsonFile'] ? req.files['jsonFile'][0] : null;
        let jsonData;

        if (!wordFile) {
            return res.status(400).send({ message: 'Word 模板文件是必需的。' });
        }
        wordFilePath = wordFile.path;

        // 获取 JSON 数据
        if (uploadedJsonFile) {
            jsonFilePath = uploadedJsonFile.path;
            const jsonDataString = fs.readFileSync(jsonFilePath, 'utf-8');
            jsonData = JSON.parse(jsonDataString);
        } else if (req.body.jsonData) {
            try {
                jsonData = JSON.parse(req.body.jsonData);
            } catch (e) {
                return res.status(400).send({ message: '提供的 JSON 文本数据无效。' });
            }
        } else {
            return res.status(400).send({ message: 'JSON 数据是必需的 (可以作为文件上传或文本输入)。' });
        }

        // 读取 Word 模板内容
        const templateContent = fs.readFileSync(wordFilePath);

        const zip = new PizZip(templateContent);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            nullGetter: function() { return ""; } // 对于缺失的标签返回空字符串而不是抛出错误
        });

        doc.setData(jsonData);

        try {
            doc.render(); // 执行模板替换
        } catch (error) {
            // 捕获 Docxtemplater 的渲染错误，例如标签未找到等
            console.error("Docxtemplater 渲染错误:", error);
            let errorMessage = '文档渲染时出错。';
            if (error.properties && error.properties.errors) {
                errorMessage += ' 详情: ' + error.properties.errors.map(e => e.message).join(', ');
            } else {
                errorMessage += ` ${error.message}`;
            }
            return res.status(400).send({ message: errorMessage });
        }

        const outputBuffer = doc.getZip().generate({
            type: 'nodebuffer',
            compression: "DEFLATE"
        });

        const outputFileName = `filled_${Date.now()}_${wordFile.originalname}`;
        const outputPath = path.join(filledDocumentsDir, outputFileName);

        fs.writeFileSync(outputPath, outputBuffer);

        res.download(outputPath, outputFileName, (err) => {
            if (err) {
                console.error('下载文件时出错:', err);
                // 即便下载失败，也应该尝试清理文件
            }
            // 下载完成后（无论成功或失败）清理生成的填充文档
            if (fs.existsSync(outputPath)) {
                fs.unlinkSync(outputPath);
            }
        });

    } catch (error) {
        console.error('处理请求时发生内部错误:', error);
        res.status(500).send({ message: '服务器内部错误。' + error.message });
    } finally {
        // 清理上传的临时文件
        if (wordFilePath && fs.existsSync(wordFilePath)) {
            fs.unlinkSync(wordFilePath);
        }
        if (jsonFilePath && fs.existsSync(jsonFilePath)) {
            fs.unlinkSync(jsonFilePath);
        }
    }
});

// 默认首页路由
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// 全局错误处理中间件 (可选, 用于处理 multer 的错误等)
app.use((err, req, res, next) => {
    if (err instanceof multer.MulterError) {
        // Multer 错误
        return res.status(400).send({ message: `文件上传错误: ${err.message}` });
    } else if (err) {
        // 其他错误
        return res.status(400).send({ message: err.message });
    }
    next();
});


app.listen(port, () => {
    console.log(`服务器正在 http://localhost:${port} 上运行`);
});