const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const app = express();
const port = process.env.PORT || 3000;

// 中间件
app.use(express.json({ limit: '50mb' })); // 用于解析 application/json，增加大小限制以支持较大的Base64文件
app.use(express.urlencoded({ extended: true, limit: '50mb' })); // 用于解析 application/x-www-form-urlencoded
app.use(express.static(path.join(__dirname, 'public')));

const uploadsDir = path.join(__dirname, 'uploads');
const filledDocumentsDir = path.join(__dirname, 'filled_documents');
if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
}
if (!fs.existsSync(filledDocumentsDir)) {
    fs.mkdirSync(filledDocumentsDir, { recursive: true });
}

const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, uploadsDir);
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
                return cb(new Error('请只上传 .docx 格式的 Word 文件! (multer)'), false);
            }
        } else if (file.fieldname === "jsonFile") {
            if (!file.originalname.match(/\.(json)$/)) {
                return cb(new Error('请只上传 .json 格式的 JSON 文件! (multer)'), false);
            }
        }
        cb(null, true);
    },
    limits: { fileSize: 50 * 1024 * 1024 } // 增加 multer 的文件大小限制
});

app.post('/api/fill-document', upload.fields([{ name: 'wordFile', maxCount: 1 }, { name: 'jsonFile', maxCount: 1 }]), async (req, res) => {
    let wordFilePath = null; // 用于 multer 上传的文件路径清理
    let jsonFilePath = null; // 用于 multer 上传的 JSON 文件路径清理
    let templateContent;
    let originalWordFileName = 'template.docx'; // 默认或从请求中获取

    try {
        // --- 1. 获取 Word 模板内容 ---
        const wordFileFromMultipart = req.files && req.files['wordFile'] ? req.files['wordFile'][0] : null;

        if (wordFileFromMultipart) {
            wordFilePath = wordFileFromMultipart.path;
            templateContent = fs.readFileSync(wordFilePath);
            originalWordFileName = wordFileFromMultipart.originalname;
        } else if (req.body.wordFileBase64) {
            originalWordFileName = req.body.wordFileName || `document_${Date.now()}.docx`; // 允许客户端指定文件名
            let base64String = req.body.wordFileBase64;

            // 移除常见的 Base64 数据 URI 前缀
            const prefixesToRemove = [
                "data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,",
                "data:application/msword;base64,",
                "data:application/octet-stream;base64,"
            ];
            for (const prefix of prefixesToRemove) {
                if (base64String.startsWith(prefix)) {
                    base64String = base64String.substring(prefix.length);
                    break;
                }
            }

            try {
                templateContent = Buffer.from(base64String, 'base64');
            } catch (e) {
                return res.status(400).send({ message: '提供的 Word 文件 Base64 数据无效。 ' + e.message });
            }
        } else {
            return res.status(400).send({ message: 'Word 模板文件是必需的 (作为文件上传或 Base64 字符串)。' });
        }

        // --- 2. 获取 JSON 数据 ---
        let jsonData;
        const jsonFileFromMultipart = req.files && req.files['jsonFile'] ? req.files['jsonFile'][0] : null;

        if (jsonFileFromMultipart) {
            jsonFilePath = jsonFileFromMultipart.path;
            const jsonDataString = fs.readFileSync(jsonFilePath, 'utf-8');
            try {
                jsonData = JSON.parse(jsonDataString);
            } catch (e) {
                 return res.status(400).send({ message: '上传的 JSON 文件内容解析失败。 ' + e.message });
            }
        } else if (req.body.jsonFileBase64) {
            let jsonBase64String = req.body.jsonFileBase64;
            const jsonPrefix = "data:application/json;base64,";
             if (jsonBase64String.startsWith(jsonPrefix)) {
                jsonBase64String = jsonBase64String.substring(jsonPrefix.length);
            }
            try {
                const jsonDataString = Buffer.from(jsonBase64String, 'base64').toString('utf-8');
                jsonData = JSON.parse(jsonDataString);
            } catch (e) {
                return res.status(400).send({ message: '提供的 JSON 文件 Base64 数据无效或解码后不是有效的 JSON。 ' + e.message });
            }
        } else if (req.body.jsonData) {
            try {
                if (typeof req.body.jsonData === 'string') {
                    jsonData = JSON.parse(req.body.jsonData);
                } else if (typeof req.body.jsonData === 'object') { // 如果 express.json() 已经解析
                    jsonData = req.body.jsonData;
                } else {
                     return res.status(400).send({ message: '提供的 jsonData 字段必须是 JSON 字符串或对象。' });
                }
            } catch (e) {
                return res.status(400).send({ message: '提供的 JSON 文本数据无效。 ' + e.message });
            }
        } else {
            return res.status(400).send({ message: 'JSON 数据是必需的 (作为文件上传、Base64 字符串或文本输入)。' });
        }

        // --- 3. 处理 Word 文档 ---
        const zip = new PizZip(templateContent);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            nullGetter: function() { return ""; }
        });

        doc.render(jsonData); // 使用更新后的 render 方法

        const outputBuffer = doc.getZip().generate({
            type: 'nodebuffer',
            compression: "DEFLATE"
        });

        const outputFileName = `filled_${Date.now()}_${path.basename(originalWordFileName)}`; // 使用 path.basename 避免路径问题
        const outputPath = path.join(filledDocumentsDir, outputFileName);

        fs.writeFileSync(outputPath, outputBuffer);

        res.download(outputPath, outputFileName, (err) => {
            if (err) {
                console.error('下载文件时出错:', err);
            }
            if (fs.existsSync(outputPath)) {
                fs.unlinkSync(outputPath); // 下载后删除生成的填充文档
            }
        });

    } catch (error) {
        console.error('处理请求时发生内部错误:', error);
        // 完善错误信息，特别是 Docxtemplater 渲染错误
        if (error.properties && error.properties.errors) {
            const templateErrors = error.properties.errors.map(e => `[${e.id}]: ${e.message} (Context: ${JSON.stringify(e.properties.context)})`).join(', ');
            return res.status(400).send({ message: `文档渲染时出错。详情: ${templateErrors}` });
        }
        res.status(500).send({ message: '服务器内部错误。' + error.message });
    } finally {
        // 清理 multer 上传的临时文件
        if (wordFilePath && fs.existsSync(wordFilePath)) {
            fs.unlinkSync(wordFilePath);
        }
        if (jsonFilePath && fs.existsSync(jsonFilePath)) {
            fs.unlinkSync(jsonFilePath);
        }
    }
});

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.use((err, req, res, next) => {
    if (err instanceof multer.MulterError) {
        return res.status(400).send({ message: `文件上传错误 (Multer): ${err.message}` });
    } else if (err) {
        // 如果是 Error 实例并且有 message，就用它
        const message = err.message || '未知错误';
        return res.status(400).send({ message: message });
    }
    next();
});

app.listen(port, () => {
    console.log(`服务器正在 http://localhost:${port} 上运行`);
});