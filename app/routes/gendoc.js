const express = require('express');
const router = express.Router();

const path = require('path');
const fs = require('fs');
const os = require('os');
const { spawn } = require('child_process');
const Excel = require('exceljs');


// template เดิม (ยังอยู่ไปก่อน)
const TEMPLATE_XLSX = path.join(__dirname, '..', 'templates', 'template.xlsx');

// ใช้เก็บไฟล์ template ที่ upload
function getTemplatePathFromId(templateId) {
    return path.join(os.tmpdir(), `${templateId}.xlsx`);
}
// ===== LibreOffice =====
const SOFFICE = process.platform === 'win32'
    ? 'C:\\Program Files\\LibreOffice\\program\\soffice.exe'
    : 'soffice';

// -------------------------------------------------------
// Utility
// -------------------------------------------------------
function get(obj, pathStr) {
    const normalized = pathStr.replace(/\[(\d+)\]/g, '.$1');
    return normalized.split('.').reduce((o, k) => (o ? o[k] ?? '' : ''), obj);
}

function replaceTokensInCell(cell, data) {
    if (typeof cell.value !== 'string') return;

    let hasArrayToken = false;

    cell.value = cell.value.replace(/{{\s*([\w.[\]0-9]+)\s*}}/g, (_, k) => {
        if (/\[\d+\]/.test(k)) hasArrayToken = true;
        return String(get(data, k));
    });

    if (hasArrayToken) {
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
        };
    }
}

function mapOrientation(input) {
    const v = String(input || '').toLowerCase();
    if (v === 'p' || v === 'portrait') return 'portrait';
    if (v === 'l' || v === 'landscape') return 'landscape';
    return 'portrait';
}

// -------------------------------------------------------
// ขยาย rows แบบ array
// -------------------------------------------------------
function expandArrayRows(ws, data) {
    for (let rowNum = ws.rowCount; rowNum >= 1; rowNum--) {
        const row = ws.getRow(rowNum);
        let arrayName = null;

        row.eachCell(cell => {
            if (typeof cell.value !== 'string') return;
            const m = cell.value.match(/{{\s*(\w+)\[0\]\.[\w]+/);
            if (m) arrayName = m[1];
        });

        if (!arrayName) continue;

        const arr = data[arrayName];
        if (!Array.isArray(arr) || arr.length <= 1) continue;

        const templateValues = row.values.slice();

        for (let i = 1; i < arr.length; i++) {
            const newRow = ws.insertRow(rowNum + i, []);
            newRow.values = templateValues;

            newRow.eachCell(cell => {
                if (typeof cell.value === 'string') {
                    cell.value = cell.value.replace(/\[0\]/g, `[${i}]`);
                }
            });
        }
    }
}

// -------------------------------------------------------
// render excel
// -------------------------------------------------------
async function fillXlsx(tplPath, data) {
    const wb = new Excel.Workbook();
    await wb.xlsx.readFile(tplPath);

    const defaultOpt = {
        paperSize: 'A4',
        orientation: 'portrait',
        margin: 0,
        pageNumber: true,
        pageNumberPosition: 'bottom-center',
        repeatHeaderRows: '',
    };

    const opt = { ...defaultOpt, ...(data.__options || {}) };

    const paperMap = { A3: 8, A4: 9, A5: 11, Letter: 1 };

    wb.eachSheet(ws => {
        ws.pageSetup.paperSize = paperMap[opt.paperSize] || 9;
        ws.pageSetup.orientation = mapOrientation(opt.orientation);

        const m = opt.margin ?? 0;
        ws.pageSetup.margins = {
            left: m,
            right: m,
            top: m,
            bottom: m,
            header: Math.max(m, 0.3),
            footer: Math.max(m, 1),
        };

        if (opt.repeatHeaderRows) {
            ws.pageSetup.printTitlesRow = opt.repeatHeaderRows;
        }

        if (opt.pageNumber) {
            const label = 'หน้า &P / &N';
            const pos = opt.pageNumberPosition || 'bottom-center';

            ws.headerFooter = ws.headerFooter || {};
            const map = {
                'top-left': 'oddHeader',
                'top-center': 'oddHeader',
                'top-right': 'oddHeader',
                'bottom-left': 'oddFooter',
                'bottom-center': 'oddFooter',
                'bottom-right': 'oddFooter',
            };

            const side = pos.includes('top') ? pos : pos;
            const field = map[pos];
            const align = pos.includes('left') ? '&L' : pos.includes('right') ? '&R' : '&C';

            ws.headerFooter[field] = `${align}${label}`;
        }

        expandArrayRows(ws, data);

        ws.eachRow(row => row.eachCell(cell => replaceTokensInCell(cell, data)));
    });

    const outXlsx = path.join(os.tmpdir(), `filled_${Date.now()}.xlsx`);
    await wb.xlsx.writeFile(outXlsx);
    return outXlsx;
}

// -------------------------------------------------------
// convert to pdf
// -------------------------------------------------------
function convertToPdf(xlsxPath) {
    const outDir = path.dirname(xlsxPath);
    return new Promise((resolve, reject) => {
        const args = [
            '--headless',
            '--nologo',
            '--nolockcheck',
            '--nodefault',
            '--norestore',
            '--convert-to',
            'pdf',
            '--outdir',
            outDir,
            xlsxPath,
        ];
        const p = spawn(SOFFICE, args, { stdio: 'ignore' });
        p.on('exit', code => {
            if (code !== 0) return reject(new Error('soffice exit ' + code));
            const base = path.basename(xlsxPath).replace(/\.[^.]+$/, '');
            resolve(path.join(outDir, `${base}.pdf`));
        });
    });
}

// -------------------------------------------------------
// Schema builder
// -------------------------------------------------------
async function buildSchemaFromTemplate(tplPath) {
    const wb = new Excel.Workbook();
    await wb.xlsx.readFile(tplPath);

    const schema = {};
    function addKey(schema, keyPath) {
        const m = keyPath.match(/^(\w+)\[(\d+)\]\.(\w+)$/);
        if (m) {
            const [, arrName, , fieldName] = m;
            if (!schema[arrName]) schema[arrName] = [];
            if (!schema[arrName][0]) schema[arrName][0] = {};
            schema[arrName][0][fieldName] = '';
        } else {
            if (!schema[keyPath]) schema[keyPath] = '';
        }
    }

    wb.eachSheet(ws => {
        ws.eachRow(row =>
            row.eachCell(cell => {
                if (typeof cell.value !== 'string') return;
                const re = /{{\s*([\w.[\]0-9]+)\s*}}/g;
                let m;
                while ((m = re.exec(cell.value))) addKey(schema, m[1]);
            })
        );
    });

    return schema;
}

// -------------------------------------------------------
// Routes
// -------------------------------------------------------

router.get('/health', (req, res) => {
    res.json({ ok: true });
});

router.get('/schema', async (req, res) => {
    try {
        /* if (!fs.existsSync(TEMPLATE_XLSX)) {
            return res.status(500).json({ error: 'Template missing' });
        }

        const schema = await buildSchemaFromTemplate(TEMPLATE_XLSX); */

        const finalSchema = {
            "__options": {
                paperSize: "A4",
                orientation: "p",
                margin: 0,
                /* pageNumber: true, */
                pageNumberPosition: "bottom-center",
                repeatHeaderRows: "",
            },
            data: {}
        };

        res.json(finalSchema);
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

router.post('/render', async (req, res) => {
    try {
        const raw = req.body || {};
        const data = Array.isArray(raw) ? { items: raw } : raw;

        if (!data.__templateId) {
            return res.status(400).json({ error: 'No template selected (__templateId missing)' });
        }

        const tplPath = getTemplatePathFromId(data.__templateId);
        if (!fs.existsSync(tplPath)) {
            return res.status(400).json({ error: 'Template file not found' });
        }

        const xlsx = await fillXlsx(tplPath, data);
        const pdf = await convertToPdf(xlsx);

        res.setHeader('Content-Type', 'application/pdf');
        fs.createReadStream(pdf).pipe(res).on('close', () => {
            fs.unlink(xlsx, () => { });
            fs.unlink(pdf, () => { });
            // ถ้าอยากลบ template ที่ upload แล้วใช้ครั้งเดียวก็ลบตรงนี้ได้ด้วย
            // if (data.__templateId) fs.unlink(tplPath, () => {});
        });
    } catch (e) {
        console.error(e);
        res.status(500).json({ error: e.message });
    }
});


// POST /api/gendoc/schema/upload
// รับไฟล์ Excel แล้วคืน schema + templateId
router.post('/schema/upload', async (req, res) => {
    try {
        // ถ้าใช้ express-fileupload ใน app.js มันจะยัดไฟล์ไว้ที่นี่
        if (!req.files || !req.files.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        const file = req.files.file;          // field name = 'file' ตรงกับ Dragger.name
        const buffer = file.data;             // เป็น Buffer ของไฟล์ Excel

        // สร้าง temp path (ตาม templateId)
        const templateId = `tpl_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
        const tplPath = getTemplatePathFromId(templateId);

        // เซฟไฟล์ลง temp
        fs.writeFileSync(tplPath, buffer);

        // ใช้ function เดิมสร้าง schema จากไฟล์นี้
        const schema = await buildSchemaFromTemplate(tplPath);

        const finalSchema = {
            __templateId: templateId,
            __options: {
                paperSize: 'A4',
                orientation: 'p',
                margin: 0,
                pageNumber: true,
                pageNumberPosition: 'bottom-center',
                repeatHeaderRows: '',
            },
            ...schema,
        };

        res.json(finalSchema);
    } catch (e) {
        console.error(e);
        res.status(500).json({ error: e.message });
    }
});



// -------------------------------------------------------

module.exports = router;
