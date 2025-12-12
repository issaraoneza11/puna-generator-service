const express = require('express');
const router = express.Router();
const databaseContextPg = require("database-context-pg");
const connectionSetting = require("../dbconnect");
const connectionConfig = connectionSetting.config;
const condb = new databaseContextPg(connectionConfig);
const path = require('path');
const fs = require('fs');
const os = require('os');
const { spawn } = require('child_process');
const Excel = require('exceljs');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const { createCanvas } = require('canvas');

async function checkPermissionUrl(url) {
    let arr_permission = [
        'http://localhost:3000/',
        'http://localhost:5173/'
    ];
    let check = await condb.clientQuery(`SELECT pu_id, pu_url, pu_key, pu_is_active
	FROM public.permission_url WHERE pu_is_active = true AND pu_url = $1`, [url]);

    /*  console.log('check',check.rows); */


    if (check.rows.length > 0) {
        return true;
    } else {
        return false;
    }
}



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
// Excel format detection / legacy XLS convert
// -------------------------------------------------------
function detectExcelFormat(buffer) {
    if (!buffer || buffer.length < 4) return 'unknown';

    // XLSX (OOXML) เป็น ZIP, ขึ้นต้นด้วย 'PK'
    if (buffer[0] === 0x50 && buffer[1] === 0x4B) {
        return 'xlsx';
    }

    // XLS เก่า (CFB/BIFF) magic: D0 CF 11 E0 ...
    if (
        buffer[0] === 0xD0 && buffer[1] === 0xCF &&
        buffer[2] === 0x11 && buffer[3] === 0xE0
    ) {
        return 'xls';
    }

    return 'unknown';
}

/**
 * ให้แน่ใจว่าเราได้เป็น XLSX จริง ๆ
 * - ถ้า buffer เป็น XLSX อยู่แล้ว → คืนกลับตรง ๆ
 * - ถ้าเป็น XLS เก่า → แปลงด้วย LibreOffice (soffice) ก่อน
 */
async function ensureXlsxFromBuffer(buffer) {
    const format = detectExcelFormat(buffer);

    if (format === 'xlsx') {
        return { buffer, convertedFromLegacyXls: false };
    }

    if (format === 'xls') {
        const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlsconv_'));
        const srcPath = path.join(tmpDir, 'input.xls');
        const outPath = path.join(tmpDir, 'input.xlsx');

        fs.writeFileSync(srcPath, buffer);

        await new Promise((resolve, reject) => {
            const args = [
                '--headless',
                '--nologo',
                '--nolockcheck',
                '--nodefault',
                '--norestore',
                '--convert-to',
                'xlsx',
                '--outdir',
                tmpDir,
                srcPath,
            ];
            const p = spawn(SOFFICE, args, { stdio: 'ignore' });
            p.on('exit', code => {
                if (code !== 0) return reject(new Error('soffice exit ' + code));
                resolve();
            });
        });

        if (!fs.existsSync(outPath)) {
            throw new Error('XLS→XLSX conversion failed');
        }

        const xlsxBuffer = fs.readFileSync(outPath);
        return { buffer: xlsxBuffer, convertedFromLegacyXls: true };
    }

    // ถ้าไม่ใช่ xls/xlsx → ไม่รองรับ
    throw new Error('Unsupported Excel format');
}

// -------------------------------------------------------
// Utility
// -------------------------------------------------------
function get(obj, pathStr) {
    const normalized = pathStr.replace(/\[(\d+)\]/g, '.$1');
    return normalized.split('.').reduce((o, k) => (o ? o[k] ?? '' : ''), obj);
}
function normalizePos(pos) {
    const map = {
        tl: 'top-left',
        tc: 'top-center',
        tr: 'top-right',
        bl: 'bottom-left',
        bc: 'bottom-center',
        br: 'bottom-right',
    };

    return map[pos] || pos; // ถ้าไม่ได้ใส่แบบใหม่ ก็ปล่อยของเดิมไว้
}
function applyDefaultStyle(cell) {
    const oldFont = cell.font || {};
    const oldAlign = cell.alignment || {};

    cell.font = {
        ...oldFont,
        // ใช้ฟอนต์เดิมถ้ามี ถ้าไม่มีค่อยใช้ TH Sarabun
        name: oldFont.name ?? 'TH SarabunPSK',
        // ❌ ไม่ยุ่ง size เลย ปล่อยตาม template / fs:
        color: oldFont.color ?? { argb: 'FF000000' },
        bold: oldFont.bold ?? false,
        italic: oldFont.italic ?? false,
        underline: oldFont.underline ?? false,
    };

    cell.alignment = {
        ...oldAlign,
        horizontal: oldAlign.horizontal ?? 'left',
        vertical: oldAlign.vertical ?? 'top',
        wrapText: oldAlign.wrapText ?? false,
    };
}



function applyInlineStyle(cell, styleTokens) {
    if (!styleTokens || styleTokens.length === 0) return;

    let wrap = null;
    let bold = null;
    let italic = null;
    let underline = null;
    let hAlign = null;
    let vAlign = null;
    let color = null;
    let fontSize = null;   // 👈 เพิ่มตัวแปรเก็บ size

    for (const raw of styleTokens) {
        const t = raw.trim();
        if (!t) continue;

        const tk = t.toLowerCase();

        // wrapText
        if (tk === 'w') { wrap = true; continue; }
        if (tk === 'nw') { wrap = false; continue; }

        // bold / italic / underline
        if (tk === 'b') { bold = true; continue; }
        if (tk === 'nb') { bold = false; continue; }
        if (tk === 'i') { italic = true; continue; }
        if (tk === 'ni') { italic = false; continue; }
        if (tk === 'u') { underline = true; continue; }
        if (tk === 'nu') { underline = false; continue; }

        // horizontal align
        if (tk === 'hl') { hAlign = 'left'; continue; }
        if (tk === 'hc') { hAlign = 'center'; continue; }
        if (tk === 'hr') { hAlign = 'right'; continue; }

        // vertical align
        if (tk === 'vt') { vAlign = 'top'; continue; }
        if (tk === 'vm') { vAlign = 'middle'; continue; }
        if (tk === 'vb') { vAlign = 'bottom'; continue; }

        // color
        if (/^#[0-9a-f]{3}([0-9a-f]{3})?$/i.test(tk)) {
            const hex = tk.slice(1).toUpperCase();
            const fullHex = hex.length === 3
                ? hex.split('').map(ch => ch + ch).join('')
                : hex;
            color = fullHex;
            continue;
        }

        // 🔹 font size: fs:14 หรือ FS:18
        if (tk.startsWith('fs:')) {
            const n = Number(tk.slice(3));
            if (Number.isFinite(n) && n > 0) {
                fontSize = n;
            }
            continue;
        }

        // token อื่นไม่รู้จัก → ข้าม
    }

    // apply font รวม size ด้วย
    if (bold !== null || italic !== null || underline !== null || color || fontSize !== null) {
        const oldFont = cell.font || {};
        cell.font = {
            ...oldFont,
            ...(bold !== null ? { bold } : {}),
            ...(italic !== null ? { italic } : {}),
            ...(underline !== null ? { underline } : {}),
            ...(color ? { color: { argb: 'FF' + color } } : {}),
            ...(fontSize !== null ? { size: fontSize } : {}),
        };
    }

    // apply alignment
    if (wrap !== null || hAlign || vAlign) {
        const oldAlign = cell.alignment || {};
        cell.alignment = {
            ...oldAlign,
            ...(wrap !== null ? { wrapText: wrap } : {}),
            ...(hAlign ? { horizontal: hAlign } : {}),
            ...(vAlign ? { vertical: vAlign } : {}),
        };
    }
}

function softWrapLabelValueCell(cell, colNumber) {
    if (!cell || typeof cell.value !== 'string') return;

    const text = cell.value;

    // จับ pattern "label: value"
    const m = text.match(/^(.*?:\s*)(.+)$/);
    if (!m) return;

    const prefix = m[1];
    const rest = m[2];

    if (!rest) return;

    // ความกว้างคอลัมน์ (หน่วยของ exceljs) → แปลงคร่าว ๆ เป็น pixel
    const ws = cell.worksheet;
    const col = ws.getColumn(colNumber);
    const colWidth = col.width || 10;
    const colPx = colWidth * 7; // ประมาณการ: 1 unit ~ 7px

    const font = cell.font || {};
    const fontSize = Number(font.size) || 16;
    const fontName = font.name || 'TH SarabunPSK';

    const lines = [];
    let current = prefix;

    for (const ch of rest) {
        const candidate = current + ch;
        const w = measureTextWidthPx(candidate, fontSize, fontName);

        // ถ้าเกินความกว้าง cell แล้ว และ current มีตัวมากกว่าพวก prefix → ตัดบรรทัด
        if (w > colPx && current !== prefix) {
            lines.push(current);
            // บรรทัดใหม่: indent ด้วยช่องว่างเท่ากับ prefix
            const indent = ' '.repeat(prefix.length);
            current = indent + ch;
        } else {
            current = candidate;
        }
    }
    if (current) {
        lines.push(current);
    }

    cell.value = lines.join('\n');

    const align = cell.alignment || {};
    cell.alignment = {
        ...align,
        wrapText: true,
        vertical: align.vertical || 'top',
    };
}

function normalizeKeyForStyle(path) {
    // แปลง goog[0].no, goog[1].no → goog[].no ให้เป็น key เดียวกัน
    return String(path || '').replace(/\[\d+\]/g, '[]');
}
function replaceTokensInCell(cell, data, defaultStyleByKey) {
    if (typeof cell.value !== 'string') return;

    let hasArrayToken = false;

    // state ต่อ 1 cell
    let mainKeyPath = null;
    let hasExplicitStyle = false;

    cell.value = cell.value.replace(/{{\s*([^{}]+?)\s*}}/g, (_, inner) => {
        const tokens = splitPlaceholder(inner);   // 👈 ใช้ฟังก์ชันใหม่
        if (tokens.length === 0) return '';

        const key = tokens[0];
        const styleTokens = tokens.slice(1);

        // 1) style
        if (key.toLowerCase() === 'style') {
            hasExplicitStyle = true;
            if (styleTokens.length > 0) {
                applyInlineStyle(cell, styleTokens);

                if (mainKeyPath) {
                    const norm = normalizeKeyForStyle(mainKeyPath);
                    defaultStyleByKey[norm] = styleTokens.slice();
                }
            }
            return '';
        }

        // 2) fx
        if (key.toLowerCase() === 'fx') {
            const result = evalFxFormula(styleTokens, data);
            if (styleTokens.some(t => /\[\d+\]/.test(t))) {
                hasArrayToken = true;
            }
            return String(result ?? '');
        }

        // 3) ปกติ: data path
        const keyPath = key;
        mainKeyPath = mainKeyPath || keyPath;

        if (/\[\d+\]/.test(keyPath)) hasArrayToken = true;

        const v = String(get(data, keyPath));
        return v;
    });

    if (hasArrayToken) {
        const oldAlign = cell.alignment || {};

        cell.alignment = {
            ...oldAlign,
            // เดิมเขียน wrapText: true,
            // แก้เป็นเซ็ตเฉพาะตอนยังไม่ได้ตั้งค่า
            ...(oldAlign.wrapText === undefined ? { wrapText: true } : {}),
            vertical: oldAlign.vertical || 'top',
        };

        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
        };
    }



    // cell ปกติ
    // cell ปกติ
    if (mainKeyPath) {
        applyDefaultStyle(cell);

        const align = cell.alignment || {};
        cell.alignment = {
            ...align,
            vertical: align.vertical || 'top',
            // ไม่ไปยุ่ง wrapText ตรงนี้ เพื่อให้ได้ค่าจาก template หรือ style (w / nw)
        };
    }

    if (mainKeyPath && !hasExplicitStyle) {
        const norm = normalizeKeyForStyle(mainKeyPath);
        const defTokens = defaultStyleByKey[norm];
        if (defTokens && defTokens.length > 0) {
            applyInlineStyle(cell, defTokens);
        }
    }


}



function isQuoted(str) {
    return /^(['"]).*\1$/.test(str);
}

function stripQuotes(str) {
    const m = str.match(/^(['"])(.*)\1$/);
    return m ? m[2] : str;
}
function splitPlaceholder(inner) {
    if (!inner) return [];

    // แปลง \r\n, \n ให้เป็น space ธรรมดา
    let s = String(inner)
        .replace(/\r\n/g, '\n')
        .replace(/\n+/g, ' ');

    // split ตาม | แล้ว trim ช่องว่างออก
    let parts = s.split('|').map(p => p.trim()).filter(Boolean);
    if (parts.length === 0) return [];

    // กรณีเคาะบรรทัดผิด เช่น "style\n  ni | hl..."
    // จะได้ "style ni" → ถ้า token แรกมี space และไม่ใช่ string แบบใส่ quote
    // ให้แตกเพิ่มตาม space เป็นหลาย token
    if (!isQuoted(parts[0]) && /\s/.test(parts[0])) {
        const firstPieces = parts[0].split(/\s+/).filter(Boolean);
        parts = [...firstPieces, ...parts.slice(1)];
    }

    return parts;
}

function resolveTokenValue(token, data) {
    if (token == null) return '';

    const trimmed = token.trim();
    if (!trimmed) return '';

    // ถ้าใส่ "..." หรือ '...' → เป็น literal ตรง ๆ
    if (isQuoted(trimmed)) {
        return stripQuotes(trimmed);
    }

    // อย่างอื่น treat เป็น key path
    return get(data, trimmed);
}
function setByPath(obj, pathStr, value) {
    if (!pathStr) return;
    const normalized = pathStr.replace(/\[(\d+)\]/g, '.$1');
    const parts = normalized.split('.').filter(Boolean);
    if (!parts.length) return;

    let o = obj;
    for (let i = 0; i < parts.length - 1; i++) {
        const k = parts[i];
        if (o[k] == null || typeof o[k] !== 'object') {
            o[k] = {};
        }
        o = o[k];
    }
    o[parts[parts.length - 1]] = value;
}

function evalFxFormula(tokens, data) {
    if (!tokens || tokens.length === 0) return '';

    // แยก core tokens กับ as:alias ออกจากกันก่อน
    let asPath = null;
    const coreTokens = [];

    for (const t of tokens) {
        const s = String(t || '').trim();
        if (!s) continue;

        if (s.toLowerCase().startsWith('as:')) {
            const alias = s.slice(3).trim();
            if (alias) asPath = alias;
        } else {
            coreTokens.push(s);
        }
    }

    if (coreTokens.length === 0) return '';

    const cmd = String(coreTokens[0] || '').toLowerCase();
    const args = coreTokens.slice(1);

    let result = '';

    // -----------------------------
    // fx|sum|qty|price|fee
    // -----------------------------
    if (cmd === 'sum') {
        let total = 0;
        let any = false;

        for (const path of args) {
            const v = resolveTokenValue(path, data);
            const n = Number(v);
            if (Number.isFinite(n)) {
                total += n;
                any = true;
            }
        }
        result = any ? total : '';
    }

    // -----------------------------
    // fx|if|qty|==|0|"ไม่มี"| "มี"
    // หรือ  fx|if|qty|0|"ไม่มี"| "มี"  (op default ==)
    // -----------------------------
    else if (cmd === 'if') {
        if (args.length >= 1) {
            const leftToken = args[0];
            let op = args[1];
            let rightToken;
            let thenToken;
            let elseToken;

            if (['==', '!=', '>', '>=', '<', '<='].includes(op)) {
                rightToken = args[2];
                thenToken = args[3];
                elseToken = args[4];
            } else {
                rightToken = op;
                op = '==';
                thenToken = args[2];
                elseToken = args[3];
            }

            const leftValRaw = resolveTokenValue(leftToken, data);
            const rightValRaw = resolveTokenValue(rightToken, data);

            const leftNum = Number(leftValRaw);
            const rightNum = Number(rightValRaw);
            const bothNum = Number.isFinite(leftNum) && Number.isFinite(rightNum);

            let cond = false;

            if (bothNum) {
                switch (op) {
                    case '==': cond = leftNum === rightNum; break;
                    case '!=': cond = leftNum !== rightNum; break;
                    case '>': cond = leftNum > rightNum; break;
                    case '>=': cond = leftNum >= rightNum; break;
                    case '<': cond = leftNum < rightNum; break;
                    case '<=': cond = leftNum <= rightNum; break;
                }
            } else {
                const L = String(leftValRaw ?? '');
                const R = String(rightValRaw ?? '');
                switch (op) {
                    case '==': cond = L === R; break;
                    case '!=': cond = L !== R; break;
                }
            }

            const chosen = cond ? thenToken : elseToken;
            result = chosen == null ? '' : resolveTokenValue(chosen, data);
        }
    }

    // ถ้ายังไม่รู้จักสูตรอื่น → ปล่อย result เป็น '' ไป

    // ถ้ามี as:xxx → เขียนค่าลง data ด้วย
    if (asPath) {
        setByPath(data, asPath, result);
    }

    return result;
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

        // หา array key จาก row นี้ เช่น goog[0].no
        row.eachCell(cell => {
            if (typeof cell.value !== 'string') return;
            const m = cell.value.match(/{{\s*([^{}]+?)\s*}}/);
            if (!m) return;

            const inner = m[1];
            const parts = inner.split('|').map(s => s.trim()).filter(Boolean);
            if (!parts.length) return;

            const key = parts[0];
            const mm = key.match(/^(\w+)\[0\]\./);
            if (mm) arrayName = mm[1];
        });

        if (!arrayName) continue;

        const arr = data[arrayName];
        if (!Array.isArray(arr) || arr.length <= 1) continue;

        // เก็บค่า template row ไว้ให้ชัด ๆ
        const templateRow = ws.getRow(rowNum);
        const templateValues = templateRow.values.slice();
        const templateHeight = templateRow.height;

        // เก็บ style ของแต่ละคอลัมน์แบบ deep copy
        const templateStyles = {};
        templateRow.eachCell({ includeEmpty: true }, (tmplCell, col) => {
            templateStyles[col] = JSON.parse(JSON.stringify(tmplCell.style || {}));
        });

        // สร้าง row เพิ่มตามจำนวน array
        for (let i = 1; i < arr.length; i++) {
            const newRow = ws.insertRow(rowNum + i, []);
            newRow.values = templateValues.slice();

            if (templateHeight != null) {
                newRow.height = templateHeight;
            }

            templateRow.eachCell({ includeEmpty: true }, (tmplCell, col) => {
                const cell = newRow.getCell(col);

                // clone style จาก template เป๊ะ ๆ
                cell.style = JSON.parse(JSON.stringify(templateStyles[col] || {}));

                // แก้ [0] → [i] เฉพาะ cell ที่เป็น string
                if (typeof cell.value === 'string') {
                    cell.value = cell.value.replace(/\[0\]/g, `[${i}]`);
                }
            });
        }
    }
}

const measureCanvas = createCanvas(1000, 100);
const measureCtx = measureCanvas.getContext('2d');
function measureTextWidthPx(text, fontSize, fontName = 'TH SarabunPSK') {
    if (!text) return 0;
    const size = Number(fontSize) || 16;
    measureCtx.font = `${size}pt "${fontName}"`;
    const metrics = measureCtx.measureText(text);
    return metrics.width || 0;
}
const IS_LINUX = process.platform === 'linux';

function autoAdjustRowHeightByWrap(ws) {
    ws.eachRow((row) => {
        let hasWrap = false;
        let maxLines = 1;
        let maxFontSize = 0;

        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const align = cell.alignment || {};
            if (!align.wrapText) return;   // ไม่ได้เปิด wrap ก็ข้าม

            hasWrap = true;

            const text = (typeof cell.value === 'string') ? cell.value : '';
            if (!text) return;

            const font = cell.font || {};
            const fontSize = Number(font.size) || 11;
            if (fontSize > maxFontSize) maxFontSize = fontSize;

            const paragraphs = text.split(/\r?\n/);
            const col = ws.getColumn(colNumber);
            const colWidth = col.width || 10;

            // กะ ๆ ว่า 1 หน่วย column ~ 7 px
            const colPx = colWidth * 7;

            let totalLines = 0;
            for (const p of paragraphs) {
                if (!p) {
                    totalLines += 1;
                    continue;
                }
                const wPx = measureTextWidthPx(p, fontSize, font.name || 'TH SarabunPSK');
                const linesForPara = Math.max(1, Math.ceil(wPx / colPx));
                totalLines += linesForPara;
            }

            if (totalLines > maxLines) maxLines = totalLines;
        });

        // แถวนี้ไม่มี cell ไหนเปิด wrap ก็ไม่ยุ่ง
        if (!hasWrap) return;

        if (!maxFontSize) maxFontSize = 11;

        const lineHeight = maxFontSize * 1.15;
        const padding = 4;
        let target = lineHeight * maxLines + padding;

        if (IS_LINUX) {
            target *= 1.05;   // เผื่อ Linux นิดนึง
        }

        row.height = target;
    });
}


function smartWrapLabelValueCell(ws, cell, colNumber) {
    if (!cell || typeof cell.value !== 'string') return;

    const align = cell.alignment || {};

    // ❗ ต้องเป็น true เท่านั้นถึงจะจัด wrap ให้
    if (!align.wrapText) return;

    const text = cell.value;

    const m = text.match(/^(.*?:\s+)(.+)$/);
    if (!m) return;

    const prefix = m[1];
    const rest = m[2];

    const col = ws.getColumn(colNumber);
    const colWidth = col.width || 10;
    const colPx = colWidth * 7;

    const font = cell.font || {};
    const fontSize = Number(font.size) || 11;
    const fontName = font.name || 'TH SarabunPSK';

    if (measureTextWidthPx(text, fontSize, fontName) <= colPx) return;

    const lines = [];
    let current = prefix;

    for (const ch of rest) {
        const candidate = current + ch;
        const w = measureTextWidthPx(candidate, fontSize, fontName);

        if (w > colPx && current !== '') {
            lines.push(current.trimEnd());
            current = ch;
        } else {
            current += ch;
        }
    }
    if (current) lines.push(current.trimEnd());

    cell.value = lines.join('\n');

    const oldAlign = cell.alignment || {};
    cell.alignment = {
        ...oldAlign,
        wrapText: true,
        vertical: oldAlign.vertical || 'top',
    };
}





// -------------------------------------------------------
// render excel
// -------------------------------------------------------
async function fillXlsx(tplPath, data) {

    const wb = new Excel.Workbook();
    await wb.xlsx.readFile(tplPath);
    const defaultStyleByKey = {};
    const defaultOpt = {
        paperSize: 'A4',
        orientation: 'portrait',
        margin: 0, // รองรับค่า default เป็นตัวเลขอยู่เหมือนเดิม
        pageNumber: true,
        pageNumberPosition: 'bottom-center',
        repeatHeaderRows: '',
        autoScaleToFitWidth: false,
    };

    const opt = { ...defaultOpt, ...(data.__options || {}) };

    const paperMap = { A3: 8, A4: 9, A5: 11, Letter: 1 };

    function getPaperWidthInches(paperSize, orientation) {
        // default A4
        let w = 8.27;
        let h = 11.69;

        if (paperSize === 'A3') {
            w = 11.69;
            h = 16.54;
        } else if (paperSize === 'Letter') {
            w = 8.5;
            h = 11;
        } else if (paperSize === 'A5') {
            w = 5.83;
            h = 8.27;
        }

        const ori = mapOrientation(orientation || 'portrait');
        if (ori === 'landscape') {
            return h; // สลับกว้าง/ยาว
        }
        return w;
    }

    function autoScaleToFitWidth(ws, opt, marginLeft, marginRight) {
        const paperWidthInches = getPaperWidthInches(opt.paperSize || 'A4', opt.orientation || 'portrait');
        const printableWidthInches = Math.max(
            0.1,
            paperWidthInches - (marginLeft + marginRight)
        );

        // หา column ที่ใช้จริง
        const usedCols = new Set();
        ws.eachRow({ includeEmpty: false }, row => {
            row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                usedCols.add(colNumber);
            });
        });

        if (usedCols.size === 0) return;

        // รวมความกว้างทุกคอลัมน์ (หน่วยนิ้ว โดยประมาณ)
        let totalWidthInches = 0;
        for (const colNumber of usedCols) {
            const col = ws.getColumn(colNumber);
            const colWidth = col.width || 8.43; // default Excel ประมาณนี้
            const colInches = (colWidth * 7) / 96; // 1 หน่วย = ~7px, 96px = 1 inch
            totalWidthInches += colInches;
        }

        if (totalWidthInches <= 0) return;

        if (opt.forceSinglePage) {
            ws.pageSetup.fitToPage = true;
            ws.pageSetup.fitToWidth = 1;
            ws.pageSetup.fitToHeight = 1;
            ws.pageSetup.scale = undefined; // ปล่อยให้ Excel คิดเอง
            return;
        }
        /*  const scaleFloat = (printableWidthInches / totalWidthInches) * 100;
         const scale = Math.floor(Math.min(100, scaleFloat));
 
         // ถ้าเกิน 100 แปลว่ากว้างพอแล้ว ไม่ต้องขยาย
         if (scale < 100 && scale > 10) {
             ws.pageSetup.fitToPage = false;
             ws.pageSetup.fitToWidth = undefined;
             ws.pageSetup.fitToHeight = undefined;
             ws.pageSetup.scale = undefined;
         } */
    }

    wb.eachSheet(ws => {
        ws.pageSetup.paperSize = paperMap[opt.paperSize] || 9;
        ws.pageSetup.orientation = mapOrientation(opt.orientation);

        // ----------------------
        // รองรับ margin ทั้งแบบตัวเดียว และแบบ object 4 ด้าน
        // ----------------------
        const margin = opt.margin ?? 0;

        let marginLeft, marginRight, marginTop, marginBottom;

        if (margin && typeof margin === 'object') {
            marginLeft = Number(margin.left ?? 0) || 0;
            marginRight = Number(margin.right ?? 0) || 0;
            marginTop = Number(margin.top ?? 0) || 0;
            marginBottom = Number(margin.bottom ?? 0) || 0;
        } else {
            const m = Number(margin) || 0;
            marginLeft = marginRight = marginTop = marginBottom = m;
        }

        // 🔹 auto ปรับ margin ตามตำแหน่งเลขหน้า (หน่วย = นิ้ว)
        if (opt.pageNumber) {
            const pos = normalizePos(opt.pageNumberPosition || 'bottom-center');
            const MIN_TOP = 0.4;     // ~1.8cm
            const MIN_BOTTOM = 0.4;  // ~1.8cm

            if (pos.startsWith('top') && marginTop < MIN_TOP) {
                marginTop = MIN_TOP;
            }
            if (pos.startsWith('bottom') && marginBottom < MIN_BOTTOM) {
                marginBottom = MIN_BOTTOM;
            }
        }

        // ✅ ค่อยเซ็ต margin เข้า Excel หลังปรับเสร็จแล้ว
        ws.pageSetup.margins = {
            left: marginLeft,
            right: marginRight,
            top: marginTop,
            bottom: marginBottom,
            header: Math.max(marginTop, 0.3),
            footer: Math.max(marginBottom, 1),
        };

        if (opt.repeatHeaderRows) {
            ws.pageSetup.printTitlesRow = opt.repeatHeaderRows;
        }
        autoScaleToFitWidth(ws, opt, marginLeft, marginRight);

        expandArrayRows(ws, data);

        // 1) แทนค่า + style จาก token
        ws.eachRow(row => row.eachCell(cell => {
            replaceTokensInCell(cell, data, defaultStyleByKey);
        }));

        // 2) บังคับ wrap แบบ label: value ยาว ๆ เช่น "ชื่อลูกค้า: xxxxxx"
        ws.eachRow(row => row.eachCell((cell, colNumber) => {
            smartWrapLabelValueCell(ws, cell, colNumber);
        }));

        // 3) คำนวณ row height ใหม่ (เฉพาะ Linux)
        if (IS_LINUX) {
            autoAdjustRowHeightByWrap(ws);
        }

        // 4) บังคับฟอนต์ TH Sarabun ให้ทุก cell
        ws.eachRow(row => {
            row.eachCell(cell => {
                const oldFont = cell.font || {};
                cell.font = {
                    ...oldFont,
                    name: 'TH SarabunPSK',
                };
            });
        });


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

async function addPageNumbers(pdfPath, options = {}) {
    const bytes = fs.readFileSync(pdfPath);
    const pdfDoc = await PDFDocument.load(bytes);

    const fontkit = require('@pdf-lib/fontkit');
    pdfDoc.registerFontkit(fontkit);

    const fontPath = path.join(__dirname, '..', 'font', 'THSarabun.ttf');
    const fontBytes = fs.readFileSync(fontPath);
    const font = await pdfDoc.embedFont(fontBytes);

    const pages = pdfDoc.getPages();
    const pageCount = pages.length;

    const pos = normalizePos(options.pageNumberPosition || 'bottom-center');
    const fontSize = options.fontSize || 16;
    const margin = options.margin || 10;

    pages.forEach((page, index) => {
        const { width, height } = page.getSize();
        const text = `หน้า ${index + 1} / ${pageCount}`;
        const textWidth = font.widthOfTextAtSize(text, fontSize);

        let x = 0;
        let y = 0;

        switch (pos) {
            case 'top-left':
                x = margin;
                y = height - margin - fontSize;
                break;
            case 'top-center':
                x = (width - textWidth) / 2;
                y = height - margin - fontSize;
                break;
            case 'top-right':
                x = width - textWidth - margin;
                y = height - margin - fontSize;
                break;
            case 'bottom-left':
                x = margin;
                y = margin;
                break;
            case 'bottom-right':
                x = width - textWidth - margin;
                y = margin;
                break;
            case 'bottom-center':
            default:
                x = (width - textWidth) / 2;
                y = margin;
        }

        page.drawText(text, {
            x,
            y,
            size: fontSize,
            font,
            color: hexToRgb('#000000'),
        });
    });

    const outBytes = await pdfDoc.save();
    const outPath = pdfPath.replace(/\.pdf$/, '_paged.pdf');
    fs.writeFileSync(outPath, outBytes);
    return outPath;
}


function hexToRgb(hex) {
    const h = hex.replace('#', '');
    const bigint = parseInt(h, 16);
    const r = ((bigint >> 16) & 255) / 255;
    const g = ((bigint >> 8) & 255) / 255;
    const b = (bigint & 255) / 255;
    return rgb(r, g, b);
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

                const re = /{{\s*([^{}]+?)\s*}}/g;
                let m;
                while ((m = re.exec(cell.value))) {
                    const inner = m[1];
                    const tokens = splitPlaceholder(inner);
                    if (tokens.length === 0) continue;

                    const key = tokens[0];
                    const rest = tokens.slice(1);
                    if (key.toLowerCase().startsWith('as:')) {
                        continue;
                    }
                    // ข้าม style: {{style|...}}
                    if (key.toLowerCase() === 'style') {
                        continue;
                    }

                    // fx: ดึง key ที่เป็น path จาก arg
                    if (key.toLowerCase() === 'fx') {
                        /*       const cmd = (rest[0] || '').toLowerCase();
                              const args = rest.slice(1);
      
                              if (cmd === 'sum') {
                                  // fx|sum|qty|price|fee|as:goog[0].total
                                  for (let tok of args) {
                                      if (!tok) continue;
                                      tok = tok.trim();
                                      if (!tok) continue;
      
                                      // 👇 ข้าม alias
                                      if (tok.toLowerCase().startsWith('as:')) continue;
      
                                      // ถ้าไม่ใช่ literal string → ถือว่าเป็น key path
                                      if (!isQuoted(tok)) {
                                          addKey(schema, tok);
                                      }
                                  }
                              } else if (cmd === 'if') {
                                  // fx|if|qty|==|0|"ไม่มี"| "มี"|as:status
                                  if (args.length >= 1) {
                                      let left = args[0];
                                      if (left) {
                                          left = left.trim();
                                          if (left && !left.toLowerCase().startsWith('as:') && !isQuoted(left)) {
                                              addKey(schema, left);
                                          }
                                      }
                                      // ถ้าอนาคตอยากดึง key จาก then/else ก็เพิ่มเหมือนด้านบนได้
                                  }
                              } */

                        continue;
                    }



                    // ปกติ: {{ customer_name }} , {{ goog[0].qty }}
                    addKey(schema, key);
                }
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
        const referer = req.get('Referer');  // หรือ req.headers.referer
        let permission = await checkPermissionUrl(referer);
        if (!permission) {
            throw Error('No Permission');
        }
        /* if (!fs.existsSync(TEMPLATE_XLSX)) {
            return res.status(500).json({ error: 'Template missing' });
        }

        const schema = await buildSchemaFromTemplate(TEMPLATE_XLSX); */

        const finalSchema = {
            "__options": {
                paperSize: "A4",
                orientation: "p",
                margin: {
                    "left": 0,
                    "right": 0,
                    "top": 0,
                    "bottom": 0
                },
                /* pageNumber: true, */
                pageNumberPosition: "bottom-center",
                repeatHeaderRows: "",
                autoScaleToFitWidth: true
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
        const referer = req.get('Referer');  // หรือ req.headers.referer
        let permission = await checkPermissionUrl(referer);
        if (!permission) {
            throw Error('No Permission');
        }
        const raw = req.body || {};
        const data = Array.isArray(raw) ? { items: raw } : raw;

        const opt = {
            pageNumber: true,
            pageNumberPosition: 'bottom-center',
            ...(data.__options || {}),
        };
        if (!data.__templateId) {
            return res.status(400).json({ error: 'No template selected (__templateId missing)' });
        }

        const tplPath = getTemplatePathFromId(data.__templateId);
        if (!fs.existsSync(tplPath)) {
            return res.status(400).json({ error: 'Template file not found' });
        }

        const xlsx = await fillXlsx(tplPath, data);
        const pdf = await convertToPdf(xlsx);

        let finalPdf = pdf;

        // ถ้า opt.pageNumber !== false → เติมเลขหน้า
        if (opt.pageNumber !== false) {
            finalPdf = await addPageNumbers(pdf, {
                pageNumberPosition: opt.pageNumberPosition,
            });
        }

        res.setHeader('Content-Type', 'application/pdf');
        fs.createReadStream(finalPdf).pipe(res).on('close', () => {
            fs.unlink(xlsx, () => { });
            fs.unlink(pdf, () => { });
            if (finalPdf !== pdf) {
                fs.unlink(finalPdf, () => { });
            }
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
        const referer = req.get('Referer');  // หรือ req.headers.referer
        let permission = await checkPermissionUrl(referer);
        if (!permission) {
            throw Error('No Permission');
        }

        // ถ้าใช้ express-fileupload ใน app.js มันจะยัดไฟล์ไว้ที่นี่
        if (!req.files || !req.files.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        const file = req.files.file;          // field name = 'file' ตรงกับ Dragger.name
        const buffer = file.data;             // เป็น Buffer ของไฟล์ Excel (xls / xlsx ก็ได้)

        // ตรวจฟอร์แมต แล้ว ensure ให้เป็น XLSX ก่อน (รองรับ .xls แบบ auto-convert)
        const { buffer: xlsxBuffer, convertedFromLegacyXls } = await ensureXlsxFromBuffer(buffer);

        // สร้าง temp path (ตาม templateId)
        const templateId = `tpl_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
        const tplPath = getTemplatePathFromId(templateId);

        // เซฟไฟล์ลง temp เป็น .xlsx เสมอ
        fs.writeFileSync(tplPath, xlsxBuffer);

        // ใช้ function เดิมสร้าง schema จากไฟล์นี้
        const schema = await buildSchemaFromTemplate(tplPath);

        const finalSchema = {
            __templateId: templateId,
            __convertedFromLegacyXls: convertedFromLegacyXls, // เผื่อ UI จะเอาไปเตือน user
            __options: {
                paperSize: 'A4',
                orientation: 'p',
                margin: {
                    left: 0,
                    right: 0,
                    top: 0,
                    bottom: 0,
                },
                pageNumber: true,
                pageNumberPosition: 'bottom-center',
                repeatHeaderRows: '',
                forceSinglePage: true
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
