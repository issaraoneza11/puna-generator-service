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
const { PDFDocument, rgb } = require('pdf-lib');
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
const TEMPLATE_TTL_MS = 6 * 60 * 60 * 1000;
const TMP_DIR = path.join(os.tmpdir(), 'gendoc_templates');

function ensureTmpDir() {
    if (!fs.existsSync(TMP_DIR)) fs.mkdirSync(TMP_DIR, { recursive: true });
}

function cleanupOldTemplates(prefix = 'tpl_', ttlMs = TEMPLATE_TTL_MS) {
    ensureTmpDir();
    const now = Date.now();
    try {
        const files = fs.readdirSync(TMP_DIR);
        for (const f of files) {
            if (!f.startsWith(prefix) || !f.endsWith('.xlsx')) continue;
            const full = path.join(TMP_DIR, f);
            try {
                const st = fs.statSync(full);
                if (now - st.mtimeMs > ttlMs) fs.unlinkSync(full);
            } catch { }
        }
    } catch { }
}
function cleanupOldFiles(ttlMs = TEMPLATE_TTL_MS) {
    ensureTmpDir();
    const now = Date.now();

    try {
        const files = fs.readdirSync(TMP_DIR);
        for (const f of files) {
            // ลบเฉพาะไฟล์ที่เราสร้างเอง
            const okPrefix = f.startsWith('tpl_') || f.startsWith('filled_');
            const okExt = f.endsWith('.xlsx') || f.endsWith('.pdf');

            if (!okPrefix || !okExt) continue;

            const full = path.join(TMP_DIR, f);
            try {
                const st = fs.statSync(full);
                if (now - st.mtimeMs > ttlMs) fs.unlinkSync(full);
            } catch { }
        }
    } catch { }
}


function resolvePathWithState(obj, pathStr) {
    if (obj == null) return { value: "", state: "NOT_FOUND" };
    if (!pathStr || typeof pathStr !== "string") return { value: "", state: "NOT_FOUND" };

    const normalized = pathStr
        .trim()
        .replace(/\[(\d+)\]/g, ".$1")
        .replace(/\[["']([^"']+)["']\]/g, ".$1")
        .replace(/\.+/g, ".")
        .replace(/^\./, "")
        .replace(/\.$/, "");

    if (!normalized) return { value: "", state: "NOT_FOUND" };

    const parts = normalized.split(".").filter(Boolean);

    let cur = obj;
    for (const p of parts) {
        if (cur == null) return { value: "", state: "NOT_FOUND" };

        // ✅ จุดสำคัญ: ต้องเช็ค "มี key จริงไหม"
        if (typeof cur === "object" && cur !== null && !(p in cur)) {
            return { value: "", state: "NOT_FOUND" };
        }
        cur = cur[p];
    }

    if (cur === undefined) return { value: "", state: "NOT_FOUND" };
    if (cur === null) return { value: "", state: "FOUND_BUT_EMPTY" };
    if (typeof cur === "string" && cur.trim() === "") return { value: "", state: "FOUND_BUT_EMPTY" };

    return { value: cur, state: "FOUND" };
}


function safeGet(obj, pathStr, fallback = "") {
    const r = resolvePathWithState(obj, pathStr);
    return (r.state === "FOUND") ? r.value : fallback;
}


// template เดิม (ยังอยู่ไปก่อน)
const TEMPLATE_XLSX = path.join(__dirname, '..', 'templates', 'template.xlsx');

// ใช้เก็บไฟล์ template ที่ upload
function getTemplatePathFromId(templateId) {
    ensureTmpDir();
    return path.join(TMP_DIR, `${templateId}.xlsx`);
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
        try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch { }
        return { buffer: xlsxBuffer, convertedFromLegacyXls: true };
    }

    // ถ้าไม่ใช่ xls/xlsx → ไม่รองรับ
    throw new Error('Unsupported Excel format');
}

// -------------------------------------------------------
// Utility
// -------------------------------------------------------

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

    // ✅ ตั้งค่าเฉพาะที่ template ไม่ได้ตั้ง
    cell.font = {
        ...oldFont,
        name: oldFont.name ?? 'TH SarabunPSK',
        // อย่าไปแตะ color ถ้าเขาตั้งมาแล้ว
        // ถ้าอยากกัน "ไม่มีสี" ค่อยใส่เฉพาะตอน undefined
        ...(oldFont.color === undefined ? { color: { argb: 'FF000000' } } : {}),
    };

    cell.alignment = {
        ...oldAlign,
        ...(oldAlign.vertical === undefined ? { vertical: 'top' } : {}),
        // wrapText อย่าไปตั้ง default มั่ว ๆ เพราะ template อาจคุมเอง
        // ...(oldAlign.wrapText === undefined ? { wrapText: false } : {}),
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


        if (w > colPx && current !== prefix) {
            lines.push(current);

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
    let hasExplicitStyle = false;            // ✅ เพิ่ม
    let mainKeyPath = null;
    let lastExplicitStyleTokens = null;
    const useStyleCache = !!cell._fromArrayTemplate;
    // ✅ มีสิทธิ์ใช้ cache เฉพาะ cell ที่มาจาก expand array เท่านั้น
    cell.value = cell.value.replace(/{{\s*([^{}]+?)\s*}}/g, (_, inner) => {
        const tokens = splitPlaceholder(inner);
        if (!tokens.length) return '';

        const key = tokens[0];
        const styleTokens = tokens.slice(1);

        if (key.toLowerCase() === 'style') {
            hasExplicitStyle = true;
            lastExplicitStyleTokens = styleTokens.slice(); // ✅ เอาตัวท้ายสุด
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

        const r = resolvePathWithState(data, keyPath);
        const v = String(r.value ?? "");
        return v;
    });

    if (hasArrayToken) {
        const oldAlign = cell.alignment || {};
        cell.alignment = {
            ...oldAlign,
            ...(oldAlign.wrapText === undefined ? { wrapText: true } : {}),
            vertical: oldAlign.vertical || 'top',
        };

        // ✅ only set border if template didn't set any

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

    if (mainKeyPath && !hasExplicitStyle && useStyleCache) {
        const norm = normalizeKeyForStyle(mainKeyPath);
        const defTokens = defaultStyleByKey[norm];
        if (defTokens?.length) applyInlineStyle(cell, defTokens);
    }
    // ✅ apply style ทีเดียวหลังจบ (ตัวท้ายสุดชนะ)
    if (hasExplicitStyle && lastExplicitStyleTokens?.length) {
        applyInlineStyle(cell, lastExplicitStyleTokens);

        // ✅ จำ style เฉพาะ array template เท่านั้น
        if (mainKeyPath && useStyleCache) {
            defaultStyleByKey[normalizeKeyForStyle(mainKeyPath)] = lastExplicitStyleTokens;
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

    let s = String(inner).replace(/\r\n/g, '\n').replace(/\n/g, ' ');

    // split ตาม | ก่อน
    let parts = s.split('|').map(p => p.trim()).filter(Boolean);
    if (parts.length === 0) return [];

    // ✅ แตก space เพิ่มเฉพาะ style เท่านั้น
    const head = parts[0];
    if (!isQuoted(head) && head.toLowerCase().startsWith('style') && /\s/.test(head)) {
        const firstPieces = head.split(/\s+/).filter(Boolean);
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
    return safeGet(data, trimmed, "");
}

function setByPath(obj, pathStr, value) {
    if (!pathStr) return;

    const normalized = String(pathStr)
        .trim()
        .replace(/\[(\d+)\]/g, '.$1')
        .replace(/\[["']([^"']+)["']\]/g, '.$1')  // ✅ เพิ่มบรรทัดนี้
        .replace(/\.+/g, ".")
        .replace(/^\./, "")
        .replace(/\.$/, "");

    const parts = normalized.split('.').filter(Boolean);
    if (!parts.length) return;

    let cur = obj;
    for (let i = 0; i < parts.length; i++) {
        const key = parts[i];
        const isLast = i === parts.length - 1;

        const nextKey = parts[i + 1];
        const nextIsIndex = nextKey != null && /^\d+$/.test(nextKey);

        if (isLast) {
            if (Array.isArray(cur) && /^\d+$/.test(key)) cur[Number(key)] = value;
            else cur[key] = value;
            return;
        }

        if (Array.isArray(cur) && /^\d+$/.test(key)) {
            const idx = Number(key);
            if (cur[idx] == null) cur[idx] = nextIsIndex ? [] : {};
            cur = cur[idx];
        } else {
            if (cur[key] == null) cur[key] = nextIsIndex ? [] : {};
            cur = cur[key];
        }
    }
}

function addKey(schema, keyPath) {
    const kp = String(keyPath || "")
        .replace(/\[i\]/g, "[0]")
        .replace(/\[\]/g, "[0]")
        .trim();
    if (!kp) return;

    setByPath(schema, kp, "");
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
function escapeRegExp(s) {
    return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function findArrayNamesInRow(row) {
    const found = new Set();

    row.eachCell({ includeEmpty: true }, (cell) => {
        if (typeof cell.value !== 'string') return;

        const re = /{{\s*([^{}]+?)\s*}}/g;
        let m;
        while ((m = re.exec(cell.value))) {
            const inner = m[1];
            const tokens = splitPlaceholder(inner);

            // เช็คทุก token (รวม fx args ด้วย เพราะมันถูก split มาแล้วเหมือนกัน)
            for (const tk of tokens) {
                const s = String(tk || '').trim();
                if (!s) continue;

                const mm = s.match(/(.+?)\[i\]/); // จับก่อน [i] ทั้งหมด (รวม dot path)
                if (mm) {
                    const name = (mm[1] || '').trim();
                    if (name) found.add(name);
                }
            }
        }
    });

    return Array.from(found);
}

function expandArrayRows(ws, data) {
    for (let rowNum = ws.rowCount; rowNum >= 1; rowNum--) {
        const row = ws.getRow(rowNum);

        const arrayNames = findArrayNamesInRow(row);
        if (!arrayNames.length) continue;

        // ✅ หา maxLen ตามกฎ: ไม่มี/ไม่ใช่ array = 0, สุดท้าย max ต้อง >= 1
        const lens = arrayNames.map((p) => {
            const r = resolvePathWithState(data, p);
            return Array.isArray(r.value) ? r.value.length : 0;
        });
        const maxLen = Math.max(1, ...lens);

        // templateRow
        const templateRow = ws.getRow(rowNum);

        // mark (เหมือนเดิม)
        templateRow._fromArrayTemplate = true;
        templateRow.eachCell({ includeEmpty: true }, (c) => { c._fromArrayTemplate = true; });

        // ✅ normalize: [i] -> [0] ทั้งแถว (ครอบคลุมทุก array)
        templateRow.eachCell({ includeEmpty: true }, (c) => {
            if (typeof c.value === 'string') c.value = c.value.replace(/\[i\]/g, '[0]');
        });

        // ถ้า maxLen = 1 ไม่ต้อง insert row เพิ่ม แค่ normalize ก็พอ
        if (maxLen <= 1) continue;

        const templateValues = templateRow.values.slice();
        const templateHeight = templateRow.height;

        const templateStyles = {};
        templateRow.eachCell({ includeEmpty: true }, (tmplCell, col) => {
            templateStyles[col] = JSON.parse(JSON.stringify(tmplCell.style || {}));
        });

        // เตรียม regex ของทุก arrayName: arr[0] -> arr[i]
        const reByArr0 = arrayNames.map((arrName) => ({
            arrName,
            re0: new RegExp(`${escapeRegExp(arrName)}\\[0\\]`, 'g'),
        }));

        // insert rows 1..maxLen-1
        for (let i = 1; i < maxLen; i++) {
            const newRow = ws.insertRow(rowNum + i, []);
            newRow.values = templateValues.slice();
            if (templateHeight != null) newRow.height = templateHeight;
            newRow._fromArrayTemplate = true;

            templateRow.eachCell({ includeEmpty: true }, (tmplCell, col) => {
                const cell = newRow.getCell(col);
                cell.style = JSON.parse(JSON.stringify(templateStyles[col] || {}));
                cell._fromArrayTemplate = true;

                if (typeof cell.value === 'string') {
                    // ✅ replace ทุก array ในลิสต์
                    for (const it of reByArr0) {
                        cell.value = cell.value.replace(it.re0, `${it.arrName}[${i}]`);
                    }
                }
            });
        }
    }
}


function assertNoILeft(ws) {
    ws.eachRow(r => r.eachCell({ includeEmpty: true }, c => {
        if (typeof c.value === 'string' && /\[i\]/.test(c.value)) {
            throw new Error(`Found [i] left at ${ws.name}!${c.address}: ${c.value}`);
        }
    }));
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



function inchToMm(v) {
    const n = Number(v);
    return Number.isFinite(n) ? (n * 25.4) : 0;
}

function mmToInch(v) {
    const n = Number(v) || 0;
    return n / 25.4;
}

// reverse map: exceljs paperSize number -> name
const PAPER_NUM_TO_NAME = {
    1: 'Letter',
    8: 'A3',
    9: 'A4',
    11: 'A5',
};

function getTemplateOptionsFromWorksheet(ws) {
    const psNum = ws?.pageSetup?.paperSize;
    const ori = String(ws?.pageSetup?.orientation || '').toLowerCase();

    const paperSize = PAPER_NUM_TO_NAME[psNum] || 'A4';
    const orientation = (ori === 'landscape') ? 'l' : 'p';

    // exceljs margins are in inches (if present)
    const m = ws?.pageSetup?.margins || {};
    const margin = {
        left: inchToMm(m.left ?? 0),
        right: inchToMm(m.right ?? 0),
        top: inchToMm(m.top ?? 0),
        bottom: inchToMm(m.bottom ?? 0),
    };

    return { paperSize, orientation, margin };
}

function mergeOptions(sysDefault, tplOpt, userOpt) {
    // sysDefault <- tplOpt <- userOpt
    const out = { ...sysDefault, ...tplOpt, ...userOpt };

    // margin ต้อง merge แบบ object
    if (sysDefault?.margin || tplOpt?.margin || userOpt?.margin) {
        out.margin = {
            ...(sysDefault.margin || {}),
            ...(tplOpt?.margin || {}),
            ...(userOpt?.margin || {}),
        };
    }
    return out;
}
function getPaperWidthInches(paperSize, orientation) {
    const size = String(paperSize || 'A4').toUpperCase();
    const ori = String(orientation || 'portrait').toLowerCase();

    // หน่วย: inch (กว้าง)
    const map = {
        A3: 11.69,   // 297mm
        A4: 8.27,    // 210mm
        A5: 5.83,    // 148mm
        LETTER: 8.5, // 8.5in
    };

    let w = map[size] ?? map.A4;
    if (ori === 'l' || ori === 'landscape') {
        // สลับกว้าง/สูง → กว้างของ landscape = “ด้านยาว”
        const longMap = { A3: 16.54, A4: 11.69, A5: 8.27, LETTER: 11 };
        w = longMap[size] ?? 11.69;
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
        ws.pageSetup.fitToHeight = 0;
        ws.pageSetup.scale = undefined;
        return;
    }


    if (!opt.autoScaleToFitWidth) {
        return; // ไม่ทำอะไรเลยถ้าไม่ได้เปิด
    }

    const scaleFloat = (printableWidthInches / totalWidthInches) * 100;
    const scale = Math.floor(Math.min(100, scaleFloat));

    if (scale < 100 && scale > 10) {
        ws.pageSetup.fitToPage = false;
        ws.pageSetup.fitToWidth = undefined;
        ws.pageSetup.fitToHeight = undefined;
        ws.pageSetup.scale = scale;
    }
}
// -------------------------------------------------------
// render excel
// -------------------------------------------------------
async function fillXlsx(tplPath, data) {
    const wb = new Excel.Workbook();
    await wb.xlsx.readFile(tplPath);

    const defaultStyleByKey = {};

    const sysDefaultOpt = {
        paperSize: 'A4',
        orientation: 'p',
        margin: { left: 0, right: 0, top: 0, bottom: 0 },
        pageNumber: true,
        pageNumberPosition: 'bottom-center',
        repeatHeaderRows: '',
        autoScaleToFitWidth: false,
    };

    const userOpt = (data.__options || {});

    const paperMap = { A3: 8, A4: 9, A5: 11, Letter: 1 };

    wb.eachSheet(ws => {
        // ✅ 1) อ่านค่าจาก template ก่อน (ต่อ sheet)
        const tplOpt = getTemplateOptionsFromWorksheet(ws);

        // ✅ 2) รวม option ตามลำดับ: system <- template <- user
        const opt = mergeOptions(sysDefaultOpt, tplOpt, userOpt);

        // ✅ 3) เซ็ตตาม opt ที่ได้มา (ถ้า user ไม่ส่งมา ก็จะเป็นของ template เอง)
        ws.pageSetup.paperSize = paperMap[opt.paperSize] || 9;
        ws.pageSetup.orientation = mapOrientation(opt.orientation);

        // -------- margin เดิมของตัว (opt.margin เป็น mm) --------
        const margin = opt.margin ?? 0;

        let marginLeft, marginRight, marginTop, marginBottom;
        if (margin && typeof margin === 'object') {
            marginLeft = mmToInch(margin.left);
            marginRight = mmToInch(margin.right);
            marginTop = mmToInch(margin.top);
            marginBottom = mmToInch(margin.bottom);
        } else {
            const mInch = mmToInch(margin);
            marginLeft = marginRight = marginTop = marginBottom = mInch;
        }

        if (opt.pageNumber) {
            const pos = normalizePos(opt.pageNumberPosition || 'bottom-center');
            const MIN_TOP = 0.4;
            const MIN_BOTTOM = 0.4;
            if (pos.startsWith('top') && marginTop < MIN_TOP) marginTop = MIN_TOP;
            if (pos.startsWith('bottom') && marginBottom < MIN_BOTTOM) marginBottom = MIN_BOTTOM;
        }

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

        // ---- ที่เหลือใช้ของเดิมได้เลย ----
        expandArrayRows(ws, data);
        assertNoILeft(ws);

        ws.eachRow(row => row.eachCell(cell => {
            replaceTokensInCell(cell, data, defaultStyleByKey);
        }));

        ws.eachRow(row => row.eachCell((cell, colNumber) => {
            smartWrapLabelValueCell(ws, cell, colNumber);
        }));

        if (IS_LINUX) autoAdjustRowHeightByWrap(ws);

        ws.eachRow(row => {
            row.eachCell(cell => {
                const oldFont = cell.font || {};
                if (!oldFont.name) cell.font = { ...oldFont, name: 'TH SarabunPSK' };
            });
        });
    });
    ensureTmpDir();
    const outXlsx = path.join(TMP_DIR, `filled_${Date.now()}_${Math.random().toString(36).slice(2, 6)}.xlsx`);

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
            '--headless', '--nologo', '--nolockcheck', '--nodefault', '--norestore',
            '--convert-to', 'pdf', '--outdir', outDir, xlsxPath,
        ];

        const p = spawn(SOFFICE, args, { stdio: 'ignore' });

        p.on('error', reject);

        p.on('exit', code => {
            if (code !== 0) return reject(new Error('soffice exit ' + code));

            const base = path.basename(xlsxPath).replace(/\.[^.]+$/, '');
            const pdfPath = path.join(outDir, `${base}.pdf`);

            if (!fs.existsSync(pdfPath)) {
                return reject(new Error('PDF not found after convert'));
            }
            resolve(pdfPath);
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
                        for (const t of rest) {
                            const s = String(t || '').trim();
                            if (!s) continue;

                            if (s.toLowerCase().startsWith('as:')) {
                                addKey(schema, s.slice(3).trim());
                                continue;
                            }

                            // ถ้าเป็น literal หรือ op ก็ข้าม
                            if (isQuoted(s)) continue;
                            if (/^(==|!=|>=|<=|>|<)$/i.test(s)) continue;
                            if (/^[+-]?(\d+(\.\d+)?|\.\d+)$/.test(s)) continue;

                            // เก็บเป็น input path
                            addKey(schema, s);
                        }
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
        /*   const referer = req.get('Referer') || '';
          let permission = await checkPermissionUrl(referer);
          if (!permission) {
              throw Error('No Permission');
          } */
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
                autoScaleToFitWidth: true,
                forceSinglePage: true
            },
            __templateId: null,
            __needUploadTemplate: true,
        };


        res.json(finalSchema);
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});
function touchFile(p) {
    const now = new Date();
    try { fs.utimesSync(p, now, now); } catch { }
}

router.post('/render', async (req, res) => {
    try {
        const referer = req.get('Referer') || '';
        let permission = await checkPermissionUrl(referer);
        if (!permission) {
            throw Error('No Permission');
        }
        cleanupOldFiles();
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
        touchFile(tplPath);
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
        let cleaned = false;
        const cleanup = () => {
            if (cleaned) return;
            cleaned = true;
            fs.unlink(xlsx, () => { });
            fs.unlink(pdf, () => { });
            if (finalPdf !== pdf) fs.unlink(finalPdf, () => { });
        };

        res.on('finish', cleanup); // ส่งครบ
        res.on('close', cleanup);  // client ตัดกลางทาง
        const stream = fs.createReadStream(finalPdf);
        stream.on('error', (err) => {
            cleanup();
            if (!res.headersSent) res.status(500).json({ error: err.message });
            else res.destroy();
        });
        stream.pipe(res);

    } catch (e) {
        console.error(e);
        res.status(500).json({ error: e.message });
    }
});



// POST /api/gendoc/schema/upload
// รับไฟล์ Excel แล้วคืน schema + templateId
router.post('/schema/upload', async (req, res) => {
    try {
        cleanupOldFiles();
        /*    const referer = req.get('Referer') || '';
           let permission = await checkPermissionUrl(referer);
           if (!permission) {
               throw Error('No Permission');
           } */

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

        const wb = new Excel.Workbook();
        await wb.xlsx.readFile(tplPath);
        const firstWs = wb.worksheets?.[0];
        const tplOpt = firstWs ? getTemplateOptionsFromWorksheet(firstWs) : {};

        const finalSchema = {
            __templateId: templateId,
            __convertedFromLegacyXls: convertedFromLegacyXls,
            __options: mergeOptions(
                {
                    paperSize: 'A4',
                    orientation: 'p',
                    margin: { left: 0, right: 0, top: 0, bottom: 0 },
                    pageNumber: true,
                    pageNumberPosition: 'bottom-center',
                    repeatHeaderRows: '',
                    autoScaleToFitWidth: true,
                    forceSinglePage: true,
                },
                tplOpt,
                {}
            ),
            ...schema,
        };

        return res.json(finalSchema);
    } catch (e) {
        console.error(e);
        res.status(500).json({ error: e.message });
    }
});


function toCsvCell(v) {
    const s = (v === null || v === undefined) ? '' : String(v);
    // escape double quotes
    const escaped = s.replace(/"/g, '""');
    // wrap if has comma, quote, newline
    if (/[",\n\r]/.test(escaped)) return `"${escaped}"`;
    return escaped;
}

function buildCsvFromIndexedObject(obj) {
    // obj = { "0": {...}, "1": {...}, start_date:"", ... }
    const rows = [];
    const keysNumeric = Object.keys(obj || {}).filter(k => /^\d+$/.test(k)).sort((a, b) => Number(a) - Number(b));
    if (keysNumeric.length === 0) return { csv: '\ufeff', rowCount: 0, columns: [] };

    // union columns ตามลำดับที่เจอ (stable)
    const colSet = new Set();
    for (const k of keysNumeric) {
        const r = obj[k];
        if (r && typeof r === 'object' && !Array.isArray(r)) {
            for (const ck of Object.keys(r)) colSet.add(ck);
        }
    }
    const columns = Array.from(colSet);

    // header
    rows.push(columns.map(toCsvCell).join(','));

    // body
    for (const k of keysNumeric) {
        const r = obj[k] || {};
        rows.push(columns.map(c => toCsvCell(r[c])).join(','));
    }

    // BOM + join
    const csv = '\ufeff' + rows.join('\r\n');
    return { csv, rowCount: keysNumeric.length, columns };
}
function convertToOdt(inputPath) {
    const outDir = path.dirname(inputPath);

    return new Promise((resolve, reject) => {
        const args = [
            '--headless', '--nologo', '--nolockcheck', '--nodefault', '--norestore',
            '--convert-to', 'odt', '--outdir', outDir, inputPath,
        ];

        const p = spawn(SOFFICE, args, { stdio: 'ignore' });
        p.on('error', reject);
        p.on('exit', code => {
            if (code !== 0) return reject(new Error('soffice exit ' + code));

            const base = path.basename(inputPath).replace(/\.[^.]+$/, '');
            const odtPath = path.join(outDir, `${base}.odt`);

            if (!fs.existsSync(odtPath)) return reject(new Error('ODT not found after convert'));
            resolve(odtPath);
        });
    });
}
// POST /api/gendoc/export?format=pdf|xlsx
router.post('/export', async (req, res) => {
    try {
        const referer = req.get('Referer') || '';
        const permission = await checkPermissionUrl(referer);
        if (!permission) throw Error('No Permission');

        cleanupOldFiles();

        const raw = req.body || {};
        const data = Array.isArray(raw) ? { items: raw } : raw;

        const format = String(req.query.format || 'pdf').toLowerCase();
        if (!['pdf', 'xlsx', 'csv', 'odt'].includes(format)) {
            return res.status(400).json({ error: 'Invalid format (use pdf, xlsx, csv, odt)' });
        }

        // ✅ CSV: ไม่ต้องใช้ template / ไม่ต้องสร้างไฟล์อื่น
        if (format === 'csv') {
            const source = (data && typeof data.data === 'object') ? data.data : data;
            const { csv } = buildCsvFromIndexedObject(source);

            res.setHeader('Content-Type', 'text/csv; charset=utf-8');
            res.setHeader('Content-Disposition', `attachment; filename="export_${Date.now()}.csv"`);
            return res.send(csv);
        }

        // options
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
        touchFile(tplPath);

        // 1) fill xlsx
        const xlsxPath = await fillXlsx(tplPath, data);

        // ✅ XLSX
        if (format === 'xlsx') {
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', `attachment; filename="export_${Date.now()}.xlsx"`);

            let cleaned = false;
            const cleanup = () => {
                if (cleaned) return;
                cleaned = true;
                fs.unlink(xlsxPath, () => { });
            };

            res.on('finish', cleanup);
            res.on('close', cleanup);

            const stream = fs.createReadStream(xlsxPath);
            stream.on('error', (err) => {
                cleanup();
                if (!res.headersSent) res.status(500).json({ error: err.message });
                else res.destroy();
            });
            return stream.pipe(res);
        }

        // 2) ทำ pdf (ใช้ร่วมกันทั้ง pdf และ odt)
        const pdfPath = await convertToPdf(xlsxPath);
        let finalPdf = pdfPath;

        if (opt.pageNumber !== false) {
            finalPdf = await addPageNumbers(pdfPath, {
                pageNumberPosition: opt.pageNumberPosition,
            });
        }

        // ✅ ODT (ทำหลังได้ finalPdf แล้ว)
        if (format === 'odt') {
            const odtPath = await convertToOdt(finalPdf);

            res.setHeader('Content-Type', 'application/vnd.oasis.opendocument.text');
            res.setHeader('Content-Disposition', `attachment; filename="export_${Date.now()}.odt"`);

            let cleaned = false;
            const cleanup = () => {
                if (cleaned) return;
                cleaned = true;
                fs.unlink(xlsxPath, () => { });
                fs.unlink(pdfPath, () => { });
                if (finalPdf !== pdfPath) fs.unlink(finalPdf, () => { });
                fs.unlink(odtPath, () => { });
            };

            res.on('finish', cleanup);
            res.on('close', cleanup);

            const stream = fs.createReadStream(odtPath);
            stream.on('error', (err) => {
                cleanup();
                if (!res.headersSent) res.status(500).json({ error: err.message });
                else res.destroy();
            });
            return stream.pipe(res);
        }

        // ✅ PDF
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename="export_${Date.now()}.pdf"`);

        let cleaned = false;
        const cleanup = () => {
            if (cleaned) return;
            cleaned = true;
            fs.unlink(xlsxPath, () => { });
            fs.unlink(pdfPath, () => { });
            if (finalPdf !== pdfPath) fs.unlink(finalPdf, () => { });
        };

        res.on('finish', cleanup);
        res.on('close', cleanup);

        const stream = fs.createReadStream(finalPdf);
        stream.on('error', (err) => {
            cleanup();
            if (!res.headersSent) res.status(500).json({ error: err.message });
            else res.destroy();
        });
        return stream.pipe(res);

    } catch (e) {
        console.error(e);
        res.status(500).json({ error: e.message });
    }
});



// -------------------------------------------------------

module.exports = router;
