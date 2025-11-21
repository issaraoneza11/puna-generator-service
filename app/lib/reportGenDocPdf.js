// pdf.tools.js  (CommonJS, สำหรับหลังบ้าน)
'use strict';
const { jsPDF } = require('jspdf');

// helper ตั้งชื่อไฟล์ (ถ้าไม่ส่งชื่อมา จะได้ report.pdf)
function beFileName(prefix = 'report') {
    return `${prefix}.pdf`;
}

function createPdfTools(opts = {}) {
    const orientation = opts.orientation || 'p';
    const unit = opts.unit || 'mm';
    const format = opts.format || 'a4';

    const padding = opts.padding ?? 20;
    const themeRGB = opts.themeRGB || [194, 194, 194];
    const [r, g, b] = themeRGB;

    const doc = new jsPDF(orientation, unit, format);
    const pageSize = doc.internal.pageSize;
    const wpage = pageSize.getWidth();
    const hpage = pageSize.getHeight();

    let heightpage = 29;        // y-cursor เริ่มต้น
    const endpage = hpage - 35; // ขอบล่างก่อนตัดหน้า


    const fontItemsRaw = Array.isArray(opts.fonts)
        ? opts.fonts
        : (Array.isArray(opts.font) ? opts.font : (opts.font ? [opts.font] : []));
    const fontRegistry = new Map();
    for (const f of fontItemsRaw) {
        if (!f || !f.name || !f.data) continue;
        const style = f.style || 'normal';
        const vfsName = f.vfsName || `${f.name}.${(f.ext || 'ttf')}`;
        doc.addFileToVFS(vfsName, f.data);
        doc.addFont(vfsName, f.name, style);
        fontRegistry.set(`${f.name}::${style}`, true);
    }
    if (fontItemsRaw.length) {
        const first = fontItemsRaw[0];
        doc.setFont(first.name, first.style || 'normal');
    }

    // ฟอนต์ (ไทย) — ต้องส่ง base64 ttf ผ่าน opts.font.data
    /*  if (opts.font && opts.font.name && opts.font.data) {
         const style = opts.font.style || 'normal';
         doc.addFileToVFS(`${opts.font.name}.ttf`, opts.font.data);
         doc.addFont(`${opts.font.name}.ttf`, opts.font.name, style);
         doc.setFont(opts.font.name, style);
     }
  */
    // ===== helpers =====
    function CutPadding(mode, percent, num) {
        if (mode === 'h') {
            const base = (hpage * percent) / 100;
            return num !== undefined ? base + num : base;
        }
        // mode === 'w'
        const base = ((wpage - (padding * 2)) * percent) / 100 + padding;
        return num !== undefined ? base + num : base;
    }

    function TexttoString(text) {
        return text == null ? '' : String(text);
    }

    function NumtoString(num, fix) {
        const n = Number(num ?? 0);
        const s = fix != null ? n.toFixed(fix) : String(n);
        return s.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    }

    function newpage(threshold) {
        const t = threshold || endpage;
        if (heightpage >= t) {
            doc.addPage();
            heightpage = 29;
            if (opts.onNewPage) opts.onNewPage(doc); // header/bg ต่อหน้า
        }
    }

    function addPage() {
        doc.addPage();
    }
    function setPage(page) {
        doc.setPage(page);
    }

    function Colposition(mode, name, col, pos, pad) {
        if (mode === 'r') {
            const poscol = (typeof pos === 'number' ? pos : 0) || 0;
            const colcount = col - 1;
            let colsum = 0;
            for (let i = colcount - 1; i > 0; i--) colsum += name[`col${i}`] || 0;
            if (colcount === 0) return CutPadding('w', 0);
            return CutPadding('w', 0, (name[`col${colcount}`] || 0) + poscol) + colsum;
        }
        // mode === 't'
        let colsum = 0;
        const paddingcol = pad || 0;
        for (let i = col - 1; i > 0; i--) colsum += name[`col${i}`] || 0;
        if (col === 0) return CutPadding('w', 0);

        if (pos === 'c') return CutPadding('w', 0, (name[`col${col}`] * (50 + paddingcol)) / 100) + colsum;
        if (pos === 'l') return CutPadding('w', 0, (name[`col${col}`] * (1 + paddingcol)) / 100) + colsum;
        if (pos === 'r') return CutPadding('w', 0, (name[`col${col}`] * (99 - paddingcol)) / 100) + colsum;

        return CutPadding('w', 0, name[`col${col}`]) + colsum;
    }

    // ===== สวิตช์คำสั่งหลัก =====
    async function Shot(mode, c1, c2, c3, c4, c5, c6, c7) {
        if (mode === 'add') return c1 !== undefined ? doc.addPage(c1) : doc.addPage();

        if (mode === 'new') {
            // บน Node ให้คืน bytes
            if (typeof window !== 'undefined' && window?.open) {
                return window.open(doc.output('bloburl'));
            }
            return doc.output('arraybuffer');
        }

        if (mode === 'save') {
            // บน Browser จะดาวน์โหลด, บน Node คืน Buffer
            const name = c1 || beFileName('report'); // ✅ แก้ default ให้ถูก
            if (typeof window !== 'undefined' && doc.save) {
                return doc.save(name);
            }
            return Buffer.from(doc.output('arraybuffer'));
        }

        if (mode === 'newsave') {
            const name = c1 || beFileName('report'); // ✅ แก้ default ให้ถูก
            if (typeof window !== 'undefined' && window?.open && doc.save) {
                window.open(doc.output('bloburl'));
                return doc.save(name);
            }
            return Buffer.from(doc.output('arraybuffer'));
        }

        if (mode === 'fs') return doc.setFontSize(c1);
        if (mode === 'dc') return doc.setDrawColor(c1, c2, c3);
        if (mode === 'tc') return doc.setTextColor(c1, c2, c3);
        if (mode === 'fc') return doc.setFillColor(c1, c2, c3);
        if (mode === 'lw') return doc.setLineWidth(c1);
        if (mode === 'sf') {
            let name, style, size;
            if (c1 && typeof c1 === 'object') ({ name, style, size } = c1);
            else { name = c1; style = c2; size = c3; }
            if (name) {
                const key = `${name}::${style || 'normal'}`;
                if (fontRegistry.has(key)) { try { doc.setFont(name, style || 'normal'); } catch (_) { } }
            }
            if (typeof size === 'number') doc.setFontSize(size);
            return;
        }
        if (mode === 'i' && c1) {
            // c1=dataURL/Uint8Array; c2=x c3=y c4=w c5=h; c6='fit'; c7={width,height}
            if (c5 !== undefined && c6 === 'fit') {
                const Width = (c7 && c7.width) || 1920;
                const Height = (c7 && c7.height) || 1080;
                const imgar = Width / Height;
                const reactar = c4 / c5;
                const scale = imgar > reactar ? (c4 / Width) : (c5 / Height);
                const nw = Width * scale, nh = Height * scale;
                const x = (c4 - nw) / 2, y = (c5 - nh) / 2;
                return doc.addImage(c1, 'JPEG', c2 + x, c3 + y, nw, nh, undefined, 'FAST');
            }
            return doc.addImage(c1, 'JPEG', c2, c3, c4, c5 || c4);
        }

        if (mode === 'r') {
            if (c5 === 'd') return doc.rect(c1, c2, c3, c4, 'D');
            if (c5 === 'fd') return doc.rect(c1, c2, c3, c4, 'FD');
            if (c5 === 'f') return doc.rect(c1, c2, c3, c4, 'F');
            return doc.rect(c1, c2, c3, c4);
        }

        if (mode === 't') {

            c1 = (c1 || '').toString().replaceAll('/n','');
            if (c4 === 'c') return doc.text(c1, c2, c3, { align: 'center' });
            if (c4 === 'r') return doc.text(c1, c2, c3, { align: 'right' });
            if (c4 === 'l') return doc.text(c1, c2, c3, { align: 'left' });
            return doc.text(c1, c2, c3);
        }
        if (mode === 'rt') {
            if (c4 === 'c') return doc.text(c1, c2, c3, { align: 'center',angle:c5 });
            if (c4 === 'r') return doc.text(c1, c2, c3, { align: 'right',angle:c5 });
            if (c4 === 'l') return doc.text(c1, c2, c3, { align: 'left',angle:c5 });
            return doc.text(c1, c2, c3);
        }
        if (mode === 'tb') {
            if (c5) {
                let align = 'left'
                if (c4 === 'c') align = 'center'
                if (c4 === 'r') align = 'right'
                for (let l = 0; l < c5; l++) {
                    doc.text(c1, c2, c3, { align: align });

                }
            } else {
                if (c4 === 'c') { doc.text(c1, c2, c3, { align: 'center' }); doc.text(c1, c2, c3, { align: 'center' }); doc.text(c1, c2, c3, { align: 'center' }); }
                else if (c4 === 'r') { doc.text(c1, c2, c3, { align: 'right' }); doc.text(c1, c2, c3, { align: 'right' }); doc.text(c1, c2, c3, { align: 'right' }); }
                else if (c4 === 'l') { doc.text(c1, c2, c3, { align: 'left' }); doc.text(c1, c2, c3, { align: 'left' }); doc.text(c1, c2, c3, { align: 'left' }); }
                else { doc.text(c1, c2, c3); doc.text(c1, c2, c3); doc.text(c1, c2, c3); }
            }


        }
    }

    // ===== ตาราง (ตามโครงเดิมของตัว) =====
    function ShotTable(mode, head, pos, col, coll, loop, color, hig, header) {
        let collengthsum = (coll == null ? 5 : coll);
        const loopsum = (loop == null ? 10 : loop);

        if (mode === 'ht') {
            for (let t = 1; t <= col; t++) {
                Shot('fc', r, g, b);
                Shot('r', Colposition('r', head, t), pos, CutPadding('w', 0, head[`col${t}`] - padding), collengthsum, color);
                if (head[`name${t}`] !== undefined) {
                    Shot('t', Colposition('t', head, t, 'c'), pos + (collengthsum / 1.6), TexttoString(head[`name${t}`]), 'c');
                }
            }
        }



        if (mode === 'lt') {
            const lengthloop = (loopsum * collengthsum) + collengthsum;
            for (let t = 1; t <= col; t++) {
                Shot('r', Colposition('r', head, t), pos, CutPadding('w', 0, head[`col${t}`] - padding), lengthloop);
            }
        }

        if (mode === 'st') {
            let lengthloop = collengthsum;
            for (let l = 0; l < loopsum; l++) {
                for (let t = 1; t <= col; t++) {
                    Shot('r', Colposition('r', head, t), pos + lengthloop, CutPadding('w', 0, head[`col${t}`] - padding), collengthsum);
                }
                lengthloop += collengthsum;
            }
        }

        if (mode === 'htc') {
            for (let t = 1; t <= col; t++) {
                Shot('fc', r, g, b);
                Shot('r', Colposition('r', head, t), pos, CutPadding('w', 0, head[`col${t}`] - padding), collengthsum * (hig || 1), color);
                if (head[`name${t}`]) {
                    for (const c of head[`name${t}`]) {
                        Shot('t', Colposition('t', head, t, 'c'), pos + (collengthsum * 0.6), TexttoString(c), 'c');
                        collengthsum += (coll || 0);
                    }
                    collengthsum = coll || 0;
                }
            }
            heightpage += collengthsum * (hig || 1);
        }

        if (mode === 'ltc') {
            let befor = 0, higbefor = pos, maxhig = 0;
            for (let t = 1; t <= col; t++) {
                Shot('fc', r, g, b);
                if (head[`name${t}`]) {
                    for (let d = befor; d < head[`name${t}`].length; d++) {
                        const c = head[`name${t}`][d];
                        if (heightpage + ((d - befor) * 6) > (hpage - 47)) {
                            if (t < col) {
                                for (let t2 = t + 1; t2 <= col; t2++) {
                                    let fakecolsum = coll || 0;
                                    for (let dx = befor; dx <= d; dx++) {
                                        if ((head[`name${t2}`] || []).length > dx) {
                                            const cx = head[`name${t2}`][dx];
                                            Shot('t', Colposition('t', head, t2, 'l', 3), higbefor + (fakecolsum), TexttoString(cx), 'l');
                                            fakecolsum += (coll || 0);
                                        } else break;
                                    }
                                }
                            }
                            for (let al = 1; al <= col; al++) {
                                Shot('r', Colposition('r', head, al), higbefor, CutPadding('w', 0, head[`col${al}`] - padding), collengthsum + 2);
                            }
                            heightpage += ((d - befor) * 6);
                            newpage(hpage - 47);
                            higbefor = heightpage;
                            collengthsum = col;
                            befor = d;
                            maxhig = col;
                            if (header) {
                                ShotTable('htc', header, heightpage, header.loop, 7, '', 'fd', header.height);
                                higbefor += 7;
                            }
                        }
                        Shot('t', Colposition('t', head, t, 'l', 3), higbefor + (collengthsum), TexttoString(c), 'l');
                        collengthsum += (coll || 0);
                    }
                    if (collengthsum > maxhig) maxhig = collengthsum;
                    collengthsum = coll || 0;
                }
            }
            for (let al = 1; al <= col; al++) {
                Shot('r', Colposition('r', head, al), higbefor, CutPadding('w', 0, head[`col${al}`] - padding), maxhig);
            }
            heightpage += maxhig;
        }

        if (mode === 'stc') {
            for (let t = 1; t <= col; t++) {
                Shot('fc', r, g, b);
                Shot('r', Colposition('r', head, t), pos, CutPadding('w', 0, head[`col${t}`] - padding), collengthsum * (hig || 1) + 4);
                if (head[`name${t}`]) {
                    for (const c of head[`name${t}`]) {
                        Shot('t', Colposition('t', head, t, 'l'), pos + (collengthsum), TexttoString(c), 'l');
                        collengthsum += (coll || 0);
                    }
                    collengthsum = coll || 0;
                }
            }
            heightpage += (collengthsum * (hig || 1)) + 4;
        }
    }

    function makeMeasurer() {
        return (s = "") => {
            // ใช้ dimensions.w แทน getTextWidth ตรงๆ
            const dim = doc.getTextDimensions(String(s));
            return dim.w; // หน่วยเดียวกับ doc
        };
    }

    /* function splitText(
        text,
        maxWidth,
        maxLinesOrResize,
        {
            locale = 'th',
            preserveEmptyLine = true,
            trimParagraph = true,
            preserveLeadingSpaces = true, // เก็บช่องว่างต้นบรรทัด
            useNbspIndent = true,         // แปลง space ตัวแรกเป็น NBSP เพื่อบังคับอินเดนต์
            tabSize = 4,                  // \t = 4 spaces                      // !! ต้องส่ง doc (jsPDF instance) เข้ามา
        } = {}
    ) {
        const w = makeMeasurer();
        const normalized = String(text ?? '')
            .replace(/<br\s*\/?>/gi, '\n')
            .replace(/\t/g, ' '.repeat(tabSize));

        const paragraphs = normalized.split('\n');
        const segWord = new Intl.Segmenter(locale, { granularity: 'word' });
        const segGrapheme = new Intl.Segmenter(locale, { granularity: 'grapheme' });

        const hardWrap = (token) => {
            const glyphs = Array.from(segGrapheme.segment(token), s => s.segment);
            let out = [], line = '';
            for (const g of glyphs) {
                const test = line + g;
                if (w(test) > maxWidth && line) {
                    out.push(line);
                    line = g;
                } else {
                    line = test;
                }
            }
            if (line) out.push(line);
            return out;
        };

        const applyLeadingNbsp = (s) => {
            if (!useNbspIndent || !s) return s;
            return s.replace(/^ +/, m => m.length ? '\u00A0' + ' '.repeat(m.length - 1) : m);
        };

        const lines = [];

        for (let para of paragraphs) {
            if (trimParagraph) para = para.replace(/\s+$/, '');

            if (para.length === 0) {
                if (preserveEmptyLine) lines.push('');
                continue;
            }

            const tokens = Array.from(segWord.segment(para), s => s.segment);
            let line = '';
            let lineHasContent = false;

            for (let t of tokens) {
                const isOnlySpaces = /^\s+$/.test(t);
                if (!preserveLeadingSpaces && !lineHasContent && isOnlySpaces) continue;

                const test = line + t;

                if (w(test) <= maxWidth) {
                    line = test;
                    if (!isOnlySpaces) lineHasContent = true;
                    continue;
                }

                // ปิดบรรทัดเดิม
                let toPush = line.replace(/\s+$/, '');
                if (preserveLeadingSpaces) toPush = applyLeadingNbsp(toPush);
                if (toPush.length || preserveEmptyLine) lines.push(toPush);

                line = '';
                lineHasContent = false;

                // ถ้า token เดียวก็ยังยาว เก็บแบบ hard wrap
                if (w(t) > maxWidth) {
                    const chunks = hardWrap(t);
                    for (let i = 0; i < chunks.length; i++) {
                        const c = chunks[i];
                        if (i < chunks.length - 1) {
                            let pushNow = c.replace(/\s+$/, '');
                            if (preserveLeadingSpaces) pushNow = applyLeadingNbsp(pushNow);
                            lines.push(pushNow);
                        } else {
                            line = c;
                            lineHasContent = !/^\s*$/.test(c);
                        }
                    }
                } else {
                    line = t;
                    if (!isOnlySpaces) lineHasContent = true;
                }
            }

            if (line || preserveEmptyLine) {
                let toPush = line.replace(/\s+$/, '');
                if (preserveLeadingSpaces) toPush = applyLeadingNbsp(toPush);
                if (toPush.length || preserveEmptyLine) lines.push(toPush);
            }
        }


        return lines;
    } */

    function splitText(
        text,
        maxWidth,
        maxLinesOrResize, // ถ้าเป็น number = จำกัดจำนวนบรรทัด (auto-resize)
        {
            locale = 'th',
            preserveEmptyLine = true,
            trimParagraph = true,
            preserveLeadingSpaces = true,
            useNbspIndent = true,
            tabSize = 4,
            // ค่าปรับขนาด
            minSize = 6,
            step = 1,
        } = {}
    ) {
        const w = makeMeasurer();

        const normalized = String(text ?? '')
            .replace(/<br\s*\/?>/gi, '\n')
            .replace(/\t/g, ' '.repeat(tabSize));

        const segWord = new Intl.Segmenter(locale, { granularity: 'word' });
        const segGrapheme = new Intl.Segmenter(locale, { granularity: 'grapheme' });

        const hardWrap = (token) => {
            const glyphs = Array.from(segGrapheme.segment(token), s => s.segment);
            let out = [], line = '';
            for (const g of glyphs) {
                const test = line + g;
                if (w(test) > maxWidth && line) {
                    out.push(line);
                    line = g;
                } else {
                    line = test;
                }
            }
            if (line) out.push(line);
            return out;
        };

        const applyLeadingNbsp = (s) => {
            if (!useNbspIndent || !s) return s;
            return s.replace(/^ +/, m => (m.length ? '\u00A0' + ' '.repeat(m.length - 1) : m));
        };

        function doSplit(currentText, width) {
            const paragraphs = currentText.split('\n');
            const lines = [];

            for (let para of paragraphs) {
                if (trimParagraph) para = para.replace(/\s+$/, '');

                // ย่อหน้าว่างจริง ๆ จาก \n\n → เก็บไว้ตาม preserveEmptyLine
                if (para.length === 0) {
                    if (preserveEmptyLine) lines.push('');
                    continue;
                }

                const tokens = Array.from(segWord.segment(para), s => s.segment);
                let line = '';
                let lineHasContent = false;

                for (let t of tokens) {
                    const isOnlySpaces = /^\s+$/.test(t);
                    if (!preserveLeadingSpaces && !lineHasContent && isOnlySpaces) continue;

                    const test = line + t;

                    if (w(test) <= width) {
                        line = test;
                        if (!isOnlySpaces) lineHasContent = true;
                        continue;
                    }

                    // -------- ปิดบรรทัดเดิม (เฉพาะถ้ามีเนื้อจริง ๆ แล้ว) --------
                    if (line) {
                        let toPush = line.replace(/\s+$/, '');
                        if (preserveLeadingSpaces) toPush = applyLeadingNbsp(toPush);

                        // ถึงแม้ preserveEmptyLine = true ก็ให้เคารพเฉพาะกรณี para ว่างด้านบน
                        if (toPush.length) {
                            lines.push(toPush);
                        }
                    }
                    line = '';
                    lineHasContent = false;
                    // -------------------------------------------------------------

                    // ถ้า token เดียวก็ยังยาว เก็บแบบ hard wrap
                    if (w(t) > width) {
                        const chunks = hardWrap(t);
                        for (let i = 0; i < chunks.length; i++) {
                            const c = chunks[i];
                            if (i < chunks.length - 1) {
                                let pushNow = c.replace(/\s+$/, '');
                                if (preserveLeadingSpaces) pushNow = applyLeadingNbsp(pushNow);
                                if (pushNow.length) lines.push(pushNow);
                            } else {
                                line = c;
                                lineHasContent = !/^\s*$/.test(c);
                            }
                        }
                    } else {
                        line = t;
                        if (!isOnlySpaces) lineHasContent = true;
                    }
                }

                // ปิดบรรทัดสุดท้ายของย่อหน้า (ต้องมีเนื้อจริง ๆ เท่านั้น)
                if (line && line.trim().length) {
                    let toPush = line.replace(/\s+$/, '');
                    if (preserveLeadingSpaces) toPush = applyLeadingNbsp(toPush);
                    if (toPush.length) {
                        lines.push(toPush);
                    }
                }
            }

            return lines;
        }

        // --- โหมดปกติ: ทำแค่ split แล้วคืน string[] (compat เดิม) ---
        if (typeof maxLinesOrResize !== 'number') {
            return doSplit(normalized, maxWidth);
        }

        const maxLines = maxLinesOrResize;
        const originalSize =
            (doc && doc.internal && typeof doc.internal.getFontSize === 'function'
                ? doc.internal.getFontSize()
                : undefined);

        let hi = (originalSize != null ? originalSize : 16); // เริ่มจากฟอนต์ตอนนี้ (ขอบบน)
        let best = { fontSize: hi, lines: [] };

        // helper วัดจำนวนบรรทัดที่ size ที่กำหนด
        const measureAt = (sz) => {
            Shot('fs', sz);
            return doSplit(normalized, maxWidth);
        };

        try {
            // 1) ถ้าตอนนี้ก็ไม่เกินแล้ว จบเลย
            {
                const linesNow = measureAt(hi);
                best = { fontSize: hi, lines: linesNow };
                if (linesNow.length <= maxLines) return best;
            }

            // 2) หา lower bound โดย "ย่อลงครึ่ง" ไปเรื่อยๆ จนกว่าจะ <= maxLines
            let lo = hi;
            let linesLo = best.lines;
            const MIN_FLOOR = 0.01;  // กันศูนย์ (แต่ไม่ใช่ min จริง แค่กัน error lib)
            const MAX_SHRINK = 60;   // safety stop
            let shrinkStep = 0;

            while (shrinkStep++ < MAX_SHRINK) {
                lo = Math.max(MIN_FLOOR, lo / 2); // ย่อลงครึ่ง
                linesLo = measureAt(lo);
                if (linesLo.length <= maxLines) {
                    break; // ได้ lower bound แล้ว
                }
            }

            // ถ้ายังไม่เข้าเงื่อนไขแม้จะเล็กมาก → เอาค่าสุดท้าย (ถือว่ายัดสุดจริงๆ)
            if (linesLo.length > maxLines) {
                best = { fontSize: lo, lines: linesLo };
                return best;
            }

            // ตอนนี้เรามีช่วง [lo, hi] โดย lo พอดี/ไม่เกิน, hi เกิน
            best = { fontSize: lo, lines: linesLo };

            // 3) Binary search หา "ฟอนต์ใหญ่ที่สุดที่ยังไม่เกิน maxLines"
            const MAX_ITER = 40;
            let iter = 0;

            while (iter++ < MAX_ITER && (hi - lo) > 0.01) {
                const mid = (hi + lo) / 2;
                const linesMid = measureAt(mid);

                if (linesMid.length <= maxLines) {
                    // ใช้ได้ เก็บไว้ แล้วลองขยายขึ้น
                    best = { fontSize: mid, lines: linesMid };
                    lo = mid;
                } else {
                    // เกิน หดลง
                    hi = mid;
                }
            }
        } finally {
            // 4) คืนฟอนต์เดิม
            if (originalSize != null) Shot('fs', originalSize);
        }

        return best;
    }


    function resizeText(
        text,
        maxWidth,
        size,
    ) {

        let width = makeMeasurer();
        let font_size = size || 10;
        Shot("fs", size);
       // console.log('text', text);

       // console.log('width', width(text));
        if (width(text) > maxWidth) {
            font_size--;
            return resizeText(text, maxWidth, font_size);
        } else {
        //    console.log('พอละ ', font_size);

            return font_size;
        }


    }

    return {
        doc,
        Shot,
        ShotTable,
        Colposition,
        CutPadding,
        splitText,
        resizeText,
        addPage,
        setPage,
        utils: { TexttoString, NumtoString, newpage, beFileName },
        state: {
            get y() { return heightpage; },
            set y(v) { heightpage = v; },
            wpage, hpage, endpage
        }
    };
}

module.exports = { createPdfTools };
