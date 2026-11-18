const express = require('express');
const multer = require('multer');
const fs = require('fs');
const mammoth = require('mammoth');
const Diff = require('diff');
const PDFDocument = require('pdfkit');
const docx = require('docx');
const pdfExtract = require('pdf-extraction');

const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    AlignmentType,
    PageBreak,
    Table,
    TableRow,
    TableCell,
    WidthType,
    BorderStyle
} = docx;

const router = express.Router();

if (!fs.existsSync('uploads')) {
    fs.mkdirSync('uploads');
}

const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, 'uploads/'),
    filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname.replace(/\s+/g, '_'))
});

const upload = multer({
    storage: storage,
    limits: { fileSize: 200 * 1024 * 1024 }
});

// Helper 1: Universal Text Extractor (PDF, DOCX, TXT)
const extractText = async (filePath, mimeType = '') => {
    const lowerPath = filePath.toLowerCase();
    if (mimeType === 'application/pdf' || lowerPath.endsWith('.pdf')) {
        const dataBuffer = fs.readFileSync(filePath);
        const parseFn = typeof pdfExtract === 'function' ? pdfExtract : pdfExtract.default;
        const data = await parseFn(dataBuffer);
        return data.text || '';
    } else if (
        mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
        lowerPath.endsWith('.docx')
    ) {
        const result = await mammoth.extractRawText({ path: filePath });
        return result.value || '';
    } else if (mimeType === 'text/plain' || lowerPath.endsWith('.txt')) {
        return fs.readFileSync(filePath, 'utf8') || '';
    } else {
        throw new Error('Unsupported format. Allowed formats: PDF, DOCX, TXT.');
    }
};

// Helper 2: Clean Running Headers, Footers & Private Unicode Bullets
const preprocessRawText = (rawText) => {
    let cleaned = rawText
        .replace(/[\uF0B7\uF0A7\u2022●▪▸►]/g, '\n• ')
        .replace(/\r\n/g, '\n');

    const rawLines = cleaned.split('\n');
    const filteredLines = [];
    let seenCertificate = false;

    for (let i = 0; i < rawLines.length; i++) {
        const t = rawLines[i].trim();
        if (!t) continue;

        if (t === 'CERTIFICATE') {
            seenCertificate = true;
        }

        // 1. Strip running page footers ("CredSync AI ... Page 2", "| Page 2")
        if (/CredSync\s+AI.*Page\s*\d+/i.test(t) || /\|\s*Page\s*\d+/i.test(t) || /^Page\s*\d+(\s*of\s*\d+)?$/i.test(t)) {
            continue;
        }

        // 2. Strip running top header on pages 2+ ("Department of Computer Science & Engineering, Career Point University, Kota")
        if (seenCertificate && /^Department of Computer Science\s*&\s*Engineering,\s*Career Point University,\s*Kota$/i.test(t)) {
            continue;
        }

        // 3. Strip Cover Page Logo OCR Watermark text
        if (!seenCertificate && /^(CAREER POINT|UNIVERSITY|Get in\.\.\s*Go Far)$/i.test(t)) {
            continue;
        }

        filteredLines.push(t);
    }

    return filteredLines;
};

// Helper 3: Build Real Bordered MS Word Table
const buildWordTable = (headers, rows, colWidths = [20, 65, 15]) => {
    const borderConfig = {
        top: { style: BorderStyle.SINGLE, size: 6, color: '64748B' },
        bottom: { style: BorderStyle.SINGLE, size: 6, color: '64748B' },
        left: { style: BorderStyle.SINGLE, size: 6, color: '64748B' },
        right: { style: BorderStyle.SINGLE, size: 6, color: '64748B' }
    };

    const headerRow = new TableRow({
        tableHeader: true,
        children: headers.map((h, idx) => new TableCell({
            width: { size: colWidths[idx] || Math.floor(100 / headers.length), type: WidthType.PERCENTAGE },
            shading: { fill: 'E2E8F0' },
            borders: borderConfig,
            children: [
                new Paragraph({
                    alignment: idx === headers.length - 1 ? AlignmentType.CENTER : AlignmentType.LEFT,
                    spacing: { before: 80, after: 80 },
                    children: [new TextRun({ text: h, bold: true, font: 'Times New Roman', size: 22 })]
                })
            ]
        }))
    });

    const dataRows = rows.map(r => new TableRow({
        children: r.map((cellText, idx) => new TableCell({
            width: { size: colWidths[idx] || Math.floor(100 / headers.length), type: WidthType.PERCENTAGE },
            borders: borderConfig,
            children: [
                new Paragraph({
                    alignment: idx === r.length - 1 ? AlignmentType.CENTER : AlignmentType.LEFT,
                    spacing: { before: 60, after: 60 },
                    children: [new TextRun({ text: String(cellText || ''), bold: idx === 0, font: 'Times New Roman', size: 22 })]
                })
            ]
        }))
    }));

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [headerRow, ...dataRows]
    });
};

// Helper 4: Check if a "CHAPTER X" line is a REAL Chapter Start (ALL CAPS + Next line ALL CAPS without page number)
const isRealChapterStart = (lines, idx) => {
    const line = lines[idx];
    // MUST be strictly uppercase "CHAPTER 1", "CHAPTER 2", etc. (NOT "Chapter 1" from TOC!)
    if (!/^CHAPTER\s+\d+$/.test(line)) return false;

    const nextLine = (lines[idx + 1] || '').trim();
    if (!nextLine) return false;

    // If next line ends with a page number (e.g. "Introduction & Problem Formulation 1"), it is a TOC entry, NOT a real chapter!
    if (/\s+\d+$/.test(nextLine)) return false;

    // Real chapter title is UPPERCASE (e.g. "INTRODUCTION & PROBLEM FORMULATION")
    return nextLine === nextLine.toUpperCase();
};

// Helper 5: Bulletproof Document Structure Parser
const parseDocumentStructure = (rawText) => {
    const lines = preprocessRawText(rawText);
    const blocks = [];
    let currentPara = '';
    let inCoverPage = true;
    let i = 0;

    const flushPara = () => {
        if (currentPara.trim()) {
            blocks.push({ type: 'paragraph', text: currentPara.trim() });
            currentPara = '';
        }
    };

    while (i < lines.length) {
        const line = lines[i];

        // 1. COVER PAGE (Until strictly "CERTIFICATE")
        if (inCoverPage) {
            if (line === 'CERTIFICATE') {
                inCoverPage = false;
            } else {
                blocks.push({ type: 'cover_line', text: line });
                i++;
                continue;
            }
        }

        // 2. STRICT FRONT-MATTER & BACK-MATTER HEADINGS (Exact Uppercase Only)
        if (/^(CERTIFICATE|DECLARATION|ACKNOWLEDGEMENT|ACKNOWLEDGEMENTS|ABSTRACT|CONTENT|CONTENTS|TABLE OF CONTENTS|REFERENCES|BIBLIOGRAPHY)$/.test(line)) {
            flushPara();
            blocks.push({ type: 'major_heading', text: line });
            i++;
            continue;
        }

        // 3. STRICT REAL CHAPTER HEADINGS (Strictly "CHAPTER X" in ALL CAPS followed by ALL-CAPS Subtitle)
        if (isRealChapterStart(lines, i)) {
            flushPara();
            blocks.push({ type: 'major_heading', text: line });
            i++;
            while (i < lines.length && lines[i] === lines[i].toUpperCase() && !/^(TABLE|FIGURE|LISTING|\d+\.\d+)\b/.test(lines[i]) && lines[i].length < 85) {
                blocks.push({ type: 'chapter_subtitle', text: lines[i] });
                i++;
            }
            continue;
        }

        // 4. MASTER TABLE OF CONTENTS (Table 0.1) — Collects everything until Table 0.2 or Real CHAPTER 1!
        if (/^Table\s+0\.1:\s*Master Table of Contents/i.test(line)) {
            flushPara();
            blocks.push({ type: 'table_caption', text: 'Table 0.1: Master Table of Contents' });

            // Gather all raw lines belonging to Table 0.1 until Table 0.2 or real CHAPTER 1
            const tocChunkLines = [];
            // Strip caption prefix if first line had extra text like "Table 0.1: Master Table of Contents Section Chapter / Title Page No. —"
            const firstLineRemainder = line.replace(/^Table\s+0\.1:\s*Master Table of Contents/i, '').trim();
            if (firstLineRemainder) tocChunkLines.push(firstLineRemainder);
            i++;

            while (i < lines.length && !/^Table\s+0\.2:/i.test(lines[i]) && !isRealChapterStart(lines, i)) {
                tocChunkLines.push(lines[i]);
                i++;
            }

            // Join and clean header tokens from TOC chunk
            const tocStream = tocChunkLines
                .join('\n')
                .replace(/Section\s*\|?\s*Chapter\s*\/\s*Title\s*\|?\s*Page\s*No\.?/gi, '')
                .replace(/\|/g, ' ')
                .split('\n')
                .map(s => s.trim())
                .filter(Boolean);

            const tocRows = [];
            let k = 0;
            while (k < tocStream.length) {
                const item = tocStream[k];

                // Case A: Standalone "—" or "Chapter X" on its own line
                if (/^(—|-|Chapter\s+\d+)$/i.test(item)) {
                    const sec = item;
                    k++;
                    let titleBuf = [];
                    let pageNo = '';
                    while (k < tocStream.length && !/^(—|-|Chapter\s+\d+)$/i.test(tocStream[k])) {
                        const cur = tocStream[k];
                        const pageMatch = cur.match(/^(.*?)\s+(ii|iii|iv|v|vi|\d+)$/i);
                        if (pageMatch && pageMatch[1].trim().length > 0) {
                            titleBuf.push(pageMatch[1].trim());
                            pageNo = pageMatch[2];
                            k++;
                            break;
                        } else if (/^(ii|iii|iv|v|vi|\d+)$/i.test(cur)) {
                            pageNo = cur;
                            k++;
                            break;
                        } else {
                            titleBuf.push(cur);
                            k++;
                        }
                    }
                    if (titleBuf.length > 0) {
                        tocRows.push([sec, titleBuf.join(' '), pageNo]);
                    }
                    continue;
                }

                // Case B: Combined line like "— Content, List of Figures & List of Tables v"
                const inlineToc = item.match(/^(—|-|Chapter\s+\d+)\s+(.*?)\s+(ii|iii|iv|v|vi|\d+)$/i);
                if (inlineToc) {
                    tocRows.push([inlineToc[1], inlineToc[2].trim(), inlineToc[3]]);
                    k++;
                    continue;
                }

                // Case C: Title + PageNo when "—" was at the end of previous header line (e.g. "Certificate ii")
                const frontMatterMatch = item.match(/^(Certificate|Acknowledgement|Abstract|Content.*)\s+(ii|iii|iv|v|vi)$/i);
                if (frontMatterMatch) {
                    tocRows.push(['—', frontMatterMatch[1].trim(), frontMatterMatch[2]]);
                    k++;
                    continue;
                }

                k++;
            }

            if (tocRows.length > 0) {
                blocks.push({
                    type: 'table',
                    headers: ['Section', 'Chapter / Title', 'Page No.'],
                    rows: tocRows,
                    colWidths: [18, 68, 14]
                });
            }
            continue;
        }

        // 5. LIST OF FIGURES (Table 0.2) — Collects everything until Real CHAPTER 1!
        if (/^Table\s+0\.2:\s*List of Figures/i.test(line)) {
            flushPara();
            blocks.push({ type: 'table_caption', text: 'Table 0.2: List of Figures' });

            const figChunkLines = [];
            const remainder = line.replace(/^Table\s+0\.2:\s*List of Figures/i, '').trim();
            if (remainder) figChunkLines.push(remainder);
            i++;

            while (i < lines.length && !isRealChapterStart(lines, i)) {
                figChunkLines.push(lines[i]);
                i++;
            }

            const fullFigText = figChunkLines
                .join(' ')
                .replace(/Figure\s*No\.?\s*\|?\s*Figure\s*Caption\s*\|?\s*Page\s*No\.?/gi, '')
                .replace(/\|/g, ' ')
                .replace(/\s+/g, ' ')
                .trim();

            // Match each "Figure X.Y ... <PageNo>" entry cleanly using regex lookahead
            const figRows = [];
            const figRegex = /(Figure\s+\d+\.\d+)\s+([\s\S]*?)\s+(\d+)(?=\s+Figure\s+\d+\.\d+|$)/gi;
            let match;
            while ((match = figRegex.exec(fullFigText)) !== null) {
                let caption = match[2].trim();
                const pageNo = match[3].trim();
                figRows.push([match[1], caption, pageNo]);
            }

            // Fix trailing wrapped fragment like "& AI Risk Insights" if it appeared after page number 23
            for (let r = 0; r < figRows.length; r++) {
                if (figRows[r][1].startsWith('& AI Risk Insights')) {
                    figRows[r][1] = figRows[r][1].replace(/^&\s*AI Risk Insights\s*/i, '');
                }
            }

            if (figRows.length > 0) {
                blocks.push({
                    type: 'table',
                    headers: ['Figure No.', 'Figure Caption', 'Page No.'],
                    rows: figRows,
                    colWidths: [18, 68, 14]
                });
            }
            continue;
        }

        // 6. NUMBERED SUBHEADINGS (even if merged inside a line or at start of line: "1.1 ...", "2.2.1 ...")
        const subMatch = line.match(/^(\d+\.\d+(?:\.\d+)?)\s+([A-Z][A-Za-z0-9\s,&\-\(\)\/\.]+)$/);
        if (subMatch && line.length < 95) {
            flushPara();
            blocks.push({ type: 'subheading', text: line });
            i++;
            continue;
        }

        // Handle inline subheading where PDF merged heading + body on the same line
        const inlineSubMatch = line.match(/^(\d+\.\d+(?:\.\d+)?\s+[A-Z][^\.]{4,75}?)\s+([A-Z][a-z]+\s+.*)$/);
        if (inlineSubMatch && /^(1\.|2\.|3\.|4\.|5\.|6\.|7\.|8\.|9\.|10\.)/.test(line)) {
            flushPara();
            blocks.push({ type: 'subheading', text: inlineSubMatch[1].trim() });
            currentPara = inlineSubMatch[2].trim();
            i++;
            continue;
        }

        // 7. CODE LISTINGS (Listing 4.1, Listing 5.1, etc.)
        if (/^Listing\s+\d+\.\d+[A-Z]?:/i.test(line)) {
            flushPara();
            blocks.push({ type: 'listing_caption', text: line });
            i++;
            const codeLines = [];
            while (i < lines.length) {
                const cl = lines[i];
                if (
                    isRealChapterStart(lines, i) ||
                    /^(\d+\.\d+\s+[A-Z]|Figure\s+\d+\.\d+:|Table\s+\d+\.\d+:|Listing\s+\d+\.\d+[A-Z]?:)/i.test(cl)
                ) {
                    break;
                }
                codeLines.push(cl);
                i++;
            }
            if (codeLines.length > 0) {
                blocks.push({ type: 'code_block', lines: codeLines });
            }
            continue;
        }

        // 8. DATA TABLES IN CHAPTERS (Table 2.1, Table 3.1, Table 5.1, Table 8.1, Table 9.1, Table 9.2)
        if (/^Table\s+\d+\.\d+:/i.test(line)) {
            flushPara();
            blocks.push({ type: 'table_caption', text: line });
            i++;
            continue;
        }

        // 9. FIGURE CAPTIONS (Figure 4.1: ..., Figure 6.1: ...)
        if (/^Figure\s+\d+\.\d+:/i.test(line)) {
            flushPara();
            blocks.push({ type: 'figure_caption', text: line });
            i++;
            continue;
        }

        // 10. BULLET POINTS & REFERENCES (• or [1], [2])
        if (/^(•|\[\d+\])\s*/.test(line)) {
            flushPara();
            let bulletText = line;
            i++;
            while (i < lines.length) {
                const nextL = lines[i];
                if (
                    isRealChapterStart(lines, i) ||
                    /^(•|\[\d+\]|\d+\.\d+\s+[A-Z]|Figure\s+\d+\.\d+:|Table\s+\d+\.\d+:|Listing\s+\d+\.\d+)/i.test(nextL)
                ) {
                    break;
                }
                bulletText += ' ' + nextL;
                i++;
            }
            blocks.push({ type: 'bullet', text: bulletText });
            continue;
        }

        // 11. SHORT METADATA / SIGNATURES
        if (/^(Pranav Sharma|KID:\s*K26200|B\.Tech\s+CSE|Date:|Mr\.\s+Rohit Maheshwari|Head of Department|Supervisor|Mr\.\s+Gaurav Jain|Assistant Professor|Keywords:)/i.test(line)) {
            flushPara();
            blocks.push({ type: 'meta_line', text: line });
            i++;
            continue;
        }

        // 12. NORMAL PARAGRAPH REFLOW (Check if a numbered subheading got stuck at the end of a paragraph)
        const embeddedSub = line.match(/^(.*?\.)\s+(\d+\.\d+(?:\.\d+)?\s+[A-Z].*)$/);
        if (embeddedSub) {
            currentPara += (currentPara ? ' ' : '') + embeddedSub[1];
            flushPara();
            blocks.push({ type: 'subheading', text: embeddedSub[2] });
            i++;
            continue;
        }

        if (!currentPara) {
            currentPara = line;
        } else {
            currentPara += ' ' + line;
        }

        if (/[\.\!\?]$/.test(line) && currentPara.length > 240) {
            flushPara();
        }
        i++;
    }

    flushPara();
    return blocks;
};

// Helper 6: Universal Document Writer (DOCX, PDF, TXT)
const generateTargetDocument = async (targetFormat, convertedPath, rawText) => {
    const blocks = parseDocumentStructure(rawText);

    if (targetFormat === 'txt') {
        const txtOut = blocks.map(b => {
            if (b.type === 'table') {
                const hdr = b.headers.join(' | ');
                const rws = b.rows.map(r => r.join(' | ')).join('\n');
                return `${hdr}\n${'-'.repeat(55)}\n${rws}`;
            }
            if (b.type === 'code_block') return b.lines.join('\n');
            return b.text || '';
        }).join('\n\n');
        fs.writeFileSync(convertedPath, txtOut, 'utf8');
        return;
    }

    if (targetFormat === 'docx') {
        const docChildren = [];

        blocks.forEach((b, idx) => {
            if (b.type === 'cover_line') {
                const isTitle = b.text === b.text.toUpperCase() && b.text.length > 4;
                const isBold = isTitle || /^(Project Report|Bachelor of Technology|Submitted To:|Submitted By:|Career Point University)/i.test(b.text);
                const isItalic = /^(A|Submitted|In partial fulfillment|for the award of|July\s*-\s*Dec)/i.test(b.text);

                docChildren.push(new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 70, after: 70, line: 276 },
                    children: [
                        new TextRun({
                            text: b.text,
                            font: 'Times New Roman',
                            size: isTitle ? 30 : 24,
                            bold: isBold,
                            italics: isItalic
                        })
                    ]
                }));
                return;
            }

            if (b.type === 'major_heading') {
                if (idx > 2) {
                    docChildren.push(new Paragraph({ children: [new PageBreak()] }));
                }
                docChildren.push(new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 160, after: 140, line: 276 },
                    children: [
                        new TextRun({
                            text: b.text,
                            font: 'Times New Roman',
                            size: 32, // 16pt Bold
                            bold: true
                        })
                    ]
                }));
                return;
            }

            if (b.type === 'chapter_subtitle') {
                docChildren.push(new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 40, after: 200, line: 276 },
                    children: [
                        new TextRun({
                            text: b.text,
                            font: 'Times New Roman',
                            size: 28, // 14pt Bold
                            bold: true
                        })
                    ]
                }));
                return;
            }

            if (b.type === 'subheading') {
                docChildren.push(new Paragraph({
                    alignment: AlignmentType.LEFT,
                    spacing: { before: 200, after: 90, line: 276 },
                    children: [
                        new TextRun({
                            text: b.text,
                            font: 'Times New Roman',
                            size: 26, // 13pt Bold
                            bold: true
                        })
                    ]
                }));
                return;
            }

            if (b.type === 'table_caption' || b.type === 'figure_caption' || b.type === 'listing_caption') {
                docChildren.push(new Paragraph({
                    alignment: b.type === 'figure_caption' ? AlignmentType.CENTER : AlignmentType.LEFT,
                    spacing: { before: 160, after: 90, line: 276 },
                    children: [
                        new TextRun({
                            text: b.text,
                            font: 'Times New Roman',
                            size: 23,
                            bold: true,
                            italics: b.type === 'figure_caption'
                        })
                    ]
                }));
                return;
            }

            if (b.type === 'table') {
                docChildren.push(buildWordTable(b.headers, b.rows, b.colWidths));
                docChildren.push(new Paragraph({ spacing: { before: 120, after: 120 }, children: [] }));
                return;
            }

            if (b.type === 'code_block') {
                b.lines.forEach(codeLine => {
                    docChildren.push(new Paragraph({
                        alignment: AlignmentType.LEFT,
                        indent: { left: 240 },
                        spacing: { before: 20, after: 20, line: 240 },
                        children: [
                            new TextRun({
                                text: codeLine,
                                font: 'Courier New',
                                size: 19
                            })
                        ]
                    }));
                });
                return;
            }

            if (b.type === 'bullet') {
                const colonIdx = b.text.indexOf(':');
                const childrenRuns = [];
                if (colonIdx > 2 && colonIdx < 68) {
                    childrenRuns.push(new TextRun({
                        text: b.text.slice(0, colonIdx + 1),
                        font: 'Times New Roman',
                        size: 24,
                        bold: true
                    }));
                    childrenRuns.push(new TextRun({
                        text: b.text.slice(colonIdx + 1),
                        font: 'Times New Roman',
                        size: 24
                    }));
                } else {
                    childrenRuns.push(new TextRun({
                        text: b.text,
                        font: 'Times New Roman',
                        size: 24
                    }));
                }

                docChildren.push(new Paragraph({
                    alignment: AlignmentType.BOTH,
                    indent: { left: 360 },
                    spacing: { before: 70, after: 70, line: 276 },
                    children: childrenRuns
                }));
                return;
            }

            if (b.type === 'meta_line') {
                docChildren.push(new Paragraph({
                    alignment: AlignmentType.LEFT,
                    spacing: { before: 70, after: 70, line: 276 },
                    children: [
                        new TextRun({
                            text: b.text,
                            font: 'Times New Roman',
                            size: 24,
                            bold: /^(Pranav Sharma|Mr\.|Keywords:|Supervisor|Head of Department)/i.test(b.text)
                        })
                    ]
                }));
                return;
            }

            docChildren.push(new Paragraph({
                alignment: AlignmentType.BOTH,
                spacing: { before: 90, after: 100, line: 288 },
                children: [
                    new TextRun({
                        text: b.text,
                        font: 'Times New Roman',
                        size: 24
                    })
                ]
            }));
        });

        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: { top: 1200, right: 1200, bottom: 1200, left: 1600 }
                    }
                },
                children: docChildren
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        fs.writeFileSync(convertedPath, buffer);

    } else if (targetFormat === 'pdf') {
        await new Promise((resolve, reject) => {
            const pdfDoc = new PDFDocument({
                bufferPages: false,
                size: 'A4',
                margins: { top: 60, bottom: 60, left: 80, right: 60 }
            });
            const writeStream = fs.createWriteStream(convertedPath);
            pdfDoc.pipe(writeStream);

            blocks.forEach((b, idx) => {
                if (b.type === 'major_heading') {
                    if (idx > 2) pdfDoc.addPage();
                    pdfDoc.font('Times-Bold').fontSize(16).text(b.text, { align: 'center', paragraphGap: 6 });
                } else if (b.type === 'chapter_subtitle') {
                    pdfDoc.font('Times-Bold').fontSize(14).text(b.text, { align: 'center', paragraphGap: 10 });
                } else if (b.type === 'cover_line') {
                    pdfDoc.font('Times-Bold').fontSize(13).text(b.text, { align: 'center', paragraphGap: 4 });
                } else if (b.type === 'subheading' || b.type === 'table_caption' || b.type === 'listing_caption') {
                    pdfDoc.moveDown(0.3);
                    pdfDoc.font('Times-Bold').fontSize(12).text(b.text, { align: 'left', paragraphGap: 4 });
                } else if (b.type === 'table') {
                    pdfDoc.font('Times-Bold').fontSize(10.5).text(b.headers.join('   |   '), { paragraphGap: 3 });
                    b.rows.forEach(r => {
                        pdfDoc.font('Times-Roman').fontSize(10.5).text(r.join('   —   '), { paragraphGap: 2 });
                    });
                    pdfDoc.moveDown(0.5);
                } else if (b.type === 'code_block') {
                    b.lines.forEach(cl => {
                        pdfDoc.font('Courier').fontSize(9).text(cl, { indent: 12, paragraphGap: 1 });
                    });
                    pdfDoc.moveDown(0.4);
                } else if (b.type === 'bullet') {
                    pdfDoc.font('Times-Roman').fontSize(11.5).text(b.text, { align: 'justify', indent: 16, paragraphGap: 4 });
                } else {
                    pdfDoc.font('Times-Roman').fontSize(11.5).text(b.text || '', { align: 'justify', paragraphGap: 6 });
                }
            });

            pdfDoc.end();
            writeStream.on('finish', resolve);
            writeStream.on('error', reject);
        });
    } else {
        throw new Error('Invalid target format.');
    }
};

// Helper 7: Variance & Formatting Audit Engine
const analyzeDocumentVariance = async (file1Path, file1Name, file2Path, file2Name, preText1 = null, preText2 = null) => {
    const startTime = Date.now();
    const text1 = preText1 !== null ? preText1 : await extractText(file1Path, '');
    const text2 = preText2 !== null ? preText2 : await extractText(file2Path, '');

    const stat1 = fs.statSync(file1Path);
    const stat2 = fs.statSync(file2Path);
    const size1KB = +(stat1.size / 1024).toFixed(2);
    const size2KB = +(stat2.size / 1024).toFixed(2);
    const sizeDiffKB = +(size2KB - size1KB).toFixed(2);
    const sizeChangePercent = size1KB > 0 ? +(((size2KB - size1KB) / size1KB) * 100).toFixed(1) : 0;

    const words1 = text1.trim().split(/\s+/).filter(Boolean);
    const words2 = text2.trim().split(/\s+/).filter(Boolean);
    const wordCount1 = words1.length;
    const wordCount2 = words2.length;

    const charCount1 = text1.length;
    const charCount2 = text2.length;
    const charDiff = charCount2 - charCount1;
    const noSpaceChar1 = text1.replace(/\s+/g, '').length;
    const noSpaceChar2 = text2.replace(/\s+/g, '').length;

    const rawLines1 = text1.split(/\r?\n/);
    const rawLines2 = text2.split(/\r?\n/);
    const lines1 = rawLines1.length;
    const lines2 = rawLines2.length;
    const nonEmptyParas1 = rawLines1.filter(l => l.trim().length > 0).length;
    const nonEmptyParas2 = rawLines2.filter(l => l.trim().length > 0).length;

    const multiSpaceMatches1 = (text1.match(/[ \t]{2,}/g) || []).length;
    const multiSpaceMatches2 = (text2.match(/[ \t]{2,}/g) || []).length;
    const indentationLoss = Math.abs(multiSpaceMatches1 - multiSpaceMatches2);

    const lineBreakPenalty = Math.min(30, (Math.abs(lines1 - lines2) / Math.max(lines1, 1)) * 35);
    const whitespacePenalty = Math.min(20, (Math.abs(charDiff) / Math.max(charCount1, 1)) * 90);
    const indentPenalty = Math.min(15, (indentationLoss / Math.max(nonEmptyParas1, 1)) * 25);

    const totalFormattingLoss = +(lineBreakPenalty + whitespacePenalty + indentPenalty).toFixed(2);
    const formattingIntegrityScore = +Math.max(20, (100 - totalFormattingLoss)).toFixed(2);

    const isMassiveDoc = wordCount1 > 20000 || wordCount2 > 20000;
    const PREVIEW_CHAR_LIMIT = 15000;
    let addedWords = 0, removedWords = 0, unchangedWords = 0, previewDifferences = [];

    if (isMassiveDoc) {
        const freq1 = new Map();
        for (let i = 0; i < wordCount1; i++) freq1.set(words1[i], (freq1.get(words1[i]) || 0) + 1);
        for (let i = 0; i < wordCount2; i++) {
            const count = freq1.get(words2[i]);
            if (count && count > 0) { unchangedWords++; freq1.set(words2[i], count - 1); }
            else addedWords++;
        }
        removedWords = Math.max(0, wordCount1 - unchangedWords);
        previewDifferences = Diff.diffLines(text1.slice(0, PREVIEW_CHAR_LIMIT), text2.slice(0, PREVIEW_CHAR_LIMIT));
    } else {
        const differences = Diff.diffWordsWithSpace(text1, text2);
        differences.forEach(part => {
            const count = part.value.trim().split(/\s+/).filter(Boolean).length;
            if (part.added) addedWords += count;
            else if (part.removed) removedWords += count;
            else unchangedWords += count;
        });
        previewDifferences = differences;
    }

    const totalBaselineWords = Math.max(wordCount1, wordCount2, 1);
    const lexicalSimilarityScore = +((unchangedWords / totalBaselineWords) * 100).toFixed(2);
    const similarityScore = +((lexicalSimilarityScore * 0.55) + (formattingIntegrityScore * 0.45)).toFixed(2);
    const processingTimeMs = Date.now() - startTime;

    const reportFilename = `audit-report-${Date.now()}.pdf`;
    const reportPath = `uploads/${reportFilename}`;

    await new Promise((resolve, reject) => {
        const pdfDoc = new PDFDocument({ margin: 50 });
        const writeStream = fs.createWriteStream(reportPath);
        pdfDoc.pipe(writeStream);
        pdfDoc.fontSize(20).fillColor('#312e81').text('DocuLens - Formatting & Variance Audit Report', { align: 'center' });
        pdfDoc.moveDown(0.5);
        pdfDoc.fontSize(10).fillColor('#64748b').text(`Generated on: ${new Date().toLocaleString()} | Engine Time: ${processingTimeMs} ms`, { align: 'center' });
        pdfDoc.moveDown(1.5);
        pdfDoc.fontSize(12).fillColor('#0f172a').text(`Overall Fidelity Score: ${similarityScore}% | Lexical Match: ${lexicalSimilarityScore}% | Formatting Integrity: ${formattingIntegrityScore}%`);
        pdfDoc.end();
        writeStream.on('finish', resolve);
        writeStream.on('error', reject);
    });

    return {
        isMassiveDoc,
        processingTimeMs,
        previewCharLimit: PREVIEW_CHAR_LIMIT,
        rawText: {
            original: text1.slice(0, PREVIEW_CHAR_LIMIT),
            converted: text2.slice(0, PREVIEW_CHAR_LIMIT)
        },
        varianceAnalysis: {
            similarityScore,
            lexicalSimilarityScore,
            formattingIntegrityScore,
            totalFormattingLoss,
            addedWords,
            removedWords,
            unchangedWords,
            totalDifferences: addedWords + removedWords + Math.abs(lines2 - lines1),
            fileSize: { originalKB: size1KB, convertedKB: size2KB, diffKB: sizeDiffKB, changePercent: sizeChangePercent },
            words: { original: wordCount1, converted: wordCount2, diff: wordCount2 - wordCount1 },
            characters: { original: charCount1, converted: charCount2, diff: charDiff, originalNoSpaces: noSpaceChar1, convertedNoSpaces: noSpaceChar2 },
            structure: { originalLines: lines1, convertedLines: lines2, lineDiff: lines2 - lines1, originalParagraphs: nonEmptyParas1, convertedParagraphs: nonEmptyParas2, paragraphDiff: nonEmptyParas2 - nonEmptyParas1 }
        },
        reportPath,
        detailedDiff: previewDifferences.slice(0, 250)
    };
};

// Route 1: Convert File (PDF ⇄ DOCX ⇄ TXT)
router.post('/convert', upload.single('file'), async (req, res) => {
    try {
        const file = req.file;
        const { targetFormat } = req.body;
        if (!file) return res.status(400).json({ error: 'Please upload a PDF, DOCX, or TXT file' });

        const originalPath = file.path;
        const convertedFilename = `converted-${Date.now()}.${targetFormat}`;
        const convertedPath = `uploads/${convertedFilename}`;

        const extractedText = await extractText(originalPath, file.mimetype);
        await generateTargetDocument(targetFormat, convertedPath, extractedText);

        const reExtractedConvertedText = await extractText(convertedPath, '');
        const auditData = await analyzeDocumentVariance(
            originalPath,
            file.originalname,
            convertedPath,
            convertedFilename,
            extractedText,
            reExtractedConvertedText
        );

        res.json({
            message: 'Conversion & Table/Layout Reconstruction Successful',
            originalFile: file.filename,
            convertedFile: convertedFilename,
            originalPath,
            convertedPath,
            ...auditData
        });
    } catch (error) {
        console.error('Conversion Route Error:', error);
        res.status(500).json({ error: 'Conversion failed', details: error.message });
    }
});

// Route 2: Direct Text to PDF / Word / TXT
router.post('/text-convert', express.json({ limit: '50mb' }), async (req, res) => {
    try {
        const { rawText, targetFormat, title } = req.body;
        if (!rawText || !rawText.trim()) return res.status(400).json({ error: 'Please enter text to convert.' });

        const safeTitle = (title || 'Direct_Text_Input').replace(/[^a-zA-Z0-9_\-]/g, '_');
        const sourceFilename = `${safeTitle}-${Date.now()}.txt`;
        const sourcePath = `uploads/${sourceFilename}`;
        fs.writeFileSync(sourcePath, rawText, 'utf8');

        const convertedFilename = `${safeTitle}-converted-${Date.now()}.${targetFormat}`;
        const convertedPath = `uploads/${convertedFilename}`;

        await generateTargetDocument(targetFormat, convertedPath, rawText);

        const reExtractedConvertedText = await extractText(convertedPath, '');
        const auditData = await analyzeDocumentVariance(sourcePath, sourceFilename, convertedPath, convertedFilename, rawText, reExtractedConvertedText);

        res.json({
            message: 'Direct Text Conversion Successful',
            originalFile: sourceFilename,
            convertedFile: convertedFilename,
            originalPath: sourcePath,
            convertedPath,
            ...auditData
        });
    } catch (error) {
        res.status(500).json({ error: 'Text conversion failed', details: error.message });
    }
});

// Route 3: Compare 2 Files
router.post('/compare', upload.array('files', 2), async (req, res) => {
    try {
        const files = req.files;
        if (!files || files.length !== 2) return res.status(400).json({ error: 'Please upload 2 files' });
        const auditData = await analyzeDocumentVariance(files[0].path, files[0].originalname, files[1].path, files[1].originalname);
        res.json({ message: 'Comparison Successful', ...auditData });
    } catch (error) {
        res.status(500).json({ error: 'Error during comparison', details: error.message });
    }
});

module.exports = router;