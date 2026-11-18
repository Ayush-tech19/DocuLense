const express = require('express');
const multer = require('multer');
const fs = require('fs');
const mammoth = require('mammoth');
const Diff = require('diff');
const PDFDocument = require('pdfkit');
const docx = require('docx');
const pdfExtract = require('pdf-extraction');

const { Document, Packer, Paragraph, TextRun, AlignmentType, PageBreak } = docx;
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

// Helper 1: Universal Text Extractor (Supports PDF, DOCX, and TXT!)
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

// Helper 1B: Extract Rich Structured Blocks from DOCX
const extractStructuredDocxLines = async (filePath) => {
    try {
        const htmlResult = await mammoth.convertToHtml({ path: filePath });
        const html = htmlResult.value || '';
        if (!html) return null;

        const blockRegex = /<(h[1-6]|p|li)[^>]*>([\s\S]*?)<\/\1>/gi;
        const blocks = [];
        let match;

        while ((match = blockRegex.exec(html)) !== null) {
            const tag = match[1].toLowerCase();
            const innerHtml = match[2];
            const cleanText = innerHtml
                .replace(/<br\s*\/?>/gi, '\n')
                .replace(/<[^>]+>/g, '')
                .replace(/&nbsp;/g, ' ')
                .replace(/&amp;/g, '&')
                .replace(/&lt;/g, '<')
                .replace(/&gt;/g, '>')
                .replace(/&quot;/g, '"')
                .trim();

            if (!cleanText) continue;
            blocks.push({
                tag,
                text: tag === 'li' ? `• ${cleanText}` : cleanText,
                hasBold: /<(strong|b)\b/i.test(innerHtml),
                hasItalic: /<(em|i)\b/i.test(innerHtml)
            });
        }
        return blocks.length > 0 ? blocks : null;
    } catch (e) {
        return null;
    }
};

// Helper 2: Smart Paragraph Reflow Engine (Fixes 38 pages turning into 69 pages!)
// Merges broken PDF line-wraps back into natural paragraphs while keeping headings/lists separate
const reflowLinesIntoParagraphs = (rawText) => {
    const rawLines = rawText.split(/\r?\n/);
    const blocks = [];
    let currentPara = '';
    let inCoverArea = true;

    const isStandaloneLine = (trimmed, idx) => {
        if (!trimmed) return true;
        // Stop cover area once main content or TOC starts
        if (/^(CERTIFICATE|DECLARATION|ACKNOWLEDGEMENT|ABSTRACT|TABLE OF CONTENTS|CHAPTER\s+1|Follow following points)/i.test(trimmed)) {
            inCoverArea = false;
        }
        // Keep cover page lines separate (first 28 lines if short)
        if (inCoverArea && idx < 28 && trimmed.length < 75) return true;
        // Major section headers
        if (/^(CERTIFICATE|DECLARATION|ACKNOWLEDGEMENT|ACKNOWLEDGEMENTS|ABSTRACT|TABLE OF CONTENTS|LIST OF FIGURES|LIST OF TABLES|CHAPTER\s+\d+|REFERENCES|BIBLIOGRAPHY)/i.test(trimmed)) return true;
        // TOC lines with dots or trailing page numbers (e.g. "1. Introduction ...... 5" or "Chapter 1   12")
        if (/\.{3,}|\s{3,}\d+$/.test(trimmed)) return true;
        // Bullet points or numbered lists
        if (/^([•●▪▸►\-\*]|\d+[\.\)]|\([a-z0-9]+\))\s+/i.test(trimmed)) return true;
        // Short ALL-CAPS headings or lines ending with colon
        if (trimmed.length < 75 && trimmed === trimmed.toUpperCase() && /[A-Z]/.test(trimmed)) return true;
        if (trimmed.endsWith(':') && trimmed.length < 65) return true;
        // Short metadata labels
        if (/^(Submitted To|Submitted By|Guided By|Supervisor|Mr\.|Ms\.|Dr\.|Prof\.|Department of|Career Point|Rajasthan|Page\s*\d+)/i.test(trimmed) && trimmed.length < 75) return true;

        return false;
    };

    for (let i = 0; i < rawLines.length; i++) {
        const trimmed = rawLines[i].trim();

        // Remove running page footers like "... | Page 2" or "Page 12 of 38" so they don't break paragraphs
        if (/\|\s*Page\s*\d+/i.test(trimmed) || /^Page\s*\d+(\s*of\s*\d+)?$/i.test(trimmed)) {
            continue;
        }

        if (!trimmed) {
            if (currentPara) {
                blocks.push({ text: currentPara, meta: null });
                currentPara = '';
            }
            continue;
        }

        if (isStandaloneLine(trimmed, i)) {
            if (currentPara) {
                blocks.push({ text: currentPara, meta: null });
                currentPara = '';
            }
            blocks.push({ text: trimmed, meta: null });
        } else {
            // Merge wrapped body lines into a single paragraph!
            if (!currentPara) {
                currentPara = trimmed;
            } else {
                currentPara += ' ' + trimmed;
            }
            // If current paragraph ends with sentence punctuation and next line looks like a new sentence/heading, flush if reasonably long
            if (/[\.\!\?]$/.test(trimmed) && currentPara.length > 260) {
                blocks.push({ text: currentPara, meta: null });
                currentPara = '';
            }
        }
    }

    if (currentPara) {
        blocks.push({ text: currentPara, meta: null });
    }

    return blocks;
};

// Helper 3: Smart Typography & Anti-Blank-Page Classifier
const classifyBlockStyle = (rawLine, blockIndex, isCoverSection, blocksSinceLastBreak, htmlMeta = null) => {
    const trimmed = rawLine.trim();
    if (!trimmed) return { type: 'empty' };

    // Detect if this line is a Table of Contents entry (e.g., "CHAPTER 1 ..... 4" or ends with page number)
    const isTocEntry = /\.{2,}|\s{2,}\d+$/.test(trimmed) || (trimmed.length > 28 && /\d+$/.test(trimmed));

    // Strict Major Section Header (ONLY exact standalone titles trigger a Page Break, NEVER TOC lines!)
    const isStrictMajorSection = !isTocEntry && trimmed.length < 40 && /^(CERTIFICATE|DECLARATION|ACKNOWLEDGEMENT|ACKNOWLEDGEMENTS|ABSTRACT|TABLE OF CONTENTS|CHAPTER\s+\d+(\s*[:\-].*)?|REFERENCES|BIBLIOGRAPHY)$/i.test(trimmed);

    const isAllCaps = trimmed.length > 2 && trimmed.length < 85 && trimmed === trimmed.toUpperCase() && /[A-Z]/.test(trimmed);
    const isHtmlHeading = htmlMeta && /^h[1-3]$/.test(htmlMeta.tag);

    const isCoverCenteredPattern = /^(A|An|Project Report|Project\/Internship Report Guidelines|Instructions for Report Writing|On|Submitted|In partial fulfillment|for the award of|Bachelor of Technology|Master of Technology|In Department of|\(Specialization:|July\s*-\s*Dec|Jan\s*-\s*Jun|Session\s*\d{4}|Department of Computer Science|Career Point University|Rajasthan\s*\(India\)|Kota-\d+|Paste University Logo)/i.test(trimmed);

    const isBullet = /^([•●▪▸►\-\*]|\d+[\.\)]|\([a-z0-9]+\))\s+/i.test(trimmed) || (htmlMeta && htmlMeta.tag === 'li');
    const isSubHeading = /^(Submitted To:?|Submitted By:?|Guided By:?|Supervisor:?|Follow following points:?|Note:?)/i.test(trimmed) || (trimmed.endsWith(':') && trimmed.length < 65);

    // Only allow page break if at least 4 content blocks have rendered since the last page break (Prevents Blank Pages!)
    if (isStrictMajorSection) {
        return {
            type: 'major_heading',
            pageBreakBefore: blockIndex > 5 && blocksSinceLastBreak >= 4,
            alignment: AlignmentType.CENTER,
            bold: true,
            italics: false,
            sizePt: 15,
            spaceBefore: 160,
            spaceAfter: 120
        };
    }

    if (isCoverCenteredPattern || (isCoverSection && (isAllCaps || trimmed.length < 70))) {
        const isBoldCover = isAllCaps || /^(Project Report|Project\/Internship|Instructions for|Bachelor of Technology|Career Point University|In Department of|Paste University Logo)/i.test(trimmed) || (htmlMeta && htmlMeta.hasBold);
        const isItalicCover = /^(A|Submitted|In partial fulfillment|for the award of|July\s*-|Jan\s*-)/i.test(trimmed) || (htmlMeta && htmlMeta.hasItalic);

        return {
            type: 'cover_centered',
            pageBreakBefore: false,
            alignment: AlignmentType.CENTER,
            bold: isBoldCover,
            italics: isItalicCover,
            sizePt: isAllCaps ? 14 : 12,
            spaceBefore: 60,
            spaceAfter: 60
        };
    }

    if (isAllCaps || isHtmlHeading) {
        return {
            type: 'section_heading',
            pageBreakBefore: false,
            alignment: trimmed.length < 50 ? AlignmentType.CENTER : AlignmentType.LEFT,
            bold: true,
            italics: false,
            sizePt: 13,
            spaceBefore: 140,
            spaceAfter: 80
        };
    }

    if (isSubHeading) {
        return {
            type: 'subheading',
            pageBreakBefore: false,
            alignment: AlignmentType.LEFT,
            bold: true,
            italics: /^Follow following points/i.test(trimmed) || (htmlMeta && htmlMeta.hasItalic),
            sizePt: 12,
            spaceBefore: 100,
            spaceAfter: 40
        };
    }

    if (isBullet) {
        return {
            type: 'bullet',
            pageBreakBefore: false,
            alignment: AlignmentType.LEFT,
            bold: htmlMeta ? htmlMeta.hasBold : false,
            italics: htmlMeta ? htmlMeta.hasItalic : false,
            sizePt: 11.5,
            indentLeft: 320,
            spaceBefore: 50,
            spaceAfter: 50
        };
    }

    // Standard Compact Body Paragraph (Prevents page inflation!)
    return {
        type: 'body',
        pageBreakBefore: false,
        alignment: trimmed.length > 90 ? AlignmentType.BOTH : AlignmentType.LEFT,
        bold: htmlMeta ? htmlMeta.hasBold : false,
        italics: htmlMeta ? htmlMeta.hasItalic : false,
        sizePt: 11.5,
        spaceBefore: 50,
        spaceAfter: 70
    };
};

// Helper 4: Universal Document Builder (Writes Formatted DOCX, PDF, or TXT)
const generateTargetDocument = async (targetFormat, convertedPath, rawText, structuredBlocks = null) => {
    // Use structured DOCX blocks if available, else reflow raw text paragraphs
    const blocks = structuredBlocks || reflowLinesIntoParagraphs(rawText);
    const MAX_BLOCKS = 18000;
    const blocksToProcess = blocks.length > MAX_BLOCKS ? blocks.slice(0, MAX_BLOCKS) : blocks;

    if (targetFormat === 'txt') {
        const cleanTxt = blocksToProcess.map(b => b.text).join('\n\n');
        fs.writeFileSync(convertedPath, cleanTxt, 'utf8');
        return;
    }

    if (targetFormat === 'docx') {
        const docParagraphs = [];
        let inCoverPage = true;
        let blocksSinceLastBreak = 0;

        blocksToProcess.forEach((item, idx) => {
            const rawLine = item.text || '';
            const trimmed = rawLine.trim();
            if (!trimmed) return;

            if (/^(CERTIFICATE|DECLARATION|ACKNOWLEDGEMENT|ABSTRACT|TABLE OF CONTENTS|CHAPTER\s+1|Follow following points)/i.test(trimmed)) {
                inCoverPage = false;
            }

            const style = classifyBlockStyle(rawLine, idx, inCoverPage && idx < 25, blocksSinceLastBreak, item.meta);

            if (style.pageBreakBefore) {
                docParagraphs.push(new Paragraph({ children: [new PageBreak()] }));
                blocksSinceLastBreak = 0;
            } else {
                blocksSinceLastBreak++;
            }

            docParagraphs.push(new Paragraph({
                alignment: style.alignment,
                indent: style.indentLeft ? { left: style.indentLeft } : undefined,
                spacing: {
                    before: style.spaceBefore || 50,
                    after: style.spaceAfter || 60,
                    line: 276 // 1.15x Compact Professional Spacing (Keeps 38 pages = ~38 pages!)
                },
                children: [
                    new TextRun({
                        text: trimmed,
                        font: 'Times New Roman',
                        size: Math.round((style.sizePt || 11.5) * 2),
                        bold: !!style.bold,
                        italics: !!style.italics
                    })
                ]
            }));
        });

        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: { top: 1150, right: 1150, bottom: 1150, left: 1500 } // Compact Academic Margins
                    }
                },
                children: docParagraphs
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        fs.writeFileSync(convertedPath, buffer);

    } else if (targetFormat === 'pdf') {
        await new Promise((resolve, reject) => {
            const pdfDoc = new PDFDocument({ 
                bufferPages: false, 
                size: 'A4',
                margins: { top: 55, bottom: 55, left: 75, right: 55 } 
            });
            const writeStream = fs.createWriteStream(convertedPath);
            pdfDoc.pipe(writeStream);

            let inCoverPage = true;
            let blocksSinceLastBreak = 0;

            blocksToProcess.forEach((item, idx) => {
                const rawLine = item.text || '';
                const trimmed = rawLine.trim();
                if (!trimmed) return;

                if (/^(CERTIFICATE|DECLARATION|ACKNOWLEDGEMENT|ABSTRACT|TABLE OF CONTENTS|CHAPTER\s+1|Follow following points)/i.test(trimmed)) {
                    inCoverPage = false;
                }

                const style = classifyBlockStyle(rawLine, idx, inCoverPage && idx < 25, blocksSinceLastBreak, item.meta);

                if (style.pageBreakBefore) {
                    pdfDoc.addPage();
                    blocksSinceLastBreak = 0;
                } else {
                    blocksSinceLastBreak++;
                }

                let pdfFont = 'Times-Roman';
                if (style.bold && style.italics) pdfFont = 'Times-BoldItalic';
                else if (style.bold) pdfFont = 'Times-Bold';
                else if (style.italics) pdfFont = 'Times-Italic';

                let pdfAlign = 'left';
                if (style.alignment === AlignmentType.CENTER) pdfAlign = 'center';
                else if (style.alignment === AlignmentType.RIGHT) pdfAlign = 'right';
                else if (style.alignment === AlignmentType.BOTH) pdfAlign = 'justify';

                pdfDoc
                    .font(pdfFont)
                    .fontSize(style.sizePt || 11.5)
                    .text(trimmed, {
                        align: pdfAlign,
                        indent: style.indentLeft ? 15 : 0,
                        lineGap: 2,
                        paragraphGap: 5
                    });
            });

            pdfDoc.end();
            writeStream.on('finish', resolve);
            writeStream.on('error', reject);
        });
    } else {
        throw new Error('Invalid target format. Choose docx, pdf, or txt.');
    }
};

// Helper 5: Deep Formatting, Layout & Lexical Variance Engine
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

    const emptyLines1 = lines1 - nonEmptyParas1;
    const emptyLines2 = lines2 - nonEmptyParas2;
    const spacingCollapse = Math.abs(emptyLines1 - emptyLines2);

    const bullets1 = (text1.match(/[•●▪▸►\-\*\#\_\=\|]/g) || []).length;
    const bullets2 = (text2.match(/[•●▪▸►\-\*\#\_\=\|]/g) || []).length;

    let shiftedLinesCount = 0;
    const minLines = Math.min(nonEmptyParas1, nonEmptyParas2, 5000);
    const cleanLines1 = rawLines1.filter(l => l.trim().length > 0);
    const cleanLines2 = rawLines2.filter(l => l.trim().length > 0);
    for (let i = 0; i < minLines; i++) {
        if (cleanLines1[i].trim() !== cleanLines2[i].trim() || cleanLines1[i].length !== cleanLines2[i].length) {
            shiftedLinesCount++;
        }
    }
    const lineAlignmentShift = shiftedLinesCount + Math.abs(nonEmptyParas1 - nonEmptyParas2);

    const lineBreakPenalty = Math.min(35, (Math.abs(lines1 - lines2) / Math.max(lines1, 1)) * 45);
    const whitespacePenalty = Math.min(25, (Math.abs(charDiff) / Math.max(charCount1, 1)) * 120);
    const indentPenalty = Math.min(20, (indentationLoss / Math.max(nonEmptyParas1, 1)) * 30);
    const containerShiftPenalty = Math.min(15, (Math.abs(sizeChangePercent) / 100) * 12);

    const totalFormattingLoss = +(lineBreakPenalty + whitespacePenalty + indentPenalty + containerShiftPenalty).toFixed(2);
    const formattingIntegrityScore = +Math.max(15, (100 - totalFormattingLoss)).toFixed(2);

    const isMassiveDoc = wordCount1 > 20000 || wordCount2 > 20000;
    const PREVIEW_CHAR_LIMIT = 15000;

    let addedWords = 0, removedWords = 0, unchangedWords = 0, previewDifferences = [];

    if (isMassiveDoc) {
        const freq1 = new Map();
        for (let i = 0; i < wordCount1; i++) {
            const w = words1[i];
            freq1.set(w, (freq1.get(w) || 0) + 1);
        }
        for (let i = 0; i < wordCount2; i++) {
            const w = words2[i];
            const count = freq1.get(w);
            if (count && count > 0) {
                unchangedWords++;
                freq1.set(w, count - 1);
            } else {
                addedWords++;
            }
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
    const similarityScore = +((lexicalSimilarityScore * 0.45) + (formattingIntegrityScore * 0.55)).toFixed(2);
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

        pdfDoc.fontSize(13).fillColor('#0f172a').text('1. Executive Fidelity Summary', { underline: true });
        pdfDoc.moveDown(0.4);
        pdfDoc.fontSize(11).fillColor('#334155')
            .text(`• Overall Document Fidelity Score: ${similarityScore}%`)
            .text(`• Lexical (Text) Preservation: ${lexicalSimilarityScore}%`)
            .text(`• Formatting & Layout Integrity Score: ${formattingIntegrityScore}% (${totalFormattingLoss}% Distortion Detected)`)
            .text(`• Source: ${file1Name} (${size1KB} KB) -> Target: ${file2Name} (${size2KB} KB)`);
        pdfDoc.moveDown(1.2);

        pdfDoc.fontSize(13).fillColor('#0f172a').text('2. Formatting & Structural Distortion Breakdown', { underline: true });
        pdfDoc.moveDown(0.4);
        pdfDoc.fontSize(11).fillColor('#334155')
            .text(`• Line-Break & Paragraph Reflow: ${lines1.toLocaleString()} lines -> ${lines2.toLocaleString()} paragraphs/lines`)
            .text(`• Blank Spacer Adjustment: ${emptyLines1.toLocaleString()} spacers -> ${emptyLines2.toLocaleString()} spacers`)
            .text(`• Indentation & Table-Space Shift: ${indentationLoss.toLocaleString()} multi-space blocks normalized`)
            .text(`• Character Stream Delta: ${charCount1.toLocaleString()} chars -> ${charCount2.toLocaleString()} chars`)
            .text(`• Words Added: +${addedWords.toLocaleString()} | Words Removed: -${removedWords.toLocaleString()}`);
        pdfDoc.moveDown(1.2);

        pdfDoc.fontSize(13).fillColor('#0f172a').text('3. Auditor Verdict', { underline: true });
        pdfDoc.moveDown(0.4);
        pdfDoc.fontSize(11).fillColor('#b45309').text(
            `STRUCTURAL AUDIT COMPLETE: ${lexicalSimilarityScore}% of lexical tokens were preserved while paragraph reflow and container conversion introduced a ${totalFormattingLoss}% structural shift.`
        );

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
            totalDifferences: addedWords + removedWords + Math.abs(lines2 - lines1) + indentationLoss,
            formattingBreakdown: {
                indentationLoss,
                spacingCollapse,
                lineAlignmentShift,
                bulletShift: Math.abs(bullets1 - bullets2)
            },
            fileSize: {
                originalKB: size1KB,
                convertedKB: size2KB,
                diffKB: sizeDiffKB,
                changePercent: sizeChangePercent
            },
            words: {
                original: wordCount1,
                converted: wordCount2,
                diff: wordCount2 - wordCount1
            },
            characters: {
                original: charCount1,
                converted: charCount2,
                diff: charDiff,
                originalNoSpaces: noSpaceChar1,
                convertedNoSpaces: noSpaceChar2
            },
            structure: {
                originalLines: lines1,
                convertedLines: lines2,
                lineDiff: lines2 - lines1,
                originalParagraphs: nonEmptyParas1,
                convertedParagraphs: nonEmptyParas2,
                paragraphDiff: nonEmptyParas2 - nonEmptyParas1
            }
        },
        reportPath,
        detailedDiff: previewDifferences.slice(0, 250)
    };
};

// Route 1: Universal File Converter (PDF ⇄ DOCX ⇄ TXT)
router.post('/convert', upload.single('file'), async (req, res) => {
    try {
        const file = req.file;
        const { targetFormat } = req.body;
        if (!file) return res.status(400).json({ error: 'Please upload a PDF, DOCX, or TXT file' });

        const originalPath = file.path;
        const isSourceDocx = file.mimetype.includes('word') || originalPath.toLowerCase().endsWith('.docx');
        const convertedFilename = `converted-${Date.now()}.${targetFormat}`;
        const convertedPath = `uploads/${convertedFilename}`;

        const extractedText = await extractText(originalPath, file.mimetype);
        const structuredBlocks = isSourceDocx ? await extractStructuredDocxLines(originalPath) : null;

        await generateTargetDocument(targetFormat, convertedPath, extractedText, structuredBlocks);

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
            message: 'Conversion & Smart Reflow Audit Successful', 
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

// Route 2: Direct Raw Text to PDF / Word (.docx) / TXT Generator & Auditor!
router.post('/text-convert', express.json({ limit: '50mb' }), async (req, res) => {
    try {
        const { rawText, targetFormat, title } = req.body;
        if (!rawText || !rawText.trim()) {
            return res.status(400).json({ error: 'Please enter or paste some text to convert.' });
        }

        const safeTitle = (title || 'Direct_Text_Input').replace(/[^a-zA-Z0-9_\-]/g, '_');
        const sourceFilename = `${safeTitle}-${Date.now()}.txt`;
        const sourcePath = `uploads/${sourceFilename}`;
        fs.writeFileSync(sourcePath, rawText, 'utf8');

        const convertedFilename = `${safeTitle}-converted-${Date.now()}.${targetFormat}`;
        const convertedPath = `uploads/${convertedFilename}`;

        await generateTargetDocument(targetFormat, convertedPath, rawText, null);

        const reExtractedConvertedText = await extractText(convertedPath, '');
        const auditData = await analyzeDocumentVariance(
            sourcePath,
            sourceFilename,
            convertedPath,
            convertedFilename,
            rawText,
            reExtractedConvertedText
        );

        res.json({
            message: 'Direct Text Conversion & Audit Successful',
            originalFile: sourceFilename,
            convertedFile: convertedFilename,
            originalPath: sourcePath,
            convertedPath,
            ...auditData
        });
    } catch (error) {
        console.error('Text Convert Error:', error);
        res.status(500).json({ error: 'Text conversion failed', details: error.message });
    }
});

// Route 3: Compare Any 2 Custom Documents (PDF, DOCX, or TXT)
router.post('/compare', upload.array('files', 2), async (req, res) => {
    try {
        const files = req.files;
        if (!files || files.length !== 2) {
            return res.status(400).json({ error: 'Please upload exactly 2 files for comparison' });
        }

        const auditData = await analyzeDocumentVariance(
            files[0].path,
            files[0].originalname,
            files[1].path,
            files[1].originalname
        );

        res.json({ 
            message: 'Comparison Successful', 
            ...auditData
        });

    } catch (error) {
        console.error('Compare Route Error:', error);
        res.status(500).json({ error: 'Error during comparison', details: error.message });
    }
});

module.exports = router;