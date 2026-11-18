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

// Helper 1: Fast Raw Text Extraction
const extractText = async (filePath, mimeType) => {
    if (mimeType === 'application/pdf' || filePath.toLowerCase().endsWith('.pdf')) {
        const dataBuffer = fs.readFileSync(filePath);
        const parseFn = typeof pdfExtract === 'function' ? pdfExtract : pdfExtract.default;
        const data = await parseFn(dataBuffer);
        return data.text || '';
    } else if (
        mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || 
        filePath.toLowerCase().endsWith('.docx')
    ) {
        const result = await mammoth.extractRawText({ path: filePath });
        return result.value || '';
    } else {
        throw new Error('Unsupported format. Only PDF or DOCX allowed.');
    }
};

// Helper 1B: Extract Rich Structured Blocks from DOCX (Preserves Bold, Italic, Headings, Lists!)
const extractStructuredDocxLines = async (filePath) => {
    try {
        const htmlResult = await mammoth.convertToHtml({ path: filePath });
        const html = htmlResult.value || '';
        if (!html) return null;

        // Split HTML by block-level tags
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

            const hasBold = /<(strong|b)\b/i.test(innerHtml);
            const hasItalic = /<(em|i)\b/i.test(innerHtml);

            blocks.push({
                tag,
                text: tag === 'li' ? `• ${cleanText}` : cleanText,
                hasBold,
                hasItalic
            });
        }
        return blocks.length > 0 ? blocks : null;
    } catch (e) {
        return null;
    }
};

// Helper 2: Smart Typography & Layout Classifier (Reconstructs Academic & Enterprise Formatting)
const classifyLineStyle = (rawLine, lineIndex, isCoverSection, htmlMeta = null) => {
    const trimmed = rawLine.trim();
    if (!trimmed) return { type: 'empty' };

    // 1. Page Break & Major Section Triggers
    const isPageHeaderTrigger = /^(CERTIFICATE|DECLARATION|ACKNOWLEDGEMENT|ACKNOWLEDGEMENTS|ABSTRACT|TABLE OF CONTENTS|LIST OF FIGURES|LIST OF TABLES|CHAPTER\s+\d+|REFERENCES|BIBLIOGRAPHY)/i.test(trimmed);
    const isFooterLine = /\|\s*Page\s*\d+/i.test(trimmed) || /^Page\s*\d+\s*(of\s*\d+)?$/i.test(trimmed);

    // 2. Headings Detection (ALL CAPS or HTML h1-h3)
    const isAllCaps = trimmed.length > 2 && trimmed.length < 90 && trimmed === trimmed.toUpperCase() && /[A-Z]/.test(trimmed);
    const isHtmlHeading = htmlMeta && /^h[1-3]$/.test(htmlMeta.tag);

    // 3. University / Project Report Cover Page Patterns (Center-Aligned)
    const isCoverCenteredPattern = /^(A|An|Project Report|Project\/Internship Report Guidelines|Instructions for Report Writing|On|Submitted|In partial fulfillment|for the award of|Bachelor of Technology|Master of Technology|In Department of|\(Specialization:|July\s*-\s*Dec|Jan\s*-\s*Jun|Session\s*\d{4}|Department of Computer Science|Career Point University|Rajasthan\s*\(India\)|Kota-\d+|Paste University Logo)/i.test(trimmed);

    // 4. Bullet Points & Numbered Lists
    const isBullet = /^([•●▪▸►\-\*]|\d+[\.\)]|\([a-z0-9]+\))\s+/i.test(trimmed) || (htmlMeta && htmlMeta.tag === 'li');

    // 5. Subheadings & Metadata Labels ("Submitted To:", "Submitted By:", "Follow following points:")
    const isSubHeading = /^(Submitted To:?|Submitted By:?|Guided By:?|Supervisor:?|Mr\.|Ms\.|Mrs\.|Dr\.|Prof\.|Follow following points:?|Note:?)/i.test(trimmed) || (trimmed.endsWith(':') && trimmed.length < 65);

    if (isPageHeaderTrigger) {
        return {
            type: 'major_heading',
            pageBreakBefore: lineIndex > 6,
            alignment: AlignmentType.CENTER,
            bold: true,
            italics: false,
            sizePt: 16,
            spaceBefore: 280,
            spaceAfter: 220
        };
    }

    if (isFooterLine) {
        return {
            type: 'footer',
            pageBreakBefore: false,
            alignment: AlignmentType.RIGHT,
            bold: false,
            italics: true,
            sizePt: 10,
            spaceBefore: 160,
            spaceAfter: 240
        };
    }

    if (isCoverCenteredPattern || (isCoverSection && (isAllCaps || trimmed.length < 75))) {
        const isBoldCover = isAllCaps || /^(Project Report|Project\/Internship|Instructions for|Bachelor of Technology|Career Point University|In Department of|Paste University Logo)/i.test(trimmed) || (htmlMeta && htmlMeta.hasBold);
        const isItalicCover = /^(A|Submitted|In partial fulfillment|for the award of|July\s*-|Jan\s*-)/i.test(trimmed) || (htmlMeta && htmlMeta.hasItalic);

        return {
            type: 'cover_centered',
            pageBreakBefore: false,
            alignment: AlignmentType.CENTER,
            bold: isBoldCover,
            italics: isItalicCover,
            sizePt: isAllCaps ? 15 : 13,
            spaceBefore: 110,
            spaceAfter: 110
        };
    }

    if (isAllCaps || isHtmlHeading) {
        return {
            type: 'section_heading',
            pageBreakBefore: false,
            alignment: trimmed.length < 60 ? AlignmentType.CENTER : AlignmentType.LEFT,
            bold: true,
            italics: false,
            sizePt: 14,
            spaceBefore: 220,
            spaceAfter: 130
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
            spaceBefore: 140,
            spaceAfter: 70
        };
    }

    if (isBullet) {
        return {
            type: 'bullet',
            pageBreakBefore: false,
            alignment: AlignmentType.LEFT,
            bold: htmlMeta ? htmlMeta.hasBold : false,
            italics: htmlMeta ? htmlMeta.hasItalic : false,
            sizePt: 12,
            indentLeft: 360,
            spaceBefore: 90,
            spaceAfter: 90
        };
    }

    // Standard Body Paragraph (Justified for full paragraphs, Left for short lines)
    return {
        type: 'body',
        pageBreakBefore: false,
        alignment: trimmed.length > 85 ? AlignmentType.BOTH : AlignmentType.LEFT,
        bold: htmlMeta ? htmlMeta.hasBold : false,
        italics: htmlMeta ? htmlMeta.hasItalic : false,
        sizePt: 12,
        spaceBefore: 100,
        spaceAfter: 110
    };
};

// Helper 3: Deep Formatting, Layout & Lexical Variance Engine
const analyzeDocumentVariance = async (file1Path, file1Name, file2Path, file2Name, preText1 = null, preText2 = null) => {
    const startTime = Date.now();
    const text1 = preText1 !== null ? preText1 : await extractText(file1Path, '');
    const text2 = preText2 !== null ? preText2 : await extractText(file2Path, '');

    // 1. File Size Metrics (KB)
    const stat1 = fs.statSync(file1Path);
    const stat2 = fs.statSync(file2Path);
    const size1KB = +(stat1.size / 1024).toFixed(2);
    const size2KB = +(stat2.size / 1024).toFixed(2);
    const sizeDiffKB = +(size2KB - size1KB).toFixed(2);
    const sizeChangePercent = size1KB > 0 ? +(((size2KB - size1KB) / size1KB) * 100).toFixed(1) : 0;

    // 2. Word & Character Metrics
    const words1 = text1.trim().split(/\s+/).filter(Boolean);
    const words2 = text2.trim().split(/\s+/).filter(Boolean);
    const wordCount1 = words1.length;
    const wordCount2 = words2.length;

    const charCount1 = text1.length;
    const charCount2 = text2.length;
    const charDiff = charCount2 - charCount1;
    const noSpaceChar1 = text1.replace(/\s+/g, '').length;
    const noSpaceChar2 = text2.replace(/\s+/g, '').length;

    // 3. Structural & Formatting Distortion Metrics
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

    // 4. Lexical Word Diff (Hybrid High-Scale Engine)
    const isMassiveDoc = wordCount1 > 20000 || wordCount2 > 20000;
    const PREVIEW_CHAR_LIMIT = 15000;

    let addedWords = 0;
    let removedWords = 0;
    let unchangedWords = 0;
    let previewDifferences = [];

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
        previewDifferences = Diff.diffLines(
            text1.slice(0, PREVIEW_CHAR_LIMIT), 
            text2.slice(0, PREVIEW_CHAR_LIMIT)
        );
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

    // 5. Generate Official Audit Report PDF
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
            .text(`• Original File: ${file1Name} (${size1KB} KB) -> Target File: ${file2Name} (${size2KB} KB)`);
        pdfDoc.moveDown(1.2);

        pdfDoc.fontSize(13).fillColor('#0f172a').text('2. Formatting & Structural Distortion Breakdown', { underline: true });
        pdfDoc.moveDown(0.4);
        pdfDoc.fontSize(11).fillColor('#334155')
            .text(`• Line-Break & Page-Wrap Shift: ${lines1.toLocaleString()} lines -> ${lines2.toLocaleString()} lines (Delta: ${lines2 - lines1})`)
            .text(`• Paragraph / Block Collapse: ${emptyLines1.toLocaleString()} blank spacers -> ${emptyLines2.toLocaleString()} spacers (Lost: ${spacingCollapse})`)
            .text(`• Indentation & Table-Space Loss: ${indentationLoss.toLocaleString()} multi-space alignment blocks altered`)
            .text(`• Whitespace & Character Drift: ${charCount1.toLocaleString()} chars -> ${charCount2.toLocaleString()} chars (Diff: ${charDiff})`)
            .text(`• Words Added: +${addedWords.toLocaleString()} | Words Removed: -${removedWords.toLocaleString()}`);
        pdfDoc.moveDown(1.2);

        pdfDoc.fontSize(13).fillColor('#0f172a').text('3. Auditor Verdict', { underline: true });
        pdfDoc.moveDown(0.4);
        pdfDoc.fontSize(11).fillColor('#b45309').text(
            `STRUCTURAL VARIANCE DETECTED: While ${lexicalSimilarityScore}% of raw words were preserved, conversion induced a ${totalFormattingLoss}% formatting distortion across line wraps, paragraph spacing, and container structure.`
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

// Route 1: Smart Structure-Preserving Convert + Re-Extracted Audit
router.post('/convert', upload.single('file'), async (req, res) => {
    try {
        const file = req.file;
        const { targetFormat } = req.body;
        if (!file) return res.status(400).json({ error: 'Please upload a file' });

        const originalPath = file.path;
        const isSourceDocx = file.mimetype.includes('word') || originalPath.toLowerCase().endsWith('.docx');
        const convertedFilename = `converted-${Date.now()}.${targetFormat}`;
        const convertedPath = `uploads/${convertedFilename}`;

        console.log(`📄 Extracting baseline text & structure from ${file.originalname}...`);
        const extractedText = await extractText(originalPath, file.mimetype);
        const structuredBlocks = isSourceDocx ? await extractStructuredDocxLines(originalPath) : null;

        // Build unified line items with optional HTML style hints
        const rawLines = structuredBlocks
            ? structuredBlocks.map(b => ({ text: b.text, meta: b }))
            : extractedText.split(/\r?\n/).map(l => ({ text: l, meta: null }));

        const MAX_LINES = 15000;
        const linesToProcess = rawLines.length > MAX_LINES ? rawLines.slice(0, MAX_LINES) : rawLines;

        console.log(`⚙️ Smart-Formatting & Converting to .${targetFormat}...`);

        if (targetFormat === 'docx') {
            const docParagraphs = [];
            let emptyStreak = 0;
            let inCoverPage = true;

            linesToProcess.forEach((item, idx) => {
                const rawLine = item.text || '';
                const trimmed = rawLine.trim();

                // Exit cover-page centering once main sections begin
                if (/^(CERTIFICATE|DECLARATION|ACKNOWLEDGEMENT|ABSTRACT|CHAPTER\s+1|Follow following points)/i.test(trimmed)) {
                    inCoverPage = false;
                }

                // Preserve vertical blank spacing (up to 2 consecutive empty lines)
                if (!trimmed) {
                    emptyStreak++;
                    if (emptyStreak <= 2) {
                        docParagraphs.push(new Paragraph({
                            spacing: { before: 60, after: 60 },
                            children: [new TextRun({ text: '', font: 'Times New Roman', size: 24 })]
                        }));
                    }
                    return;
                }
                emptyStreak = 0;

                const style = classifyLineStyle(rawLine, idx, inCoverPage && idx < 35, item.meta);

                if (style.pageBreakBefore) {
                    docParagraphs.push(new Paragraph({ children: [new PageBreak()] }));
                }

                docParagraphs.push(new Paragraph({
                    alignment: style.alignment,
                    indent: style.indentLeft ? { left: style.indentLeft } : undefined,
                    spacing: {
                        before: style.spaceBefore || 100,
                        after: style.spaceAfter || 100,
                        line: 340 // ~1.45x Academic Line Spacing
                    },
                    children: [
                        new TextRun({
                            text: trimmed,
                            font: 'Times New Roman',
                            size: (style.sizePt || 12) * 2,
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
                            // Top: 1", Bottom: 1", Left: 1.5" (2160 twips), Right: 1"
                            margin: { top: 1440, right: 1440, bottom: 1440, left: 2160 }
                        }
                    },
                    children: docParagraphs
                }]
            });

            const buffer = await Packer.toBuffer(doc);
            fs.writeFileSync(convertedPath, buffer);

        } else if (targetFormat === 'pdf') {
            await new Promise((resolve, reject) => {
                // Top: 1" (72pt), Bottom: 1" (72pt), Left: 1.5" (108pt), Right: 1" (72pt)
                const pdfDoc = new PDFDocument({ 
                    bufferPages: false, 
                    size: 'A4',
                    margins: { top: 72, bottom: 72, left: 108, right: 72 } 
                });
                const writeStream = fs.createWriteStream(convertedPath);
                pdfDoc.pipe(writeStream);

                let emptyStreak = 0;
                let inCoverPage = true;

                linesToProcess.slice(0, 25000).forEach((item, idx) => {
                    const rawLine = item.text || '';
                    const trimmed = rawLine.trim();

                    if (/^(CERTIFICATE|DECLARATION|ACKNOWLEDGEMENT|ABSTRACT|CHAPTER\s+1|Follow following points)/i.test(trimmed)) {
                        inCoverPage = false;
                    }

                    if (!trimmed) {
                        emptyStreak++;
                        if (emptyStreak <= 2) pdfDoc.moveDown(0.55);
                        return;
                    }
                    emptyStreak = 0;

                    const style = classifyLineStyle(rawLine, idx, inCoverPage && idx < 35, item.meta);

                    if (style.pageBreakBefore) {
                        pdfDoc.addPage();
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
                        .fontSize(style.sizePt || 12)
                        .text(trimmed, {
                            align: pdfAlign,
                            indent: style.indentLeft ? 18 : 0,
                            lineGap: 4,
                            paragraphGap: 4
                        });
                });

                pdfDoc.end();
                writeStream.on('finish', resolve);
                writeStream.on('error', reject);
            });
        } else {
            return res.status(400).json({ error: 'Invalid target format' });
        }

        console.log(`🔍 Re-extracting converted file to measure formatting preservation...`);
        const reExtractedConvertedText = await extractText(convertedPath, '');

        const auditData = await analyzeDocumentVariance(
            originalPath, 
            file.originalname, 
            convertedPath, 
            convertedFilename,
            extractedText,
            reExtractedConvertedText
        );

        console.log(`✅ Conversion & Smart Layout Audit Complete in ${auditData.processingTimeMs}ms!`);
        res.json({ 
            message: 'Conversion & Smart Layout Audit Successful', 
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

// Route 2: Compare Any 2 Custom Documents
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