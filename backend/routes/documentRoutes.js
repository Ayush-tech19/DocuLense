const express = require('express');
const multer = require('multer');
const fs = require('fs');
const mammoth = require('mammoth');
const Diff = require('diff');
const PDFDocument = require('pdfkit');
const docx = require('docx');
const pdfExtract = require('pdf-extraction');

const { Document, Packer, Paragraph, TextRun } = docx;
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

// Helper 1: Fast Text Extraction
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

// Helper 2: Deep Formatting, Layout & Lexical Variance Engine
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

    // 3. Structural & Formatting Distortion Metrics (THE CORE PROJECT USP!)
    const rawLines1 = text1.split(/\r?\n/);
    const rawLines2 = text2.split(/\r?\n/);
    const lines1 = rawLines1.length;
    const lines2 = rawLines2.length;
    const nonEmptyParas1 = rawLines1.filter(l => l.trim().length > 0).length;
    const nonEmptyParas2 = rawLines2.filter(l => l.trim().length > 0).length;

    // Check Indentation / Multi-column spacing loss (e.g., tables, tabs, alignment spaces)
    const multiSpaceMatches1 = (text1.match(/[ \t]{2,}/g) || []).length;
    const multiSpaceMatches2 = (text2.match(/[ \t]{2,}/g) || []).length;
    const indentationLoss = Math.abs(multiSpaceMatches1 - multiSpaceMatches2);

    // Check Blank Line / Spacing collapse
    const emptyLines1 = lines1 - nonEmptyParas1;
    const emptyLines2 = lines2 - nonEmptyParas2;
    const spacingCollapse = Math.abs(emptyLines1 - emptyLines2);

    // Check Punctuation / Bullet / Special Formatting Symbols shift
    const bullets1 = (text1.match(/[•●▪▸►\-\*\#\_\=\|]/g) || []).length;
    const bullets2 = (text2.match(/[•●▪▸►\-\*\#\_\=\|]/g) || []).length;

    // Check Line-Length / Wrapping Variance (How much layout shifted per line)
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

    // Calculate Structural & Formatting Integrity Score (0% to 100%)
    const lineBreakPenalty = Math.min(35, (Math.abs(lines1 - lines2) / Math.max(lines1, 1)) * 45);
    const whitespacePenalty = Math.min(25, (Math.abs(charDiff) / Math.max(charCount1, 1)) * 120);
    const indentPenalty = Math.min(20, (indentationLoss / Math.max(nonEmptyParas1, 1)) * 30);
    const containerShiftPenalty = Math.min(15, (Math.abs(sizeChangePercent) / 100) * 12);

    const totalFormattingLoss = +(lineBreakPenalty + whitespacePenalty + indentPenalty + containerShiftPenalty).toFixed(2);
    const formattingIntegrityScore = +Math.max(15, (100 - totalFormattingLoss)).toFixed(2);

    // 4. Lexical Word Diff
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
        // Use diffLines or diffWordsWithSpace so formatting/line breaks show up in visual diff!
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

    // Overall Composite Document Fidelity Score (Combines Text + Formatting Layout)
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
            similarityScore, // Now reflects True Composite Fidelity (Text + Formatting)!
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

// Route 1: Convert Document + True Re-Extracted Formatting Audit
router.post('/convert', upload.single('file'), async (req, res) => {
    try {
        const file = req.file;
        const { targetFormat } = req.body;
        if (!file) return res.status(400).json({ error: 'Please upload a file' });

        const originalPath = file.path;
        const convertedFilename = `converted-${Date.now()}.${targetFormat}`;
        const convertedPath = `uploads/${convertedFilename}`;

        console.log(`📄 Extracting baseline text from ${file.originalname}...`);
        const extractedText = await extractText(originalPath, file.mimetype);

        console.log(`⚙️ Converting to .${targetFormat}...`);
        if (targetFormat === 'docx') {
            // Realistic Word paragraph conversion (collapses empty PDF line breaks & normalizes spacing)
            const rawLines = extractedText.split(/\r?\n/).filter(l => l.trim().length > 0);
            const MAX_DOCX_PARAS = 12000; 
            const linesToPack = rawLines.length > MAX_DOCX_PARAS ? rawLines.slice(0, MAX_DOCX_PARAS) : rawLines;

            const doc = new Document({
                sections: [{
                    properties: {},
                    children: linesToPack.map(line => new Paragraph({
                        // Word conversion naturally strips multiple raw PDF spaces into single spacing
                        children: [new TextRun(line.replace(/[ \t]{2,}/g, ' ').trim())]
                    }))
                }]
            });
            const buffer = await Packer.toBuffer(doc);
            fs.writeFileSync(convertedPath, buffer);
        } else if (targetFormat === 'pdf') {
            await new Promise((resolve, reject) => {
                const pdfDoc = new PDFDocument({ bufferPages: false, margin: 50 });
                const writeStream = fs.createWriteStream(convertedPath);
                pdfDoc.pipe(writeStream);
                pdfDoc.fontSize(11).text(extractedText.slice(0, 500000) || ' ');
                pdfDoc.end();
                writeStream.on('finish', resolve);
                writeStream.on('error', reject);
            });
        } else {
            return res.status(400).json({ error: 'Invalid target format' });
        }

        console.log(`🔍 Re-extracting converted file to measure real formatting distortion...`);
        // ACTUALLY READ BACK THE CONVERTED FILE FROM DISK (No more fake self-comparison!)
        const reExtractedConvertedText = await extractText(convertedPath, '');

        const auditData = await analyzeDocumentVariance(
            originalPath, 
            file.originalname, 
            convertedPath, 
            convertedFilename,
            extractedText,
            reExtractedConvertedText
        );

        console.log(`✅ Audit Complete! Fidelity: ${auditData.varianceAnalysis.similarityScore}% (Formatting Distortion: ${auditData.varianceAnalysis.totalFormattingLoss}%)`);
        res.json({ 
            message: 'Conversion & Formatting Audit Successful', 
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