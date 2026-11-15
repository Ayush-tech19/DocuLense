const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');

const TOTAL_PAGES = 2500; // Tu isko 3000 bhi kar sakta hai
const outputPath = path.join(__dirname, `Mega_Enterprise_Test_${TOTAL_PAGES}_Pages.pdf`);

console.log(`🚀 Generating ${TOTAL_PAGES}-page heavy PDF... Please wait 10-15 seconds...`);
const startTime = Date.now();

const doc = new PDFDocument({ autoFirstPage: false, bufferPages: false });
const stream = fs.createWriteStream(outputPath);
doc.pipe(stream);

const sampleParagraphs = [
    "Enterprise Variance Audit Log: Section A outlines distributed microservice architecture, financial reconciliation pipelines, and automated compliance checks across multi-region cloud clusters.",
    "Quarterly Revenue & Telemetry Metrics: Node cluster throughput reached 99.98% availability with zero packet loss during peak transactional load balancing across shard replicas.",
    "Legal & Regulatory Clause 42.8: All cryptographic hashes and document transformations must preserve lexical integrity, structural metadata, and audit trail immutability."
];

for (let i = 1; i <= TOTAL_PAGES; i++) {
    doc.addPage({ margin: 50 });
    doc.fontSize(16).text(`DOCUMENT PAGE #${i} OF ${TOTAL_PAGES}`, { underline: true });
    doc.moveDown(0.5);
    doc.fontSize(10).text(`Batch ID: BATCH-${Math.ceil(i / 100)} | Timestamp: ${new Date().toISOString()}`);
    doc.moveDown(1);

    // Har page par ~250-300 words daal rahe hain (Total ~6,50,000 words!)
    for (let p = 0; p < 8; p++) {
        const text = sampleParagraphs[(i + p) % sampleParagraphs.length];
        doc.fontSize(11).text(`[Para ${i}.${p + 1}] ${text} ${text}`);
        doc.moveDown(0.5);
    }

    if (i % 500 === 0) {
        console.log(`✅ ${i} / ${TOTAL_PAGES} pages generated...`);
    }
}

doc.end();

stream.on('finish', () => {
    const timeTaken = ((Date.now() - startTime) / 1000).toFixed(2);
    const sizeMB = (fs.statSync(outputPath).size / (1024 * 1024)).toFixed(2);
    console.log(`\n🎉 BOOM! ${TOTAL_PAGES}-Page PDF Ready!`);
    console.log(`📁 File Location: ${outputPath}`);
    console.log(`📦 File Size: ${sizeMB} MB | ⏱️ Time: ${timeTaken}s`);
});