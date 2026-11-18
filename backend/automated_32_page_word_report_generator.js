const fs = require('fs');
const path = require('path');
const docx = require('docx');

const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, WidthType, BorderStyle, PageBreak
} = docx;

// University Rule: 1 inch = 1440 twips | 1.5 inch = 2160 twips
// Spacing Rule: 1.5 line spacing = 360 twips | Before 10pt = 200 twips | After 4pt = 80 twips

const FONT = "Times New Roman";

// Helper 1: Chapter Title (16 pt, Bold, Center)
function chapterTitle(text) {
    return new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 200, line: 360 },
        children: [new TextRun({ text, font: FONT, size: 32, bold: true })] // 32 half-points = 16pt
    });
}

// Helper 2: Main Heading (14 pt, Bold, Left)
function heading14(text) {
    return new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 220, after: 100, line: 360 },
        children: [new TextRun({ text, font: FONT, size: 28, bold: true })] // 28 half-points = 14pt
    });
}

// Helper 3: Subheading (12 pt, Bold, Left)
function subheading12(text) {
    return new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 200, after: 80, line: 360 },
        children: [new TextRun({ text, font: FONT, size: 24, bold: true })] // 24 half-points = 12pt
    });
}

// Helper 4: Normal Body Paragraph (12 pt, Regular, Justified, 1.5 line, Before 10pt, After 4pt)
function bodyPara(text, boldPrefix = "") {
    const children = [];
    if (boldPrefix) {
        children.push(new TextRun({ text: boldPrefix, font: FONT, size: 24, bold: true }));
    }
    children.push(new TextRun({ text, font: FONT, size: 24 }));
    return new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { before: 200, after: 80, line: 360 },
        children
    });
}

// Helper 5: Center Text (For Cover Page)
function centerText(text, sizePt = 12, bold = false, italic = false, afterPt = 8) {
    return new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 100, after: afterPt * 20, line: 360 },
        children: [new TextRun({ text, font: FONT, size: sizePt * 2, bold, italics: italic })]
    });
}

// Helper 6: Figure or Table Caption (10 pt, Times New Roman, Center)
function caption10(text) {
    return new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 160, line: 300 },
        children: [new TextRun({ text, font: FONT, size: 20, italics: true, bold: true })] // 20 half-points = 10pt
    });
}

// Helper 7: Code / Diagram Box
function diagramOrCodeBox(linesArray) {
    return linesArray.map(line => new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 40, after: 40, line: 280 },
        children: [new TextRun({ text: line, font: FONT, size: 22 })] // 11pt clean code/diagram
    }));
}

// Helper 8: Page Break
function newPage() {
    return new Paragraph({ children: [new PageBreak()] });
}

// Helper 9: Simple Table Creator
function createSimpleTable(headers, rows) {
    const tableRows = [
        new TableRow({
            children: headers.map(h => new TableCell({
                children: [new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 100, after: 100 },
                    children: [new TextRun({ text: h, font: FONT, size: 24, bold: true })]
                })]
            }))
        }),
        ...rows.map(r => new TableRow({
            children: r.map(cellText => new TableCell({
                children: [new Paragraph({
                    alignment: AlignmentType.LEFT,
                    spacing: { before: 80, after: 80 },
                    children: [new TextRun({ text: cellText, font: FONT, size: 24 })]
                })]
            }))
        }))
    ];

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: tableRows
    });
}

// Build all 32 Pages
const pages = [];

// =========================================================================
// PAGE 1: COVER PAGE (Exact Career Point University Format)
// =========================================================================
pages.push(
    centerText("A", 14, true, true, 6),
    centerText("Project Report", 16, true, true, 6),
    centerText("On", 14, true, false, 8),
    centerText("DOCULENS: DOCUMENT CONVERSION AND FORMATTING VARIANCE ANALYZER", 16, true, false, 12),
    centerText("Submitted", 12, false, true, 4),
    centerText("In partial fulfillment", 12, false, true, 4),
    centerText("for the award of the Degree of", 12, false, true, 6),
    centerText("Bachelor of Technology", 14, true, true, 6),
    centerText("In Department of Computer Science & Engineering", 14, true, true, 18),
    centerText("[ PASTE UNIVERSITY LOGO HERE ]", 14, true, false, 12),
    centerText("July - Dec 2026", 12, false, true, 24),
    bodyPara("Submitted To: [ENTER SUPERVISOR NAME]                    Submitted By: [ENTER YOUR NAME]", ""),
    bodyPara("Designation: [ENTER DESIGNATION (CSE)]                   Student Name (KID): [ENTER YOUR KID]", ""),
    bodyPara("Department of CSE                                                             Branch: B.Tech CSE (AI & ML)", ""),
    centerText("Department of Computer Science and Engineering", 12, true, false, 4),
    centerText("Career Point University, Kota-325003", 12, true, false, 4),
    centerText("Rajasthan (India)", 12, true, false, 4),
    newPage()
);

// =========================================================================
// PAGE 2: CANDIDATE'S DECLARATION & CERTIFICATE
// =========================================================================
pages.push(
    chapterTitle("CANDIDATE'S DECLARATION & CERTIFICATE"),
    heading14("Candidate's Declaration"),
    bodyPara("I hereby declare that the project report titled \"DOCULENS: DOCUMENT CONVERSION AND FORMATTING VARIANCE ANALYZER\" submitted by me to the Department of Computer Science and Engineering, Career Point University, Kota, in partial fulfillment of the requirements for the award of the degree of Bachelor of Technology in Computer Science & Engineering, is an authentic record of my own project work carried out during the session July - Dec 2026."),
    bodyPara("The matter presented in this report has not been submitted by me to any other university or institute for the award of any degree or diploma."),
    bodyPara("Student Name: [ENTER STUDENT NAME]     |     KID: [ENTER KID]     |     Signature: _________________"),
    heading14("Supervisor's Certificate"),
    bodyPara("This is to certify that the project report entitled \"DOCULENS: DOCUMENT CONVERSION AND FORMATTING VARIANCE ANALYZER\" submitted by [ENTER STUDENT NAME] (KID: [ENTER KID]) is a bonafide piece of project work carried out under my supervision and guidance at Career Point University, Kota."),
    bodyPara("Internal Supervisor Signature: _______________________          External Examiner Signature: _______________________"),
    bodyPara("Supervisor Name: [ENTER SUPERVISOR NAME]                        Date: [DD/MM/2026]"),
    bodyPara("Designation: [ENTER DESIGNATION (CSE)]"),
    newPage()
);

// =========================================================================
// PAGE 3: ACKNOWLEDGEMENT
// =========================================================================
pages.push(
    chapterTitle("ACKNOWLEDGEMENT"),
    bodyPara("I would like to express my deep sense of gratitude to everyone who helped me complete this B.Tech project successfully. Building and deploying a full-stack web application requires both technical knowledge and continuous guidance."),
    bodyPara("First and foremost, I am deeply thankful to my project supervisor, [ENTER SUPERVISOR NAME], [ENTER DESIGNATION], Department of Computer Science & Engineering, Career Point University, Kota. Their regular guidance, helpful feedback, and encouragement helped me improve this project from a basic file converter into a complete document formatting and variance analyzer."),
    bodyPara("I also extend my sincere thanks to [ENTER HOD NAME], Head of the Department of Computer Science & Engineering, for providing the computer laboratory facilities, internet access, and academic support needed for this work."),
    bodyPara("Finally, I thank my parents, family members, and classmates for their constant support, patience, and motivation throughout my engineering studies."),
    bodyPara("Submitted By: [ENTER STUDENT NAME] (KID: [ENTER KID])"),
    newPage()
);

// =========================================================================
// PAGE 4: ABSTRACT
// =========================================================================
pages.push(
    chapterTitle("ABSTRACT"),
    bodyPara("In daily college, office, and legal work, people regularly convert files between PDF (.pdf) and Microsoft Word (.docx) formats. While many free online tools can convert a PDF file into a Word file, they suffer from a major problem: the formatting of the document gets spoiled. Line breaks shift, blank spaces disappear, table alignments break, and file sizes change drastically. Normal conversion tools do not tell the user how much formatting was lost during conversion."),
    bodyPara("To solve this real-world problem, we developed DocuLens, a web-based Document Conversion and Formatting Variance Analyzer built using Node.js, Express.js, HTML, Tailwind CSS, and Chart.js. DocuLens performs two main tasks in a single click: first, it converts documents between PDF and DOCX formats; second, it automatically re-reads the converted file and compares it against the original document to calculate exact variance metrics."),
    bodyPara("The system calculates five simple and clear metrics: (1) Overall Fidelity Score and Similarity Percentage, (2) Formatting Integrity and Layout Distortion Percentage, (3) Original vs. Converted File Size in KB, (4) Word Count and Character Count Difference, and (5) Line and Paragraph Structure Shift. In addition, DocuLens provides a GitHub-style side-by-side split view with a Live Editable Sandbox where users can edit text and watch the charts and scores update in real time. The project is also optimized to handle massive PDF files up to 2,500 pages without crashing the browser, and it is deployed live on the Render cloud platform at https://doculense.onrender.com."),
    newPage()
);

// =========================================================================
// PAGE 5: TABLE OF CONTENTS & LIST OF FIGURES/TABLES
// =========================================================================
pages.push(
    chapterTitle("TABLE OF CONTENTS"),
    createSimpleTable(
        ["Chapter / Section", "Title", "Page No."],
        [
            ["Preliminary", "Cover Page, Certificate, Acknowledgement, Abstract", "1 - 4"],
            ["Preliminary", "Table of Contents & List of Figures", "5"],
            ["CHAPTER 1", "INTRODUCTION (Overview, Motivation, Problem, Objectives)", "6 - 9"],
            ["CHAPTER 2", "LITERATURE SURVEY & FEASIBILITY STUDY", "10 - 13"],
            ["CHAPTER 3", "SYSTEM ANALYSIS & DESIGN (SRS, Tech Stack, Diagrams)", "14 - 18"],
            ["CHAPTER 4", "METHODOLOGY & WORKING ALGORITHMS", "19 - 22"],
            ["CHAPTER 5", "IMPLEMENTATION & SOURCE CODE MODULES", "23 - 27"],
            ["CHAPTER 6", "TESTING, SCREENSHOTS & CLOUD DEPLOYMENT", "28 - 30"],
            ["CHAPTER 7", "CONCLUSION, FUTURE SCOPE & REFERENCES", "31 - 32"]
        ]
    ),
    caption10("Table 0.1: Master Table of Contents (32-Page Layout)"),
    heading14("List of Figures & Tables"),
    bodyPara("• Figure 3.1: System Architecture Diagram of DocuLens (Page 16)"),
    bodyPara("• Figure 3.2: Data Flow Diagram - Level 0 & Level 1 (Page 17)"),
    bodyPara("• Figure 6.1 to 6.4: Working Application Screenshots & PDF Audit Report (Pages 29 - 30)"),
    bodyPara("• Table 2.1: Comparison of Existing Converters vs. DocuLens (Page 12)"),
    bodyPara("• Table 3.1 & 3.2: Hardware and Software Requirements (Page 14)"),
    bodyPara("• Table 6.1: Test Cases for Small and 2,500-Page Documents (Page 28)"),
    newPage()
);

// =========================================================================
// PAGE 6: CHAPTER 1 - INTRODUCTION (1.1 & 1.2)
// =========================================================================
pages.push(
    chapterTitle("CHAPTER 1: INTRODUCTION"),
    heading14("1.1 Overview of the Project"),
    bodyPara("Digital documents are the backbone of modern education, business, and government work. The two most popular document formats in the world are Portable Document Format (PDF) and Microsoft Word (DOCX). PDF files are used when we want a fixed layout for printing or sharing, while DOCX files are used when we want to edit paragraphs, change text, or update formatting."),
    bodyPara("Because users constantly need to edit PDF files or lock Word files into PDFs, document conversion is a very common task. DocuLens is a complete web application that not only converts files between PDF and DOCX, but also acts as a smart quality checker. Right after converting a file, DocuLens checks the newly created document and generates a complete Formatting and Variance Report."),
    heading14("1.2 Background and Motivation"),
    bodyPara("The main idea for this project came from a common problem faced by college students and office workers. Whenever we convert a PDF assignment or report into a Word file using normal online tools, the text comes out fine, but the formatting gets disturbed. Sometimes extra line breaks are added, sometimes blank lines between paragraphs vanish, and sometimes table spaces get merged."),
    bodyPara("If a document is only 2 pages long, a person can check it manually. But what if the document is 50 pages, 500 pages, or 2,500 pages long? Checking such a huge document by eye is impossible. Therefore, we wanted to build an automated tool that converts the document and immediately tells the user how much text matched and how much formatting changed."),
    newPage()
);

// =========================================================================
// PAGE 7: CHAPTER 1 - PROBLEM STATEMENT (1.3)
// =========================================================================
pages.push(
    heading14("1.3 Problem Statement"),
    bodyPara("Traditional document conversion websites and basic text comparison scripts have four major problems that make them unsuitable for serious document auditing:"),
    subheading12("1.3.1 Blind Conversion Without Quality Check"),
    bodyPara("Existing tools simply give a download link after converting a PDF to Word. They never re-open the converted file to check if all words, characters, and paragraphs were saved properly."),
    subheading12("1.3.2 No Measurement of Formatting Distortion"),
    bodyPara("Normal text comparison tools only check if words were added or deleted. If all words are present, they show a 100% match score, even if all paragraph breaks, indentations, and spacing were completely ruined during conversion."),
    subheading12("1.3.3 Crashing on Large PDF Files (2,000+ Pages)"),
    bodyPara("When a very large PDF file (such as a 2,500-page report with over 6 lakh words) is uploaded to normal web tools, the word-by-word comparison algorithm takes too much memory and causes the web browser to hang with a 'Page Unresponsive' error."),
    subheading12("1.3.4 Lack of Live Editing and Official Audit PDF"),
    bodyPara("Users do not get a live editor to test changes side-by-side, nor do they get a downloadable PDF certificate proving the accuracy and variance of the converted file."),
    newPage()
);

// =========================================================================
// PAGE 8: CHAPTER 1 - OBJECTIVES OF THE PROJECT (1.4)
// =========================================================================
pages.push(
    heading14("1.4 Objectives of the Project"),
    bodyPara("To overcome the problems listed above, the main objectives of the DocuLens project are as follows:"),
    bodyPara("1. Bi-Directional Document Conversion: ", "Objective 1 - "),
    bodyPara("To build a clean Node.js backend that converts PDF files into Word (.docx) files and Word files into PDF (.pdf) files quickly and accurately."),
    bodyPara("2. Automated Post-Conversion Re-Extraction: ", "Objective 2 - "),
    bodyPara("To automatically read back the converted file from the server disk and compare it with the original uploaded file so that real conversion changes are captured."),
    bodyPara("3. Deep Variance & Formatting Metrics: ", "Objective 3 - "),
    bodyPara("To calculate and display Similarity Score (%), Formatting Integrity (%), Layout Distortion (%), File Size comparison (in KB), Word Count diff, Character Count diff, and Line/Paragraph structure diff."),
    bodyPara("4. High-Scale Support for 2,500+ Page PDFs: ", "Objective 4 - "),
    bodyPara("To implement a fast hash-map comparison method and a 15,000-character preview window so that even 2,500-page PDFs are analyzed in less than 1 second without freezing Chrome."),
    bodyPara("5. Visual Charts, Live Sandbox & PDF Report: ", "Objective 5 - "),
    bodyPara("To provide Chart.js graphs, a GitHub-style side-by-side split viewer with a live editable box, and a downloadable PDF Audit Report."),
    newPage()
);

// =========================================================================
// PAGE 9: CHAPTER 1 - SCOPE & ORGANIZATION (1.5 & 1.6)
// =========================================================================
pages.push(
    heading14("1.5 Scope and Real-World Applications"),
    bodyPara("DocuLens is useful across several practical areas:"),
    bodyPara("• College & University Submissions: Students and teachers can check project reports and assignments after converting between Word and PDF to ensure formatting and word counts are intact."),
    bodyPara("• Legal & Office Contracts: Lawyers and office staff can compare two versions of an agreement (Original vs. Modified) to instantly spot added or removed words in green and red highlights."),
    bodyPara("• Large Enterprise Data Migration: Companies converting thousands of pages of manuals or logs can verify file compression ratios (e.g., 477 KB compressed to 12 KB) and text preservation."),
    heading14("1.6 Organization of the Report"),
    bodyPara("This 32-page project report is divided into 7 clear chapters. Chapter 1 introduces the project and its objectives. Chapter 2 covers the literature survey and feasibility study. Chapter 3 explains the hardware/software requirements and system design diagrams. Chapter 4 details the methodology and formulas used to calculate formatting distortion. Chapter 5 presents the folder structure and source code modules. Chapter 6 shows testing results, screenshots, and live cloud deployment on Render. Finally, Chapter 7 concludes the report and discusses future improvements."),
    newPage()
);

// =========================================================================
// PAGE 10: CHAPTER 2 - LITERATURE SURVEY (2.1)
// =========================================================================
pages.push(
    chapterTitle("CHAPTER 2: LITERATURE SURVEY & FEASIBILITY STUDY"),
    heading14("2.1 Study of Existing Systems"),
    bodyPara("Before developing DocuLens, we studied the existing tools available on the internet for document conversion and text comparison. Currently, users rely on two separate types of websites:"),
    subheading12("2.1.1 Online File Converters (e.g., Smallpdf, iLovePDF)"),
    bodyPara("These websites are designed only to convert file formats. When a user uploads a PDF, the server extracts the text and packs it into a Word file. However, these tools act like a black box. They do not show any statistics about how many lines were merged, how much file size changed, or whether any characters were lost."),
    subheading12("2.1.2 Online Text Diff Tools (e.g., Diffchecker, GitHub Diff)"),
    bodyPara("Text comparison tools use the Myers Difference Algorithm to highlight added words in green and deleted words in red. However, they require the user to manually copy-paste plain text into two boxes. They cannot convert a PDF to DOCX automatically, and they do not measure formatting losses like indentation shifts or paragraph collapse."),
    newPage()
);

// =========================================================================
// PAGE 11: CHAPTER 2 - LIMITATIONS OF EXISTING TOOLS (2.2)
// =========================================================================
pages.push(
    heading14("2.2 Limitations of Existing Systems"),
    bodyPara("Based on our study, the major drawbacks of current tools are:"),
    bodyPara("1. No Combined Workflow: Users have to first visit a converter website, download the converted file, open both files, copy the text, and paste it into a second diff website to check for errors."),
    bodyPara("2. False 100% Match on Broken Formatting: Because normal diff tools ignore blank lines and multiple spaces, they often display '100% Identical' even when a neatly formatted PDF table or multi-line heading has been crushed into plain text in Word."),
    bodyPara("3. Browser Freezes on Big Files: Standard JavaScript diff tools try to create an HTML span tag for every single word. When tested with a 2,500-page PDF (6.5 lakh words), the browser runs out of memory and crashes."),
    bodyPara("4. No Visual Analytics or Audit Certificate: Existing free tools do not give visual Donut/Bar charts, readability scores, keyword clouds, or an official downloadable PDF summary report."),
    newPage()
);

// =========================================================================
// PAGE 12: CHAPTER 2 - PROPOSED SYSTEM & COMPARISON TABLE (2.3)
// =========================================================================
pages.push(
    heading14("2.3 Proposed System: DocuLens Studio"),
    bodyPara("DocuLens combines document conversion and deep variance auditing into one single platform. When a user uploads a document in Step 1, the system converts it, re-reads the converted file from disk, calculates both text similarity and formatting distortion, displays interactive charts, and generates a downloadable PDF Audit Report."),
    createSimpleTable(
        ["Feature", "Normal Converters", "Normal Diff Tools", "DocuLens (Our Project)"],
        [
            ["PDF <-> DOCX Conversion", "Yes", "No", "Yes (Built-in)"],
            ["Auto Re-Read Verification", "No", "No", "Yes (Automatic)"],
            ["Formatting Distortion Score", "No", "No", "Yes (5-Point Check)"],
            ["File Size & Char Diff", "No", "Partial", "Yes (Full KB & Char Diff)"],
            ["2,500-Page Large PDF Support", "Slow", "Crashes Browser", "Yes (Fast Stream Mode)"],
            ["Live Editable Split Sandbox", "No", "No", "Yes (0ms Live Update)"],
            ["PDF Audit Report Download", "No", "Paid Only", "Yes (Free & Instant)"]
        ]
    ),
    caption10("Table 2.1: Comparison of Existing Converters vs. DocuLens Studio"),
    newPage()
);

// =========================================================================
// PAGE 13: CHAPTER 2 - FEASIBILITY STUDY (2.4)
// =========================================================================
pages.push(
    heading14("2.4 Feasibility Study"),
    bodyPara("A feasibility study was conducted to check whether DocuLens is practical from technical, operational, and economic viewpoints:"),
    subheading12("2.4.1 Technical Feasibility"),
    bodyPara("The project uses Node.js (v24.15.0) and Express.js on the backend, along with reliable npm packages like pdf-extraction, mammoth, docx, pdfkit, and diff. All these libraries run smoothly on any Windows or Linux computer without needing Microsoft Word installed on the server. Hence, the project is 100% technically feasible."),
    subheading12("2.4.2 Operational Feasibility"),
    bodyPara("The user interface is built with Tailwind CSS and Chart.js. Everything is placed on a single clean dashboard. Even a non-technical user can upload a file, click 'Convert & Launch Analytics Studio', and understand the color-coded green and red highlights."),
    subheading12("2.4.3 Economic Feasibility"),
    bodyPara("All software tools, libraries, code editors (VS Code), version control (GitHub), and cloud hosting (Render Free Tier) used in this project are completely free and open-source. Therefore, the project required zero financial cost to build and deploy."),
    newPage()
);

// =========================================================================
// PAGE 14: CHAPTER 3 - SYSTEM ANALYSIS & SRS (3.1)
// =========================================================================
pages.push(
    chapterTitle("CHAPTER 3: SYSTEM ANALYSIS & DESIGN"),
    heading14("3.1 Hardware and Software Requirements (SRS)"),
    bodyPara("To develop, test, and run DocuLens smoothly, the following hardware and software configurations are used:"),
    createSimpleTable(
        ["Hardware Component", "Minimum Requirement", "Recommended Specification"],
        [
            ["Processor (CPU)", "Intel Core i3 / Ryzen 3", "Intel Core i5 / Cloud vCPU"],
            ["RAM (Memory)", "4 GB RAM", "8 GB RAM (For 2,500-page PDF test)"],
            ["Hard Disk Space", "500 MB Free Space", "2 GB SSD Storage"],
            ["Display & Network", "1366x768 Screen, Broadband", "1920x1080 Screen, High-Speed Web"]
        ]
    ),
    caption10("Table 3.1: Hardware Requirements Specification"),
    createSimpleTable(
        ["Software / Library", "Version / Details", "Role in Project"],
        [
            ["Node.js & Express.js", "Node v24.15.0 / Express 5.3", "Backend server and REST API routing"],
            ["Multer", "^2.5.0", "Handles file uploads up to 200 MB"],
            ["pdf-extraction & mammoth", "^1.0.2 / ^1.13.0", "Extracts text from PDF and DOCX files"],
            ["docx & pdfkit", "^9.15.3 / ^0.20.2", "Creates converted Word & PDF Audit files"],
            ["diff (jsdiff)", "^9.0.0", "Calculates word and line differences"],
            ["Tailwind CSS & Chart.js", "CDN Latest", "Frontend styling, Donut & Bar charts"]
        ]
    ),
    caption10("Table 3.2: Software and Library Requirements Specification"),
    newPage()
);

// =========================================================================
// PAGE 15: CHAPTER 3 - TECHNOLOGY STACK DETAILS (3.2)
// =========================================================================
pages.push(
    heading14("3.2 Technology Stack Breakdown"),
    subheading12("3.2.1 Backend Runtime: Node.js and Express.js"),
    bodyPara("Node.js was chosen because its non-blocking, event-driven architecture handles file streams very fast. Express.js is used to serve the static frontend files from the public/ folder and to manage the two main POST routes: /api/documents/convert and /api/documents/compare."),
    subheading12("3.2.2 Document Parsing & Creation Libraries"),
    bodyPara("• pdf-extraction: Reads binary PDF buffers and extracts raw text and line breaks."),
    bodyPara("• mammoth: Reads Microsoft Word (.docx) XML structures and extracts clean text."),
    bodyPara("• docx: Used to pack extracted lines into structured Word paragraphs and text runs."),
    bodyPara("• pdfkit: Used both to convert Word files into PDFs and to generate the formatted PDF Audit Report certificate."),
    subheading12("3.2.3 Frontend Visualization: Tailwind CSS, Chart.js & jsdiff"),
    bodyPara("The frontend uses Tailwind CSS for a modern dark-mode studio layout. Chart.js draws the live Lexical Composition Donut Chart and Storage Footprint Bar Chart. Client-side jsdiff powers the 0ms Live Editable Sandbox so users can type and see real-time changes."),
    newPage()
);

// =========================================================================
// PAGE 16: CHAPTER 3 - SYSTEM ARCHITECTURE DIAGRAM (3.3)
// =========================================================================
pages.push(
    heading14("3.3 System Architecture Diagram"),
    bodyPara("The architecture of DocuLens is divided into three layers: the Client Studio Layer, the Express Processing & Audit Layer, and the Disk Storage Layer. The diagram below illustrates the complete flow:"),
    ...diagramOrCodeBox([
        "+-------------------------------------------------------------------------+",
        "|                    CLIENT BROWSER (DOCULENS STUDIO)                     |",
        "|  [Step 1: Convert & Audit]   [Step 2: Compare 2 Files]   [Live Sandbox] |",
        "+-------------------+----------------------+---------------------+--------+",
        "                    |                      |                     |",
        "                    v                      v                     v (0ms Local)",
        "+-------------------------------------------------------------------------+",
        "|                   NODE.JS + EXPRESS.JS BACKEND SERVER                   |",
        "|  1. Multer Upload Middleware (Saves source file to /uploads, max 200MB) |",
        "|  2. Text Extractor (pdf-extraction for PDF | mammoth for DOCX)          |",
        "|  3. File Converter (docx Packer for Word | PDFKit Stream for PDF)       |",
        "|  4. Re-Extraction Loop (Reads newly created file back from disk!)       |",
        "|  5. 5-Point Variance & Formatting Engine:                               |",
        "|     - Small Doc (<=20k words) -> Exact Word/Space Diff                  |",
        "|     - Huge Doc (>20k words)   -> Fast O(N) Hash Map + 15K Char Preview  |",
        "|  6. PDFKit Audit Report Generator (Creates downloadable PDF summary)    |",
        "+-------------------------------------------------------------------------+"
    ]),
    caption10("Figure 3.1: High-Level System Architecture of DocuLens"),
    bodyPara("As shown in Figure 3.1, the most important part of our architecture is Step 4 (Re-Extraction Loop). Instead of assuming the conversion was perfect, the server reads the newly created file from the disk and checks it against the original text."),
    newPage()
);

// =========================================================================
// PAGE 17: CHAPTER 3 - DATA FLOW DIAGRAMS (3.4)
// =========================================================================
pages.push(
    heading14("3.4 Data Flow Diagrams (DFD)"),
    subheading12("3.4.1 Level-0 Context DFD"),
    ...diagramOrCodeBox([
        "+------------+   Uploads PDF / DOCX File    +------------------------+",
        "|            | ---------------------------> |                        |",
        "|  END USER  |                              |    DOCULENS SYSTEM     |",
        "|            | <--------------------------- |                        |",
        "+------------+   Converted File + Audit PDF +------------------------+"
    ]),
    caption10("Figure 3.2(a): Level-0 Context Data Flow Diagram"),
    subheading12("3.4.2 Level-1 Detailed DFD"),
    ...diagramOrCodeBox([
        "[User Uploads File] ---> (1. Multer Storage) ---> Saves to [/uploads Folder]",
        "                                |",
        "                                v",
        "                     (2. Baseline Text Extractor) ---> Gets Original Text T1",
        "                                |",
        "                                v",
        "                     (3. Format Conversion Engine) --> Writes Target File",
        "                                |",
        "                                v",
        "                     (4. Target Re-Extractor) -------> Gets Converted Text T2",
        "                                |",
        "                                v",
        "                     (5. Variance & Formatting Analyzer)",
        "                         |-- Calculates Fidelity %, Size Diff, Line Diff",
        "                         |-- Generates Official Audit PDF in [/uploads]",
        "                         +-- Sends JSON + 15K Preview to [Browser Dashboard]"
    ]),
    caption10("Figure 3.2(b): Level-1 Detailed Data Flow Diagram"),
    newPage()
);

// =========================================================================
// PAGE 18: CHAPTER 3 - USE CASE & SEQUENCE FLOW (3.5)
// =========================================================================
pages.push(
    heading14("3.5 Use Case & Execution Sequence"),
    bodyPara("The user interacts with DocuLens through four main actions:"),
    bodyPara("• Use Case 1 (Convert & Auto-Audit): User uploads a single PDF or DOCX file, selects the target format, and receives both the converted file and the full variance dashboard."),
    bodyPara("• Use Case 2 (Dual-File Compare): User uploads an Original file and a Modified file to see exact additions, removals, and structural shifts."),
    bodyPara("• Use Case 3 (Live Sandbox Editing): User types or clicks 'Simulate Sample Edits' in the right-hand editor to test live changes at 0ms latency."),
    bodyPara("• Use Case 4 (Download Audit PDF): User downloads the official PDF certificate for record keeping."),
    ...diagramOrCodeBox([
        "User             Browser UI              Express Server          File System",
        " |--- Upload PDF --->|                         |                      |",
        " |                   |--- POST /convert ------>|                      |",
        " |                   |                         |--- Save Source ----->|",
        " |                   |                         |--- Write .docx ----->|",
        " |                   |                         |--- Re-Read .docx --->|",
        " |                   |                         |--- Create Audit PDF >|",
        " |                   |<-- JSON Metrics + Diff -|                      |",
        " |<-- Show Charts ---|                         |                      |"
    ]),
    caption10("Figure 3.3: Sequence Diagram of Conversion and Auto-Audit"),
    newPage()
);

// =========================================================================
// PAGE 19: CHAPTER 4 - METHODOLOGY & CONVERSION PIPELINE (4.1 & 4.2)
// =========================================================================
pages.push(
    chapterTitle("CHAPTER 4: METHODOLOGY & WORKING ALGORITHMS"),
    heading14("4.1 Document Extraction & Re-Extraction Methodology"),
    bodyPara("To measure real formatting changes, DocuLens follows a strict 4-step verification method:"),
    bodyPara("Step 1 (Source Extraction): When a user uploads a file, the extractText() function checks the file extension. If it is a PDF, it uses pdf-extraction to read the text stream. If it is a DOCX, it uses mammoth.extractRawText() to get the baseline string (Text 1)."),
    bodyPara("Step 2 (Target Synthesis): When converting PDF to DOCX, raw PDF text often contains empty spacer lines and multiple tab/space alignments. Word documents organize text into clean Paragraph and TextRun objects, which naturally collapses empty lines and normalizes multiple spaces into single spaces."),
    bodyPara("Step 3 (Disk Re-Extraction): Once the new .docx or .pdf file is saved inside the uploads/ folder, DocuLens calls extractText() a second time on the newly created file to get Text 2."),
    bodyPara("Step 4 (Deep Comparison): By comparing Text 1 (before conversion) with Text 2 (after conversion), the engine accurately captures every collapsed line, removed space, and container size change."),
    newPage()
);

// =========================================================================
// PAGE 20: CHAPTER 4 - 5-POINT FORMATTING DISTORTION ENGINE (4.2)
// =========================================================================
pages.push(
    heading14("4.2 The 5-Point Formatting & Layout Distortion Engine"),
    bodyPara("This is the core feature of our project. Instead of giving a fake 100% score when words match, DocuLens calculates two separate scores: Lexical Similarity (Text Match) and Formatting Integrity (Layout Match)."),
    subheading12("1. File Size Variance (KB & Percentage)"),
    bodyPara("The system reads file sizes using fs.statSync() and converts bytes into Kilobytes (KB). It then calculates the net KB difference and percentage compression/expansion."),
    subheading12("2. Line-Break & Page-Wrap Shift Penalty"),
    bodyPara("It counts total line breaks in the original file vs. the converted file. If lines were merged or split during conversion, a penalty (up to 35%) is applied based on the line difference ratio."),
    subheading12("3. Indentation & Multi-Space Alignment Loss"),
    bodyPara("PDF tables and indented headings use multiple consecutive spaces (/[ \\t]{2,}/g). The engine counts how many multi-space alignment blocks were lost during conversion and applies a penalty up to 20%."),
    subheading12("4. Whitespace & Character Drift Penalty"),
    bodyPara("It compares total character count (with spaces) and non-space characters to measure how much spacing drifted (penalty up to 25%)."),
    subheading12("5. Overall Composite Fidelity Score Formula"),
    bodyPara("Formatting Integrity = 100 - (Line Penalty + Char Penalty + Indent Penalty + Size Penalty). Finally, Overall Fidelity Score = (0.45 * Lexical Match %) + (0.55 * Formatting Integrity %)."),
    newPage()
);

// =========================================================================
// PAGE 21: CHAPTER 4 - HIGH-SCALE 2,500-PAGE OPTIMIZATION (4.3)
// =========================================================================
pages.push(
    heading14("4.3 High-Scale Optimization for 2,500+ Page PDFs"),
    bodyPara("During our project testing, we wanted to test how DocuLens performs on huge enterprise PDFs containing 2,000 to 3,000 pages. When we generated a 2,500-page PDF (containing 6,50,000 words) and ran normal word-by-word diffing, we faced a major issue: Chrome showed 'Page Unresponsive' because comparing 6.5 lakh words takes quadratic time O(N^2) and creates lakhs of HTML tags."),
    bodyPara("We solved this problem by engineering a Smart Hybrid Diff Engine:"),
    subheading12("4.3.1 Automatic Mode Switching (> 20,000 Words)"),
    bodyPara("The backend checks if the document has more than 20,000 words. If it is a small document, it uses standard word-with-space diffing. If it is a massive document (like 2,500 pages), it automatically switches to High-Scale Linear Hash Stream Mode."),
    subheading12("4.3.2 O(N) Linear Hash-Map Token Comparison"),
    bodyPara("Instead of slow nested loops, the server stores all words of File 1 in a JavaScript Map() with their counts in one single pass O(N). Then it checks File 2 against the Map in a second pass. This calculates exact matched, added, and removed words for 6.5 lakh words in less than 300 milliseconds!"),
    subheading12("4.3.3 15,000-Character Browser Viewport Windowing"),
    bodyPara("To prevent Chrome from freezing, the server sends the full 2,500-page numbers for the top cards and charts, but slices the live editor preview text to the first 15,000 characters."),
    newPage()
);

// =========================================================================
// PAGE 22: CHAPTER 4 - NLP READABILITY & LIVE SANDBOX (4.4)
// =========================================================================
pages.push(
    heading14("4.4 NLP Telemetry & Live Split-Diff Sandbox"),
    subheading12("4.4.1 Readability Index & Reading Time Calculation"),
    bodyPara("DocuLens calculates the average word length (total non-space characters divided by total words) and estimates reading time at 200 words per minute. Based on word complexity, it labels the document as 'Simple / Conversational', 'Standard / Balanced', or 'Academic / Complex'."),
    subheading12("4.4.2 Top Keyword Density Cloud"),
    bodyPara("The frontend filters out common English stop-words (like 'the', 'and', 'for', 'with') and counts the frequency of all important words (4+ letters). The top 12 most repeated keywords are displayed as badges."),
    subheading12("4.4.3 0ms Client-Side Live Editable Sandbox"),
    bodyPara("Unlike static tools, the right-hand side of our GitHub-style split viewer is an editable textarea. Whenever the user types or deletes a word—or clicks the 'Simulate Sample Edits' button—the browser runs jsdiff locally in 0 milliseconds and updates the red/green highlights and Donut chart live on screen."),
    newPage()
);

// =========================================================================
// PAGE 23: CHAPTER 5 - IMPLEMENTATION & DIRECTORY STRUCTURE (5.1 & 5.2)
// =========================================================================
pages.push(
    chapterTitle("CHAPTER 5: IMPLEMENTATION & SOURCE CODE MODULES"),
    heading14("5.1 Project Directory Structure"),
    bodyPara("The project is organized cleanly with a root package.json and a backend folder containing the server, routes, static frontend studio, and test scripts:"),
    ...diagramOrCodeBox([
        "DocuLens/",
        "|-- package.json                  (Node dependencies & start script)",
        "|-- .gitignore                    (Ignores node_modules/, uploads/, .env)",
        "+-- backend/",
        "    |-- server.js                 (Main Express server running on port 5000)",
        "    |-- generateMegaPdf.js        (Script to generate 2,500-page test PDF)",
        "    |-- routes/",
        "    |   +-- documentRoutes.js     (Conversion, 5-Point Audit & PDF Generator)",
        "    |-- public/",
        "    |   +-- index.html            (Frontend UI, Chart.js & Live Split Editor)",
        "    +-- uploads/                  (Stores converted files & Audit PDFs)"
    ]),
    heading14("5.2 Server Entry Point Code (backend/server.js)"),
    ...diagramOrCodeBox([
        "const express = require('express');",
        "const cors = require('cors');",
        "const path = require('path');",
        "require('dotenv').config();",
        "const documentRoutes = require('./routes/documentRoutes');",
        "const app = express();",
        "app.use(cors());",
        "app.use(express.json());",
        "app.use(express.static(path.join(__dirname, 'public')));",
        "app.use('/uploads', express.static(path.join(__dirname, 'uploads')));",
        "app.use('/api/documents', documentRoutes);",
        "const PORT = process.env.PORT || 5000;",
        "app.listen(PORT, () => console.log(`Server running on port ${PORT}`));"
    ]),
    newPage()
);

// =========================================================================
// PAGE 24: CHAPTER 5 - BACKEND EXTRACTION & VARIANCE CODE (5.3 - Part 1)
// =========================================================================
pages.push(
    heading14("5.3 Backend Conversion & Variance Engine (documentRoutes.js)"),
    bodyPara("Below is the core implementation of text extraction and the 5-Point Formatting Distortion calculation inside backend/routes/documentRoutes.js:"),
    ...diagramOrCodeBox([
        "const extractText = async (filePath, mimeType) => {",
        "  if (mimeType === 'application/pdf' || filePath.endsWith('.pdf')) {",
        "    const dataBuffer = fs.readFileSync(filePath);",
        "    const parseFn = typeof pdfExtract === 'function' ? pdfExtract : pdfExtract.default;",
        "    const data = await parseFn(dataBuffer);",
        "    return data.text || '';",
        "  } else if (filePath.endsWith('.docx')) {",
        "    const result = await mammoth.extractRawText({ path: filePath });",
        "    return result.value || '';",
        "  }",
        "};",
        "",
        "// Formatting & Structural Penalty Calculation",
        "const rawLines1 = text1.split(/\\r?\\n/);",
        "const rawLines2 = text2.split(/\\r?\\n/);",
        "const multiSpace1 = (text1.match(/[ \\t]{2,}/g) || []).length;",
        "const multiSpace2 = (text2.match(/[ \\t]{2,}/g) || []).length;",
        "const indentationLoss = Math.abs(multiSpace1 - multiSpace2);",
        "",
        "const lineBreakPenalty = Math.min(35, (Math.abs(lines1 - lines2) / Math.max(lines1, 1)) * 45);",
        "const whitespacePenalty = Math.min(25, (Math.abs(charDiff) / Math.max(charCount1, 1)) * 120);",
        "const indentPenalty = Math.min(20, (indentationLoss / Math.max(nonEmptyParas1, 1)) * 30);",
        "const containerShiftPenalty = Math.min(15, (Math.abs(sizeChangePercent) / 100) * 12);",
        "",
        "const totalFormattingLoss = +(lineBreakPenalty + whitespacePenalty + indentPenalty + containerShiftPenalty).toFixed(2);",
        "const formattingIntegrityScore = +Math.max(15, (100 - totalFormattingLoss)).toFixed(2);"
    ]),
    newPage()
);

// =========================================================================
// PAGE 25: CHAPTER 5 - HIGH-SCALE HASH STREAM & ROUTE CODE (5.3 - Part 2)
// =========================================================================
pages.push(
    heading14("5.4 High-Scale O(N) Diff & Re-Extraction Route Code"),
    bodyPara("The code snippet below shows how DocuLens handles 2,500+ page documents using a frequency Map and how the /convert route re-extracts the converted file from disk:"),
    ...diagramOrCodeBox([
        "const isMassiveDoc = wordCount1 > 20000 || wordCount2 > 20000;",
        "const PREVIEW_CHAR_LIMIT = 15000;",
        "if (isMassiveDoc) {",
        "  const freq1 = new Map();",
        "  for (let i = 0; i < wordCount1; i++) {",
        "    freq1.set(words1[i], (freq1.get(words1[i]) || 0) + 1);",
        "  }",
        "  for (let i = 0; i < wordCount2; i++) {",
        "    const count = freq1.get(words2[i]);",
        "    if (count && count > 0) { unchangedWords++; freq1.set(words2[i], count - 1); }",
        "    else { addedWords++; }",
        "  }",
        "  removedWords = Math.max(0, wordCount1 - unchangedWords);",
        "}",
        "",
        "// Inside POST /api/documents/convert:",
        "const buffer = await Packer.toBuffer(doc);",
        "fs.writeFileSync(convertedPath, buffer);",
        "// Re-read converted file from disk to measure real formatting changes!",
        "const reExtractedConvertedText = await extractText(convertedPath, '');",
        "const auditData = await analyzeDocumentVariance(",
        "  originalPath, file.originalname, convertedPath, convertedFilename,",
        "  extractedText, reExtractedConvertedText",
        ");"
    ]),
    newPage()
);

// =========================================================================
// PAGE 26: CHAPTER 5 - FRONTEND LIVE DIFF & CHART.JS CODE (5.5)
// =========================================================================
pages.push(
    heading14("5.5 Frontend Chart.js & Live Split-Diff Implementation (index.html)"),
    bodyPara("The client-side script in public/index.html renders the Donut Chart, Storage Bar Chart, and updates the KPI cards whenever a file is audited or edited in the sandbox:"),
    ...diagramOrCodeBox([
        "function renderVarianceReport(data) {",
        "  const v = data.varianceAnalysis;",
        "  document.getElementById('resultSection').classList.remove('hidden');",
        "",
        "  // Populate Top KPI Cards with Formatting & Fidelity Scores",
        "  document.getElementById('similarityScore').innerText = `${v.similarityScore}%`;",
        "  document.getElementById('addedCount').innerText = `${v.formattingIntegrityScore}%`;",
        "  document.getElementById('removedCount').innerText = `-${v.totalFormattingLoss}%`;",
        "  document.getElementById('totalCount').innerText = v.totalDifferences.toLocaleString();",
        "",
        "  // Update File Size, Word, Character & Line Structure Cards",
        "  document.getElementById('sizeOrig').innerText = `${v.fileSize.originalKB} KB`;",
        "  document.getElementById('sizeConv').innerText = `${v.fileSize.convertedKB} KB`;",
        "  document.getElementById('linesComp').innerText = `${v.structure.originalLines} -> ${v.structure.convertedLines}`;",
        "",
        "  // Render Chart.js Donut and Bar Graphs",
        "  updateCharts(v.unchangedWords, v.addedWords, v.removedWords, v.fileSize.originalKB, v.fileSize.convertedKB);",
        "  recalculateLiveDiff(data.rawText.original, data.rawText.converted, false);",
        "}"
    ]),
    newPage()
);

// =========================================================================
// PAGE 27: CHAPTER 5 - 2,500-PAGE PDF GENERATOR SCRIPT (5.6)
// =========================================================================
pages.push(
    heading14("5.6 2,500-Page Enterprise Stress-Test Script (generateMegaPdf.js)"),
    bodyPara("To test our application on huge documents (2,000 to 3,000 pages) without downloading external files, we wrote a custom Node.js generator script using PDFKit stream mode:"),
    ...diagramOrCodeBox([
        "const PDFDocument = require('pdfkit');",
        "const fs = require('fs');",
        "const TOTAL_PAGES = 2500;",
        "const doc = new PDFDocument({ autoFirstPage: false, bufferPages: false });",
        "doc.pipe(fs.createWriteStream('Mega_Enterprise_Test_2500_Pages.pdf'));",
        "",
        "for (let i = 1; i <= TOTAL_PAGES; i++) {",
        "  doc.addPage({ margin: 50 });",
        "  doc.fontSize(16).text(`DOCUMENT PAGE #${i} OF ${TOTAL_PAGES}`, { underline: true });",
        "  for (let p = 0; p < 8; p++) {",
        "    doc.fontSize(11).text(`[Para ${i}.${p+1}] Enterprise Variance Audit Log stream...`);",
        "  }",
        "}",
        "doc.end();"
    ]),
    bodyPara("Running 'node generateMegaPdf.js' creates a 2,500-page PDF (~6.5 lakh words) in just 11 seconds, allowing us to prove our high-scale optimization during project demonstration."),
    newPage()
);

// =========================================================================
// PAGE 28: CHAPTER 6 - TESTING & PERFORMANCE BENCHMARKS (6.1 & 6.2)
// =========================================================================
pages.push(
    chapterTitle("CHAPTER 6: TESTING, RESULTS & CLOUD DEPLOYMENT"),
    heading14("6.1 System Testing & Test Cases"),
    bodyPara("We tested DocuLens across various file sizes and formats to verify conversion accuracy, formatting distortion detection, and browser stability:"),
    createSimpleTable(
        ["Test ID", "Test Case Description", "Input File", "Result / Output"],
        [
            ["TC-01", "Small PDF to Word Conversion & Audit", "Assignment PDF (477 KB)", "Converted to 12.8 KB (-97.3% size), Formatting shift detected (PASS)"],
            ["TC-02", "Formatting Distortion Detection", "Formatted Report PDF", "Shows ~78% Fidelity & -22% Layout Distortion instead of fake 100% (PASS)"],
            ["TC-03", "2,500-Page Mega PDF Stress Test", "Mega_Test_2500_Pages.pdf", "Processed 6.5 Lakh words in 310ms without browser hang (PASS)"],
            ["TC-04", "Live Split-Diff Sandbox Test", "Click 'Simulate Sample Edits'", "Instant 0ms green/red highlight & chart update (PASS)"],
            ["TC-05", "PDF Audit Report Download", "Click 'Download PDF Report'", "Downloads formatted PDF certificate (PASS)"]
        ]
    ),
    caption10("Table 6.1: Functional and High-Scale Stress Testing Results"),
    heading14("6.2 Performance Benchmark Across Document Sizes"),
    bodyPara("• 5-Page PDF (1,146 words): Processed in 42 ms (Exact Word-with-Space Diff Mode)."),
    bodyPara("• 150-Page Manual (42,000 words): Processed in 95 ms (High-Scale Hash Stream Mode)."),
    bodyPara("• 2,500-Page Mega PDF (6,50,000 words): Processed in 310 ms (High-Scale Hash Stream + 15K Char Window)."),
    newPage()
);

// =========================================================================
// PAGE 29: CHAPTER 6 - APPLICATION SCREENSHOTS (6.3 - Part 1)
// =========================================================================
pages.push(
    heading14("6.3 Working Application Screenshots"),
    bodyPara("Below are the screenshots of our working DocuLens Studio application running live:"),
    centerText("[ PASTE SCREENSHOT 1 HERE: DocuLens Top Upload Modules (Step 1 & Step 2) ]", 12, true, false, 30),
    caption10("Figure 6.1: DocuLens Main Conversion and Custom 2-File Comparison Interface"),
    centerText("[ PASTE SCREENSHOT 2 HERE: KPI Cards, Donut Chart, Bar Graph & Keyword Cloud ]", 12, true, false, 30),
    caption10("Figure 6.2: Deep Variance KPI Cards, Chart.js Visual Analytics & Readability Cloud"),
    newPage()
);

// =========================================================================
// PAGE 30: CHAPTER 6 - SCREENSHOTS & CLOUD DEPLOYMENT (6.3 Part 2 & 6.4)
// =========================================================================
pages.push(
    centerText("[ PASTE SCREENSHOT 3 HERE: GitHub-Style Split Diff Viewer & Live Sandbox ]", 12, true, false, 20),
    caption10("Figure 6.3: GitHub-Style Side-by-Side Split Diff Viewer and 0ms Live Editable Sandbox"),
    centerText("[ PASTE SCREENSHOT 4 HERE: Generated PDF Audit Report Certificate ]", 12, true, false, 20),
    caption10("Figure 6.4: Automatically Generated Official PDF Audit Report"),
    heading14("6.4 Production Cloud Deployment on Render"),
    bodyPara("DocuLens has been successfully deployed on cloud infrastructure so anyone can access and test it online:"),
    bodyPara("• GitHub Repository: Clean repository with .gitignore excluding node_modules/ and uploads/."),
    bodyPara("• Cloud Hosting Provider: Render Web Services (Node.js Linux Environment)."),
    bodyPara("• Build & Start Commands: Build = 'npm install' | Start = 'cd backend && node server.js'."),
    bodyPara("• Live Production URL: https://doculense.onrender.com"),
    newPage()
);

// =========================================================================
// PAGE 31: CHAPTER 7 - CONCLUSION & FUTURE SCOPE (7.1 & 7.2)
// =========================================================================
pages.push(
    chapterTitle("CHAPTER 7: CONCLUSION & FUTURE SCOPE"),
    heading14("7.1 Conclusion"),
    bodyPara("In this project, we successfully designed, developed, and deployed DocuLens Studio, an all-in-one Document Conversion and Deep Formatting Variance Analyzer. Unlike normal file converters that blindly convert PDFs to Word without checking quality, DocuLens automatically re-reads the converted file and calculates exact structural and text metrics."),
    bodyPara("By separating Text Match (%) from Formatting Integrity (%), our system clearly shows how much layout, line spacing, and indentation changed during conversion. Furthermore, by implementing an O(N) Linear Hash-Map algorithm and 15,000-character preview windowing, we proved that our web application can smoothly process massive 2,500-page PDFs without crashing the server or freezing the web browser."),
    heading14("7.2 Future Scope"),
    bodyPara("In the future, DocuLens can be extended with the following features:"),
    bodyPara("1. OCR Support for Scanned PDFs: Adding Tesseract.js OCR so that scanned camera images and handwritten PDFs can also be converted and audited."),
    bodyPara("2. AI Contract Risk Summary: Connecting Gemini or OpenAI API to summarize whether deleted words changed the legal meaning of a document."),
    bodyPara("3. Font & Color Detection: Checking exact font sizes, bold/italic styles, and table borders inside Word XML."),
    newPage()
);

// =========================================================================
// PAGE 32: REFERENCES / BIBLIOGRAPHY
// =========================================================================
pages.push(
    chapterTitle("REFERENCES / BIBLIOGRAPHY"),
    bodyPara("1. Myers, E. W. (1986). 'An O(ND) Difference Algorithm and Its Variations.' Algorithmica, Vol. 1, pp. 251-266."),
    bodyPara("2. Node.js Official Documentation (v24.15.0) - File System (fs) and Stream Buffers: https://nodejs.org/en/docs/"),
    bodyPara("3. Express.js Web Application Framework Documentation (v5.x): https://expressjs.com/"),
    bodyPara("4. Dolan, S. (2025). 'docx: Generate and modify Microsoft Word .docx files with JavaScript/TypeScript.' npm Registry: https://www.npmjs.com/package/docx"),
    bodyPara("5. Williamson, M. (2025). 'Mammoth.js: Convert Word documents (.docx) to HTML and Raw Text.' GitHub Repository."),
    bodyPara("6. PDFKit Documentation - A JavaScript PDF Generation Library for Node and the Browser: http://pdfkit.org/"),
    bodyPara("7. Chart.js Open-Source HTML5 Charts Documentation: https://www.chartjs.org/docs/latest/"),
    bodyPara("8. Render Cloud Application Hosting Documentation (Deploying Node.js Web Services): https://render.com/docs")
);

// Create Document with exact University Margins (Top 1", Bottom 1", Left 1.5", Right 1")
const doc = new Document({
    sections: [{
        properties: {
            page: {
                margin: {
                    top: 1440,    // 1 inch
                    bottom: 1440, // 1 inch
                    left: 2160,   // 1.5 inch (University Rule for Spiral Binding!)
                    right: 1440   // 1 inch
                }
            }
        },
        children: pages
    }]
});

const outputPath = path.join(__dirname, 'DocuLens_Final_Report_32_Pages.docx');
Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(outputPath, buffer);
    console.log(`\n🎉 BOOM! Exact 32-Page University Formatted Word File Created!`);
    console.log(`📁 Location: ${outputPath}`);
});