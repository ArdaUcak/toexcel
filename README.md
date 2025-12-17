# PDF to Excel Converter

This repository provides tools to convert PDF files to Excel workbooks while keeping the original layout as closely as possible. Each PDF page becomes its own worksheet and lines are preserved using a monospace font.

## Quick start (browser)
1. Open `convert-pdf-to-excel.html` in your browser (double-click the file or drag it into a tab).
2. Pick a PDF and click **Convert to Excel**.
3. A download will start when the workbook is ready.

_All processing stays in your browser—no uploads or servers required._

## Node.js CLI
If you prefer running the converter in Node.js, use the CLI script.

### Prerequisites
- Node.js 18+
- Access to install npm packages (`pdfjs-dist`, `exceljs`, `yargs`).

### Installation
```
npm install
```

### Usage
Run the converter by specifying an input PDF and output Excel file:
```
node convert-pdf-to-excel.js --input ./input.pdf --output ./output.xlsx
```

Shortcut via npm script:
```
npm run convert -- --input ./input.pdf --output ./output.xlsx
```

### How it works
- Parses each PDF page with `pdfjs-dist` to extract text items and their positions.
- Groups text by line, estimating spacing to preserve the visual arrangement.
- Writes each page to a separate worksheet in an Excel file using `exceljs`, applying a monospace font for alignment.
