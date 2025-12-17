# PDF to Excel Converter

This repository provides tools to convert PDF files to Excel workbooks while keeping the original layout as closely as possible. Each PDF page becomes its own worksheet and lines are preserved using a monospace font.

## Quick start (browser)
1. Open `convert-pdf-to-excel.html` in your browser (double-click the file or drag it into a tab).
2. Pick a PDF and click **Convert to Excel**.
3. A download will start when the workbook is ready.

_All processing stays in your browser—no uploads or servers required._

### Structured maintenance sheet output
If your PDF matches the maintenance log layout shown below (lines that start with values such as `S49 | 41043142 | AMC | ...`
followed by a second line of details), the converter will automatically map each entry into columns similar to the provided
sample Excel screenshot:

- `S.NO` (auto-generated row number)
- `HAFTA` (week, e.g., `S49`)
- `İŞ EMRİ NO`, `TİP`, `NO`
- `TANIM` (description plus any follow-up code)
- `ÖNCELİK DURUMU` (combined priority/status values)
- `TARİH` (normalized as `dd.MM.[yyyy]` when a year is present in the PDF header)
- `SÜRE` (time values like `0h10` normalized to `0:10`)

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
