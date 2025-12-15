# PDF to Excel Converter

This repository provides a Node.js script that converts PDF files to Excel workbooks while keeping the original layout as closely as possible. Each PDF page becomes its own worksheet and lines are preserved using a monospace font.

## Prerequisites
- Node.js 18+
- Access to install npm packages (`pdfjs-dist`, `exceljs`, `yargs`).

## Installation
```
npm install
```

## Usage
Run the converter by specifying an input PDF and output Excel file:
```
node convert-pdf-to-excel.js --input ./input.pdf --output ./output.xlsx
```

Shortcut via npm script:
```
npm run convert -- --input ./input.pdf --output ./output.xlsx
```

## How it works
- Parses each PDF page with `pdfjs-dist` to extract text items and their positions.
- Groups text by line, estimating spacing to preserve the visual arrangement.
- Writes each page to a separate worksheet in an Excel file using `exceljs`, applying a monospace font for alignment.
