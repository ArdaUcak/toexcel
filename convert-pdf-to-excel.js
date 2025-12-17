#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const { hideBin } = require('yargs/helpers');
const yargs = require('yargs/yargs');
const ExcelJS = require('exceljs');

let pdfjsLibPromise;

async function getPdfJs() {
  if (!pdfjsLibPromise) {
    pdfjsLibPromise = import('pdfjs-dist/legacy/build/pdf.mjs');
  }

  const pdfjsLib = await pdfjsLibPromise;
  return pdfjsLib;
}

const LINE_HEIGHT_THRESHOLD = 3;
const SPACE_SCALING = 1.5;

function parseArguments() {
  const argv = yargs(hideBin(process.argv))
    .usage('Usage: $0 --input <input.pdf> --output <output.xlsx>')
    .option('input', {
      alias: 'i',
      describe: 'Path to the PDF file to convert',
      type: 'string',
      demandOption: true,
    })
    .option('output', {
      alias: 'o',
      describe: 'Path for the generated Excel file',
      type: 'string',
      demandOption: true,
    })
    .example('$0 -i report.pdf -o report.xlsx', 'Convert report.pdf into an Excel file')
    .help()
    .argv;

  return {
    input: path.resolve(argv.input),
    output: path.resolve(argv.output),
  };
}

function groupItemsIntoLines(textItems) {
  const sorted = textItems
    .map((item) => ({
      text: item.str,
      x: item.transform[4],
      y: item.transform[5],
      width: item.width,
    }))
    .sort((a, b) => b.y - a.y || a.x - b.x);

  const lines = [];
  sorted.forEach((item) => {
    const existingLine = lines.find((line) => Math.abs(line.y - item.y) < LINE_HEIGHT_THRESHOLD);
    if (existingLine) {
      existingLine.items.push(item);
      existingLine.y = (existingLine.y * existingLine.items.length + item.y) / (existingLine.items.length + 1);
    } else {
      lines.push({ y: item.y, items: [item] });
    }
  });

  return lines
    .sort((a, b) => b.y - a.y)
    .map((line) => {
      const ordered = line.items.sort((a, b) => a.x - b.x);
      let assembled = '';
      for (let i = 0; i < ordered.length; i += 1) {
        const current = ordered[i];
        if (i === 0) {
          assembled += current.text;
          continue;
        }
        const prev = ordered[i - 1];
        const gap = current.x - (prev.x + prev.width);
        const spaces = gap > 0 ? Math.round((gap / prev.width) * SPACE_SCALING) : 0;
        assembled += ' '.repeat(Math.max(0, spaces)) + current.text;
      }
      return assembled.trimEnd();
    });
}

async function extractPages(pdfPath) {
  const pdfjsLib = await getPdfJs();
  const data = new Uint8Array(fs.readFileSync(pdfPath));
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const pages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    const lines = groupItemsIntoLines(content.items);
    pages.push(lines);
  }

  return pages;
}

async function writeWorkbook(pages, outputPath) {
  const workbook = new ExcelJS.Workbook();

  pages.forEach((lines, index) => {
    const worksheet = workbook.addWorksheet(`Page ${index + 1}`);
    worksheet.columns = [{ header: 'Content', key: 'content', width: 120 }];
    lines.forEach((line) => {
      const row = worksheet.addRow({ content: line || ' ' });
      row.getCell(1).font = { name: 'Consolas', family: 2, size: 11 };
      row.alignment = { vertical: 'top', wrapText: true };
    });
  });

  await workbook.xlsx.writeFile(outputPath);
}

async function main() {
  const { input, output } = parseArguments();

  if (!fs.existsSync(input)) {
    throw new Error(`Input file does not exist: ${input}`);
  }

  const pages = await extractPages(input);
  await writeWorkbook(pages, output);
  // eslint-disable-next-line no-console
  console.log(`Converted ${input} -> ${output}`);
}

main().catch((error) => {
  // eslint-disable-next-line no-console
  console.error(error.message || error);
  process.exitCode = 1;
});
