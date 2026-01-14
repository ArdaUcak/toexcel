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

function parsePeriodYear(lines) {
  const periodLine = lines.find((line) => /Periyod\s*:\s*\d{2}\/\d{2}\/\d{2}/i.test(line));
  if (!periodLine) return undefined;
  const match = periodLine.match(/Periyod\s*:\s*\d{2}\/\d{2}\/(\d{2})/i);
  if (!match) return undefined;
  const twoDigitYear = match[1];
  const year = Number(twoDigitYear);
  const century = year >= 70 ? 1900 : 2000; // crude pivot to keep near future years in 2000s
  return century + year;
}

function normalizeDate(value, fallbackYear) {
  if (!value) return '';
  const trimmed = value.trim();
  const match = trimmed.match(/^(\d{1,2})[./](\d{1,2})(?:[./](\d{2,4}))?/);
  if (!match) return trimmed;
  const day = match[1].padStart(2, '0');
  const month = match[2].padStart(2, '0');
  let year = match[3];
  if (!year && fallbackYear) {
    year = String(fallbackYear);
  } else if (year && year.length === 2) {
    year = String(Number(year) >= 70 ? 1900 + Number(year) : 2000 + Number(year));
  }
  return [day, month, year].filter(Boolean).join('.');
}

function normalizeDuration(value) {
  if (!value) return '';
  const compact = value.replace(/\s+/g, '');
  const match = compact.match(/^(\d+)h(\d{1,2})$/i);
  if (match) {
    const hours = match[1];
    const minutes = match[2].padStart(2, '0');
    return `${hours}:${minutes}`;
  }
  return value.trim();
}

function splitColumns(line, expectedCount, mergeIndex) {
  const parts = line.split('|').map((part) => part.trim());
  if (parts.length > expectedCount && mergeIndex >= 0) {
    const tailCount = expectedCount - mergeIndex - 1;
    const head = parts.slice(0, mergeIndex);
    const tail = parts.slice(parts.length - tailCount);
    const middle = parts.slice(mergeIndex, parts.length - tailCount).join(' | ');
    return [...head, middle, ...tail];
  }
  while (parts.length < expectedCount) {
    parts.push('');
  }
  return parts;
}

function sanitizeCellValue(value) {
  if (!value) return '';
  return value.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '').trim();
}

function parseCustomRows(lines) {
  const yearHint = parsePeriodYear(lines);
  const records = [];

  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    if (!/^S\d+\s*\|/i.test(line)) continue;

    const mainParts = splitColumns(line, 7, 4);

    const detailLine = lines[i + 1] && lines[i + 1].includes('|') ? lines[i + 1] : '';
    const detailParts = splitColumns(detailLine.replace(/^[_\s]+/, ''), 8, 4);

    const [week, workOrder, type, number, description, plannedDate, plannedDuration] = mainParts;
    const [detailDate, priority, statusCode, controller, code, actualDate, actualDuration, downtime] = detailParts;

    const weekValue = (week || '').replace(/^S\s*/i, '').trim();
    const descriptionWithCode = [description, code].filter(Boolean).join(' | ');
    const priorityValue = priority || statusCode || '';
    const status = '';
    const date = actualDate || detailDate || plannedDate || '';
    const duration = normalizeDuration(actualDuration || plannedDuration || downtime);

    records.push({
      week: sanitizeCellValue(weekValue),
      workOrder: sanitizeCellValue(workOrder || ''),
      type: sanitizeCellValue(type || ''),
      number: sanitizeCellValue(number || ''),
      description: sanitizeCellValue(descriptionWithCode || ''),
      priority: sanitizeCellValue(priorityValue),
      date: sanitizeCellValue(date),
      duration: sanitizeCellValue(duration),
      status: sanitizeCellValue(status),
    });

    if (detailLine) {
      i += 1;
    }
  }

  return records;
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
    const customRows = parseCustomRows(lines);
    const worksheet = workbook.addWorksheet(`Page ${index + 1}`);

    if (customRows.length) {
      worksheet.columns = [
        { header: 'S.NO', key: 'sno', width: 8 },
        { header: 'HAFTA', key: 'week', width: 10 },
        { header: 'İŞ EMRİ NO', key: 'workOrder', width: 15 },
        { header: 'TİP', key: 'type', width: 8 },
        { header: 'NO', key: 'number', width: 12 },
        { header: 'TANIM', key: 'description', width: 60 },
        { header: 'ÖNCELİK DURUMU', key: 'priority', width: 18 },
        { header: 'TARİH', key: 'date', width: 14 },
        { header: 'SÜRE', key: 'duration', width: 12 },
        { header: 'DURUM', key: 'status', width: 14 },
      ];

      customRows.forEach((row, rowIndex) => {
        worksheet.addRow({
          sno: rowIndex + 1,
          week: row.week,
          workOrder: row.workOrder,
          type: row.type,
          number: row.number,
          description: row.description,
          priority: row.priority,
          date: row.date,
          duration: row.duration,
          status: row.status,
        });
      });

      worksheet.getRow(1).font = { bold: true };
      worksheet.columns.forEach((column) => {
        column.alignment = { vertical: 'middle', wrapText: true };
      });
    } else {
      worksheet.columns = [{ header: 'Content', key: 'content', width: 120 }];
      lines.forEach((line) => {
        const row = worksheet.addRow({ content: sanitizeCellValue(line || ' ') });
        row.getCell(1).font = { name: 'Consolas', family: 2, size: 11 };
        row.alignment = { vertical: 'top', wrapText: true };
      });
    }
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
