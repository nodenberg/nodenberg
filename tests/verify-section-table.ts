import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import * as fs from 'fs';
import * as path from 'path';
import { PlaceholderReplacer } from '../src/lib/placeholderReplacer';

function getPrintArea(workbookXml: string): string {
  const match = workbookXml.match(/name="_xlnm\.Print_Area"[^>]*>([^<]*)<\/definedName>/);
  if (!match) {
    throw new Error('print area is missing');
  }
  return match[1];
}

function parseMaxPrintAreaRow(printArea: string): number {
  const rowNumbers = Array.from(printArea.matchAll(/\$[A-Z]+\$(\d+)/g)).map((match) => Number(match[1]));
  return rowNumbers.length > 0 ? Math.max(...rowNumbers) : 0;
}

function parsePrintAreaRanges(printArea: string): Array<{ startRow: number; endRow: number }> {
  return printArea
    .split(',')
    .map((range) => {
      const match = range.match(/\$[A-Z]+\$(\d+):\$[A-Z]+\$(\d+)/);
      if (!match) {
        throw new Error(`invalid print area range: ${range}`);
      }
      return {
        startRow: Number(match[1]),
        endRow: Number(match[2]),
      };
    });
}

async function verifySectionTableExpansion() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('請求書');

  ws.pageSetup = {
    paperSize: 9,
    orientation: 'portrait',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    printArea: 'A1:D10',
  };

  ws.getCell('A1').value = '{{会社名}}';
  ws.getCell('A5').value = {
    richText: [
      {
        font: { name: 'Meiryo', bold: true, size: 12 },
        text: '{{##請求.明細.番号}}',
      },
    ],
  };
  ws.getCell('B5').value = {
    richText: [
      {
        font: { name: 'Meiryo', italic: true, size: 11 },
        text: '{{##請求.明細.項目}}',
      },
    ],
  };
  ws.getCell('A6').value = '{{##請求.明細.数量}}';
  ws.getCell('B6').value = '{{##請求.明細.単価}}';
  ws.getCell('C6').value = { formula: 'A6*B6' };
  ws.getCell('D6').value = '{{##請求.明細.日付}}';
  ws.getCell('C9').value = 'FOOTER';

  const template = Buffer.from(await wb.xlsx.writeBuffer());
  const replacer = new PlaceholderReplacer();

  const result = await replacer.replacePlaceholders(template, {
    会社名: 'テスト株式会社',
    請求: {
      明細: [
        { 番号: 1, 項目: 'A', 数量: 2, 単価: 100 },
        { 番号: 2, 項目: 'B', 数量: 3, 単価: 200, 日付: new Date('2026-01-31T00:00:00.000Z') },
        { 番号: 3, 項目: 'C', 数量: 4, 単価: 300 },
      ],
    },
  });

  const zip = await JSZip.loadAsync(result);
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')?.async('string');
  const sharedStringsXml = await zip.file('xl/sharedStrings.xml')?.async('string');

  if (!workbookXml || !sheetXml || !sharedStringsXml) {
    throw new Error('generated xlsx structure is invalid');
  }

  if (!sharedStringsXml.includes('テスト株式会社')) {
    throw new Error('top-level placeholder replacement failed');
  }

  if (!sharedStringsXml.includes('A') || !sharedStringsXml.includes('B') || !sharedStringsXml.includes('C')) {
    throw new Error('section-table data replacement failed');
  }

  if (!sharedStringsXml.includes('2026/01/31')) {
    throw new Error('section-table date formatting failed');
  }

  if (!/<si><r><rPr>[\s\S]*?<t>A<\/t><\/r><\/si>/.test(sharedStringsXml)) {
    throw new Error('section-table first-run style preservation failed');
  }

  const printArea = getPrintArea(workbookXml);
  if (parseMaxPrintAreaRow(printArea) < 10) {
    throw new Error(`print area does not cover expanded rows: ${printArea}`);
  }

  if (!/<c r="C8"[^>]*><f[^>]*>A8\*B8<\/f>/.test(sheetXml)) {
    throw new Error('formula row-shift failed for duplicated block row C8');
  }

  if (!/<c r="C10"[^>]*><f[^>]*>A10\*B10<\/f>/.test(sheetXml)) {
    throw new Error('formula row-shift failed for duplicated block row C10');
  }

  [5, 6, 7, 8, 9, 10].forEach((row) => {
    if (!new RegExp(`<row[^>]*r="${row}"`).test(sheetXml)) {
      throw new Error(`expected expanded row is missing: r=${row}`);
    }
  });
}

async function verifyLegacyTableExpansion() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Legacy');

  ws.pageSetup = {
    paperSize: 9,
    orientation: 'portrait',
    printArea: 'A1:C5',
  };

  ws.getCell('A1').value = '{{会社名}}';
  ws.getCell('A3').value = '{{#明細.番号}}';
  ws.getCell('B3').value = '{{#明細.項目}}';
  ws.getCell('C3').value = '{{#明細.数量}}';

  const template = Buffer.from(await wb.xlsx.writeBuffer());
  const replacer = new PlaceholderReplacer();

  const result = await replacer.replacePlaceholders(template, {
    会社名: 'テスト株式会社',
    明細: [
      { 番号: 1, 項目: 'A', 数量: 2 },
      { 番号: 2, 項目: 'B', 数量: 3 },
      { 番号: 3, 項目: 'C', 数量: 4 },
    ],
  });

  const zip = await JSZip.loadAsync(result);
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')?.async('string');
  const sharedStringsXml = await zip.file('xl/sharedStrings.xml')?.async('string');

  if (!workbookXml || !sheetXml || !sharedStringsXml) {
    throw new Error('generated legacy xlsx structure is invalid');
  }

  if (!sharedStringsXml.includes('テスト株式会社')) {
    throw new Error('legacy top-level placeholder replacement failed');
  }

  if (!sharedStringsXml.includes('A') || !sharedStringsXml.includes('B') || !sharedStringsXml.includes('C')) {
    throw new Error('legacy table data replacement failed');
  }

  [3, 4, 5].forEach((row) => {
    if (!new RegExp(`<row[^>]*r="${row}"`).test(sheetXml)) {
      throw new Error(`legacy expanded row is missing: r=${row}`);
    }
  });

  const printArea = getPrintArea(workbookXml);
  if (parseMaxPrintAreaRow(printArea) < 5) {
    throw new Error(`legacy print area was not updated: ${printArea}`);
  }
}

async function verifySectionPagingWithFixture() {
  const projectRoot = path.resolve(process.cwd(), '..');
  const templatePath = path.join(projectRoot, '06_test-assets', 'template_section_02_with-tall-cell.xlsx');
  const requestPath = path.join(projectRoot, '06_test-assets', 'request_section_with-tall-cell.json');

  const template = fs.readFileSync(templatePath);
  const data = JSON.parse(fs.readFileSync(requestPath, 'utf8'));
  const replacer = new PlaceholderReplacer();
  const result = await replacer.replacePlaceholders(template, data);

  const zip = await JSZip.loadAsync(result);
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')?.async('string');

  if (!workbookXml || !sheetXml) {
    throw new Error('section paging fixture output is invalid');
  }

  const printArea = getPrintArea(workbookXml);
  const ranges = parsePrintAreaRanges(printArea);
  if (ranges.length !== 2 || ranges[0].startRow !== 1 || ranges[0].endRow !== 37 || ranges[1].startRow !== 38 || ranges[1].endRow !== 69) {
    throw new Error(`section paging print area is unexpected: ${printArea}`);
  }

  if (!/<row[^>]*r="33"/.test(sheetXml) || !/<row[^>]*r="38"/.test(sheetXml) || !/<row[^>]*r="69"/.test(sheetXml)) {
    throw new Error('section paging boundary rows are missing');
  }

  [34, 35, 36, 37].forEach((row) => {
    if (new RegExp(`<row[^>]*r="${row}"`).test(sheetXml)) {
      throw new Error(`section paging gap row should be blank: r=${row}`);
    }
  });
}

async function main() {
  await verifySectionTableExpansion();
  await verifyLegacyTableExpansion();
  await verifySectionPagingWithFixture();
  console.log('verify-section-table: OK');
}

main().catch((err) => {
  console.error('verify-section-table: FAILED');
  console.error(err);
  process.exit(1);
});
