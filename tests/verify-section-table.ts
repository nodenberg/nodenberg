import ExcelJS from 'exceljs';
import JSZip from 'jszip';
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

async function verifySectionPagingBreaks() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('明細書');

  ws.pageSetup = {
    paperSize: 9,
    orientation: 'portrait',
    printArea: 'A1:D10',
    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  };

  ws.getCell('A1').value = 'タイトル {{会社名}}';
  ws.getCell('A5').value = '{{##請求.明細.番号}}';
  ws.getCell('B5').value = '{{##請求.明細.項目}}';
  ws.getCell('A6').value = '{{##請求.明細.数量}}';
  ws.getCell('B6').value = '{{##請求.明細.単価}}';
  ws.getCell('C9').value = 'FOOTER';

  const template = Buffer.from(await wb.xlsx.writeBuffer());
  const replacer = new PlaceholderReplacer();
  const recordCount = 60;

  const result = await replacer.replacePlaceholders(template, {
    会社名: 'テスト株式会社',
    請求: {
      明細: Array.from({ length: recordCount }, (_, i) => ({
        番号: i + 1,
        項目: `品目${i + 1}`,
        数量: i,
        単価: 100 * i,
      })),
    },
  });

  const zip = await JSZip.loadAsync(result);
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')?.async('string');

  if (!workbookXml || !sheetXml) {
    throw new Error('section paging output is invalid');
  }

  // Print_Areaは分割されず単一範囲のまま、挿入行数ぶん拡張される
  const insertedRows = (recordCount - 1) * 2;
  const expectedEndRow = 10 + insertedRows;
  const printArea = getPrintArea(workbookXml);
  if (printArea.includes(',')) {
    throw new Error(`print area must be a single range: ${printArea}`);
  }
  const ranges = parsePrintAreaRanges(printArea);
  if (ranges.length !== 1 || ranges[0].startRow !== 1 || ranges[0].endRow !== expectedEndRow) {
    throw new Error(`section paging print area is unexpected: ${printArea}`);
  }

  // 改ページは手動rowBreaksで表現される
  const rowBreaksMatch = sheetXml.match(/<rowBreaks count="(\d+)" manualBreakCount="(\d+)">([\s\S]*?)<\/rowBreaks>/);
  if (!rowBreaksMatch) {
    throw new Error('manual rowBreaks are missing');
  }

  const breakIds = Array.from(rowBreaksMatch[3].matchAll(/<brk id="(\d+)" max="16383" man="1"\/>/g))
    .map((m) => Number(m[1]));
  if (breakIds.length === 0 || breakIds.length !== Number(rowBreaksMatch[1]) || breakIds.length !== Number(rowBreaksMatch[2])) {
    throw new Error(`rowBreaks count mismatch: ${rowBreaksMatch[0]}`);
  }

  // 改ページ位置はレコードブロック（2行）の境界に揃う（レコード行は5行目から2行刻み）
  const lastRecordRow = 4 + recordCount * 2;
  breakIds.forEach((id) => {
    if (id < 1 || id >= expectedEndRow) {
      throw new Error(`break id out of print area: ${id}`);
    }
    if (id >= 5 && id <= lastRecordRow && (id - 4) % 2 !== 0) {
      throw new Error(`break splits a record block: id=${id}`);
    }
  });

  // 各ページの行高さ合計がページ容量（A4縦 841.89pt - 上下マージン108pt）を超えない
  const pageCapacity = 841.89 - 72 * 1.5;
  const defaultRowHeight = 15;
  const boundaries = [0, ...breakIds, expectedEndRow];
  for (let i = 1; i < boundaries.length; i++) {
    const rowsInPage = boundaries[i] - boundaries[i - 1];
    if (rowsInPage * defaultRowHeight > pageCapacity + 0.000001) {
      throw new Error(`page ${i} exceeds capacity: rows=${rowsInPage}`);
    }
  }

  // 改ページ用の空白行パディングは挿入されない（レコード行はすべて連続して存在する）
  for (let row = 5; row <= lastRecordRow; row++) {
    if (!new RegExp(`<row[^>]*r="${row}"`).test(sheetXml)) {
      throw new Error(`record row is missing (unexpected blank padding): r=${row}`);
    }
  }
}

function countCellValues(ws: ExcelJS.Worksheet, predicate: (text: string) => boolean): number {
  let count = 0;
  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (predicate(String(cell.value ?? ''))) count += 1;
    });
  });
  return count;
}

async function verifyTallSectionBlocks() {
  const recordCount = 30;
  const blockHeight = 5;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('工事台帳');

  ws.pageSetup = {
    paperSize: 9,
    orientation: 'portrait',
    printArea: 'A1:D12',
    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  };

  ws.getCell('A1').value = '◆工事台帳タイトル';
  ws.getCell('A2').value = '{{会社名}}';
  ws.getCell('A4').value = '◆一覧ヘッダー';

  // 1レコード5行のブロック（行5〜9）
  ws.getCell('A5').value = '{{##工事.明細.番号}}';
  ws.getCell('B5').value = '{{##工事.明細.名称}}';
  ws.getCell('B6').value = '{{##工事.明細.担当}}';
  ws.getCell('C7').value = '{{##工事.明細.数量}}';
  ws.getCell('D8').value = '{{##工事.明細.金額}}';
  ws.getCell('B9').value = '{{##工事.明細.備考}}';

  ws.getCell('A11').value = '◆フッター';

  const template = Buffer.from(await wb.xlsx.writeBuffer());
  const replacer = new PlaceholderReplacer();

  const result = await replacer.replacePlaceholders(template, {
    会社名: 'テスト株式会社',
    工事: {
      明細: Array.from({ length: recordCount }, (_, i) => ({
        番号: i + 1,
        名称: `工事${i + 1}`,
        担当: `担当${i + 1}`,
        数量: i + 1,
        金額: (i + 1) * 1000,
        備考: `備考${i + 1}`,
      })),
    },
  });

  const zip = await JSZip.loadAsync(result);
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')?.async('string');
  if (!workbookXml || !sheetXml) {
    throw new Error('tall section output is invalid');
  }

  const insertedRows = (recordCount - 1) * blockHeight;
  const printArea = getPrintArea(workbookXml);
  if (printArea.includes(',') || !printArea.includes(`$A$1:$D$${12 + insertedRows}`)) {
    throw new Error(`tall section print area is unexpected: ${printArea}`);
  }

  // 改ページは5行ブロックの境界に揃う（レコード行は5行目から5行刻み）
  const lastRecordRow = 4 + recordCount * blockHeight;
  const breakIds = Array.from(
    (sheetXml.match(/<rowBreaks[^>]*>([\s\S]*?)<\/rowBreaks>/)?.[1] || '').matchAll(/\bid="(\d+)"/g)
  ).map((m) => Number(m[1]));
  if (breakIds.length === 0) {
    throw new Error('tall section rowBreaks are missing');
  }
  breakIds.forEach((id) => {
    if (id >= 5 && id <= lastRecordRow && (id - 4) % blockHeight !== 0) {
      throw new Error(`break splits a tall record block: id=${id}`);
    }
  });

  // 読み戻して、ヘッダー/フッター/タイトルが複製されていないこと、各レコードが1回ずつ存在することを確認
  const readBack = new ExcelJS.Workbook();
  await readBack.xlsx.load(result as unknown as ArrayBuffer);
  const sheet = readBack.getWorksheet('工事台帳');
  if (!sheet) throw new Error('tall section sheet is missing');

  [['◆工事台帳タイトル', 1], ['◆一覧ヘッダー', 1], ['◆フッター', 1]].forEach(([text, expected]) => {
    const count = countCellValues(sheet, (v) => v === text);
    if (count !== expected) {
      throw new Error(`"${text}" should appear ${expected} time(s), got ${count}`);
    }
  });

  for (let i = 1; i <= recordCount; i++) {
    const count = countCellValues(sheet, (v) => v === `工事${i}`);
    if (count !== 1) {
      throw new Error(`record 工事${i} should appear once, got ${count}`);
    }
  }
}

async function verifySingleRowSectionBlocks() {
  const recordCount = 100;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('名簿');

  ws.pageSetup = {
    paperSize: 9,
    orientation: 'portrait',
    printArea: 'A1:C8',
    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  };

  ws.getCell('A1').value = '◆名簿タイトル';
  ws.getCell('A4').value = '◆No.ヘッダー';
  ws.getCell('B4').value = '◆氏名ヘッダー';

  // 1レコード1行のブロック（行5）
  ws.getCell('A5').value = '{{##名簿.行.番号}}';
  ws.getCell('B5').value = '{{##名簿.行.氏名}}';

  ws.getCell('A7').value = '◆フッター';

  const template = Buffer.from(await wb.xlsx.writeBuffer());
  const replacer = new PlaceholderReplacer();

  const result = await replacer.replacePlaceholders(template, {
    名簿: {
      行: Array.from({ length: recordCount }, (_, i) => ({
        番号: i + 1,
        氏名: `会員${i + 1}`,
      })),
    },
  });

  const zip = await JSZip.loadAsync(result);
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')?.async('string');
  if (!workbookXml || !sheetXml) {
    throw new Error('single-row section output is invalid');
  }

  const insertedRows = recordCount - 1;
  const printArea = getPrintArea(workbookXml);
  if (printArea.includes(',') || !printArea.includes(`$A$1:$C$${8 + insertedRows}`)) {
    throw new Error(`single-row section print area is unexpected: ${printArea}`);
  }

  // レコード行（5〜104）がすべて存在し、改ページ用の空白パディングがないこと
  const lastRecordRow = 4 + recordCount;
  for (let row = 5; row <= lastRecordRow; row++) {
    if (!new RegExp(`<row[^>]*r="${row}"`).test(sheetXml)) {
      throw new Error(`single-row record row is missing: r=${row}`);
    }
  }

  const breakIds = Array.from(
    (sheetXml.match(/<rowBreaks[^>]*>([\s\S]*?)<\/rowBreaks>/)?.[1] || '').matchAll(/\bid="(\d+)"/g)
  ).map((m) => Number(m[1]));
  if (breakIds.length === 0) {
    throw new Error('single-row section rowBreaks are missing');
  }

  // 読み戻して、ヘッダー/タイトル/フッターが複製されていないことを確認
  const readBack = new ExcelJS.Workbook();
  await readBack.xlsx.load(result as unknown as ArrayBuffer);
  const sheet = readBack.getWorksheet('名簿');
  if (!sheet) throw new Error('single-row section sheet is missing');

  ['◆名簿タイトル', '◆No.ヘッダー', '◆氏名ヘッダー', '◆フッター'].forEach((text) => {
    const count = countCellValues(sheet, (v) => v === text);
    if (count !== 1) {
      throw new Error(`"${text}" should appear once, got ${count} (header duplicated?)`);
    }
  });

  for (let i = 1; i <= recordCount; i++) {
    const count = countCellValues(sheet, (v) => v === `会員${i}`);
    if (count !== 1) {
      throw new Error(`record 会員${i} should appear once, got ${count}`);
    }
  }
}

async function main() {
  await verifySectionTableExpansion();
  await verifyLegacyTableExpansion();
  await verifySectionPagingBreaks();
  await verifyTallSectionBlocks();
  await verifySingleRowSectionBlocks();
  console.log('verify-section-table: OK');
}

main().catch((err) => {
  console.error('verify-section-table: FAILED');
  console.error(err);
  process.exit(1);
});
