import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { execFileSync } from 'child_process';
import { PlaceholderReplacer } from '../src/lib/placeholderReplacer';
import { SofficeConverter } from '../src/lib/sofficeConverter';

const RECORD_COUNT = 60;
const RECORD_BLOCK_HEIGHT = 2;
const RECORD_START_ROW = 5;
const TEMPLATE_PRINT_AREA_END_ROW = 9;

// A4縦・上下マージン0.75インチ → 本文 841.89 - 108 = 733.89pt（デフォルト行15ptで48行/ページ）
const PAGE_CAPACITY_POINTS = 841.89 - 0.75 * 72 * 2;
const DEFAULT_ROW_HEIGHT = 15;

function commandExists(command: string): boolean {
  try {
    execFileSync('which', [command], { stdio: 'pipe' });
    return true;
  } catch {
    return false;
  }
}

async function buildTemplate(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('請求書');

  ws.pageSetup = {
    paperSize: 9,
    orientation: 'portrait',
    printArea: `A1:D${TEMPLATE_PRINT_AREA_END_ROW}`,
    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  };

  ws.mergeCells('A1:D1');
  ws.getCell('A1').value = '請求書 {{会社名}} 御中';
  ws.getCell('A2').value = '発行日: {{発行日}}';
  ws.getCell('A4').value = 'No.';
  ws.getCell('B4').value = '品目';
  ws.getCell('C4').value = '数量';
  ws.getCell('D4').value = '単価';

  ws.getCell('A5').value = '{{##請求.明細.番号}}';
  ws.getCell('B5').value = '{{##請求.明細.項目}}';
  ws.getCell('C6').value = '{{##請求.明細.数量}}';
  ws.getCell('D6').value = '{{##請求.明細.単価}}';

  ws.getCell('C8').value = '合計';
  ws.getCell('D8').value = '{{合計}}';

  return Buffer.from(await wb.xlsx.writeBuffer());
}

async function main() {
  const template = await buildTemplate();
  const replacer = new PlaceholderReplacer();

  const result = await replacer.replacePlaceholders(template, {
    会社名: 'テスト株式会社',
    発行日: new Date('2026-06-10T00:00:00.000Z'),
    合計: 123456,
    請求: {
      明細: Array.from({ length: RECORD_COUNT }, (_, i) => ({
        番号: i + 1,
        項目: `品目${i + 1}`,
        数量: i + 1,
        単価: (i + 1) * 100,
      })),
    },
  });

  // --- 1. 生成されたxlsxのXML検証 ---
  const zip = await JSZip.loadAsync(result);
  const workbookXml = await zip.file('xl/workbook.xml')!.async('string');
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');

  const printArea = workbookXml.match(/name="_xlnm\.Print_Area"[^>]*>([^<]*)<\/definedName>/)?.[1] || '';
  const insertedRows = (RECORD_COUNT - 1) * RECORD_BLOCK_HEIGHT;
  const expectedEndRow = TEMPLATE_PRINT_AREA_END_ROW + insertedRows;

  if (printArea.includes(',')) {
    throw new Error(`print area must be a single range: ${printArea}`);
  }
  if (!printArea.includes(`$A$1:$D$${expectedEndRow}`)) {
    throw new Error(`unexpected print area: ${printArea} (expected $A$1:$D$${expectedEndRow})`);
  }

  const rowBreaksMatch = sheetXml.match(/<rowBreaks count="(\d+)" manualBreakCount="\1">([\s\S]*?)<\/rowBreaks>/);
  if (!rowBreaksMatch) {
    throw new Error('manual rowBreaks are missing');
  }
  const breakIds = Array.from(rowBreaksMatch[2].matchAll(/<brk id="(\d+)" max="16383" man="1"\/>/g))
    .map((m) => Number(m[1]));

  const lastRecordRow = RECORD_START_ROW - 1 + RECORD_COUNT * RECORD_BLOCK_HEIGHT;
  breakIds.forEach((id) => {
    if (id >= RECORD_START_ROW && id <= lastRecordRow && (id - (RECORD_START_ROW - 1)) % RECORD_BLOCK_HEIGHT !== 0) {
      throw new Error(`break splits a record block: id=${id}`);
    }
  });

  const boundaries = [0, ...breakIds, expectedEndRow];
  for (let i = 1; i < boundaries.length; i++) {
    const rowsInPage = boundaries[i] - boundaries[i - 1];
    if (rowsInPage * DEFAULT_ROW_HEIGHT > PAGE_CAPACITY_POINTS + 0.000001) {
      throw new Error(`page ${i} exceeds capacity: rows=${rowsInPage}`);
    }
  }

  for (let row = RECORD_START_ROW; row <= lastRecordRow; row++) {
    if (!new RegExp(`<row[^>]*r="${row}"`).test(sheetXml)) {
      throw new Error(`record row is missing (unexpected blank padding): r=${row}`);
    }
  }

  const expectedPages = breakIds.length + 1;
  console.log(`xlsx OK: printArea=${printArea.replace(/&apos;/g, "'")}, breaks=[${breakIds.join(', ')}], expectedPages=${expectedPages}`);

  const outDir = path.join(os.tmpdir(), 'nodenberg-paging-e2e');
  fs.mkdirSync(outDir, { recursive: true });
  const xlsxPath = path.join(outDir, 'invoice.xlsx');
  fs.writeFileSync(xlsxPath, result);
  console.log(`xlsx written: ${xlsxPath}`);

  // --- 2. LibreOfficeでPDF変換し、改ページ位置が実際に反映されるか検証 ---
  const converter = new SofficeConverter();
  if (!(await converter.checkSofficeInstalled())) {
    console.warn('soffice not installed; skipping PDF verification');
    return;
  }

  const pdfBuffer = await converter.convertExcelToPDF(result);
  const pdfPath = path.join(outDir, 'invoice.pdf');
  fs.writeFileSync(pdfPath, pdfBuffer);
  console.log(`pdf written: ${pdfPath}`);

  if (!commandExists('pdfinfo') || !commandExists('pdftotext')) {
    console.warn('poppler-utils not installed; skipping PDF page inspection');
    return;
  }

  const pdfInfo = execFileSync('pdfinfo', [pdfPath], { encoding: 'utf8' });
  const pageCount = Number(pdfInfo.match(/^Pages:\s+(\d+)$/m)?.[1] || '0');
  if (pageCount !== expectedPages) {
    throw new Error(`PDF page count mismatch: expected ${expectedPages}, got ${pageCount}`);
  }

  // 各ページ先頭のレコードが手動改ページの位置と一致するか確認
  for (let page = 2; page <= pageCount; page++) {
    const pageText = execFileSync(
      'pdftotext',
      ['-f', String(page), '-l', String(page), pdfPath, '-'],
      { encoding: 'utf8' }
    );
    const firstItem = pageText.match(/品目(\d+)/)?.[1];
    const breakRow = breakIds[page - 2];
    const expectedFirstRecord = Math.floor((breakRow + 1 - RECORD_START_ROW) / RECORD_BLOCK_HEIGHT) + 1;
    if (Number(firstItem) !== expectedFirstRecord) {
      throw new Error(
        `page ${page} starts with 品目${firstItem}, expected 品目${expectedFirstRecord} (break after row ${breakRow})`
      );
    }
  }

  console.log(`pdf OK: ${pageCount} pages, page starts match manual row breaks`);
  console.log('verify-paging-e2e: OK');
}

main().catch((err) => {
  console.error('verify-paging-e2e: FAILED');
  console.error(err);
  process.exit(1);
});
