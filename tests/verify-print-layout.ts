import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import * as fs from 'fs';
import * as path from 'path';
import { PlaceholderReplacer } from '../src/lib/placeholderReplacer';
import { PDFGenerator } from '../src/lib/pdfGenerator';

function parseAttrs(tag: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  for (const match of tag.matchAll(/([A-Za-z_:][\w:.-]*)="([^"]*)"/g)) {
    attrs[match[1]] = match[2];
  }
  return attrs;
}

function getTagAttrs(xml: string, name: string): Record<string, string> {
  const tag = xml.match(new RegExp(`<${name}\\b([^>]*)/>`));
  if (!tag) {
    throw new Error(`${name} is missing`);
  }
  return parseAttrs(tag[1]);
}

async function buildTemplate(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('請求書');

  ws.pageSetup = {
    paperSize: 9,
    orientation: 'portrait',
    fitToWidth: 1,
    fitToHeight: 1,
    printArea: 'A1:D12',
    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  };

  ws.getCell('A1').value = '請求書 {{会社名}}';
  ws.getCell('A5').value = '{{##請求.明細.番号}}';
  ws.getCell('B5').value = '{{##請求.明細.項目}}';
  ws.getCell('A6').value = '{{##請求.明細.数量}}';
  ws.getCell('B6').value = '{{##請求.明細.単価}}';
  ws.getCell('A10').value = 'END';

  return Buffer.from(await wb.xlsx.writeBuffer());
}

function buildData(recordCount = 3) {
  return {
    会社名: 'テスト株式会社',
    請求: {
      明細: Array.from({ length: recordCount }, (_, i) => ({
        番号: i + 1,
        項目: `品目${i + 1}`,
        数量: i + 1,
        単価: (i + 1) * 100,
      })),
    },
  };
}

async function getSheetXml(buffer: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(buffer);
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')?.async('string');
  if (!sheetXml) {
    throw new Error('sheet1.xml is missing');
  }
  return sheetXml;
}

async function verifyMarginPresetNarrow() {
  const template = await buildTemplate();
  const result = await new PlaceholderReplacer().replacePlaceholders(template, buildData(), {
    printLayout: { marginPreset: 'narrow' },
  });

  const sheetXml = await getSheetXml(result);
  const margins = getTagAttrs(sheetXml, 'pageMargins');

  if (margins.left !== '0.25' || margins.right !== '0.25') {
    throw new Error(`narrow margins were not applied: ${JSON.stringify(margins)}`);
  }
  if (margins.top !== '0.75' || margins.bottom !== '0.75' || margins.header !== '0.3' || margins.footer !== '0.3') {
    throw new Error(`narrow margin defaults are unexpected: ${JSON.stringify(margins)}`);
  }
}

async function verifyMarginPresetWithPartialOverride() {
  const template = await buildTemplate();
  const result = await new PlaceholderReplacer().replacePlaceholders(template, buildData(60), {
    printLayout: {
      marginPreset: 'narrow',
      margins: { top: 0, bottom: 0, header: 0, footer: 0 },
      fit: { width: 1, height: 0 },
      paperSize: 'A4',
      orientation: 'portrait',
      recalculatePagination: true,
    },
  });

  const sheetXml = await getSheetXml(result);
  const margins = getTagAttrs(sheetXml, 'pageMargins');
  const setup = getTagAttrs(sheetXml, 'pageSetup');

  if (margins.left !== '0.25' || margins.right !== '0.25') {
    throw new Error(`partial override did not preserve narrow side margins: ${JSON.stringify(margins)}`);
  }
  if (margins.top !== '0' || margins.bottom !== '0' || margins.header !== '0' || margins.footer !== '0') {
    throw new Error(`partial override vertical margins were not applied: ${JSON.stringify(margins)}`);
  }
  if (setup.paperSize !== '9' || setup.orientation !== 'portrait' || setup.fitToWidth !== '1' || setup.fitToHeight !== '0') {
    throw new Error(`pageSetup was not applied: ${JSON.stringify(setup)}`);
  }
  if (!/<pageSetUpPr\b[^>]*fitToPage="1"/.test(sheetXml)) {
    throw new Error('pageSetUpPr fitToPage="1" was not applied for fit settings');
  }
}

async function expectInvalidPrintLayout(label: string, printLayout: unknown) {
  const template = await buildTemplate();
  try {
    await new PlaceholderReplacer().replacePlaceholders(template, buildData(), {
      printLayout: printLayout as any,
    });
  } catch {
    return;
  }
  throw new Error(`${label} should have failed`);
}

async function verifyInvalidPrintLayout() {
  await expectInvalidPrintLayout('empty margins', { margins: {} });
  await expectInvalidPrintLayout('unknown margin key', { margins: { top: 0, gutter: 0.2 } });
  await expectInvalidPrintLayout('negative margin', { margins: { top: -0.1 } });
  await expectInvalidPrintLayout('non-number margin', { margins: { top: '0.5' } });
  await expectInvalidPrintLayout('non-integer fit', { fit: { width: 1.5 } });
  await expectInvalidPrintLayout('unknown paperSize', { paperSize: 'B4' });
  await expectInvalidPrintLayout('unknown orientation', { orientation: 'sideways' });
  await expectInvalidPrintLayout('non-boolean recalculatePagination', { recalculatePagination: 'true' });
  await expectInvalidPrintLayout('false recalculatePagination', { recalculatePagination: false });
  await expectInvalidPrintLayout('null marginPreset', { marginPreset: null });
  await expectInvalidPrintLayout('null paperSize', { paperSize: null });
}

async function verifyPdfGenerationUsesPrintLayout() {
  const generator = new PDFGenerator();
  if (!(await generator.checkLibreOfficeInstalled())) {
    console.warn('soffice not installed; skipping printLayout PDF generation verification');
    return;
  }

  const template = await buildTemplate();
  const pdf = await generator.generatePDF(template.toString('base64'), buildData(20), {
    printLayout: {
      marginPreset: 'narrow',
      margins: { top: 0, bottom: 0, header: 0, footer: 0 },
      fit: { width: 1, height: 0 },
    },
  });

  if (pdf.length === 0 || pdf.subarray(0, 4).toString('utf8') !== '%PDF') {
    throw new Error('PDF generation with printLayout failed');
  }
}

async function verifyTallCellFixturePdfGeneration() {
  const generator = new PDFGenerator();
  if (!(await generator.checkLibreOfficeInstalled())) {
    console.warn('soffice not installed; skipping tall-cell fixture PDF verification');
    return;
  }

  const templatePath = path.join(__dirname, 'template_section_02_with-tall-cell.xlsx');
  const requestPath = path.join(__dirname, 'request_section_with-tall-cell.json');
  if (!fs.existsSync(templatePath) || !fs.existsSync(requestPath)) {
    console.warn('tall-cell fixture files not found; skipping tall-cell fixture PDF verification');
    return;
  }

  const templateBase64 = fs.readFileSync(templatePath).toString('base64');
  const data = JSON.parse(fs.readFileSync(requestPath, 'utf8'));
  const pdf = await generator.generatePDF(templateBase64, data, {
    printLayout: {
      marginPreset: 'narrow',
      margins: { top: 0, bottom: 0, header: 0, footer: 0 },
      fit: { width: 1, height: 0 },
    },
  });

  if (pdf.length === 0 || pdf.subarray(0, 4).toString('utf8') !== '%PDF') {
    throw new Error('tall-cell fixture PDF generation with printLayout failed');
  }
}

async function main() {
  await verifyMarginPresetNarrow();
  await verifyMarginPresetWithPartialOverride();
  await verifyInvalidPrintLayout();
  await verifyPdfGenerationUsesPrintLayout();
  await verifyTallCellFixturePdfGeneration();
  console.log('verify-print-layout: OK');
}

main().catch((err) => {
  console.error('verify-print-layout: FAILED');
  console.error(err);
  process.exit(1);
});
