import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { PlaceholderReplacer } from '../src/lib/placeholderReplacer';
import { ExcelGenerator } from '../src/lib/excelGenerator';

const SAMPLE_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAAAAAA6fptVAAAACklEQVR4nGNgAAAAAgABSK+kcQAAAABJRU5ErkJggg==';
const SAMPLE_JPEG_BASE64 =
  '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////2wBDAf//////////////////////////////////////////////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIQAxAAAAH/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oACAEBAAEFAqf/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oACAEDAQE/Aaf/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oACAECAQE/Aaf/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oACAEBAAY/Aqf/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oACAEBAAE/IV//2gAMAwEAAgADAAAAEP/EABQRAQAAAAAAAAAAAAAAAAAAABD/2gAIAQMBAT8QH//EABQRAQAAAAAAAAAAAAAAAAAAABD/2gAIAQIBAT8QH//EABQQAQAAAAAAAAAAAAAAAAAAABD/2gAIAQEAAT8QH//Z';

async function verifySingleSectionImagePlaceholder() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Image Sheet');

  worksheet.pageSetup = {
    paperSize: 9,
    orientation: 'portrait',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    printArea: 'A1:D20',
  };

  worksheet.mergeCells('B10:D10');
  worksheet.getRow(10).height = 121.7;
  worksheet.getCell('B10').value = '{{##請求.明細.image}}';

  const template = Buffer.from(await workbook.xlsx.writeBuffer());
  const replacer = new PlaceholderReplacer();

  const result = await replacer.replacePlaceholders(template, {
    請求: {
      明細: [
        {
          image: {
            base64: SAMPLE_PNG_BASE64,
            contentType: 'image/png',
            name: 'logo',
          },
        },
      ],
    },
  });

  const zip = await JSZip.loadAsync(result);
  const sharedStringsXml = await zip.file('xl/sharedStrings.xml')?.async('string');
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')?.async('string');
  const sheetRelsXml = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')?.async('string');
  const drawingXml = await zip.file('xl/drawings/drawing1.xml')?.async('string');
  const drawingRelsXml = await zip.file('xl/drawings/_rels/drawing1.xml.rels')?.async('string');
  const mediaBuffer = await zip.file('xl/media/image1.png')?.async('nodebuffer');

  if (!sharedStringsXml || !sheetXml || !sheetRelsXml || !drawingXml || !drawingRelsXml || !mediaBuffer) {
    throw new Error('generated xlsx structure is invalid');
  }

  if (sharedStringsXml.includes('__section_image_')) {
    throw new Error('generated section image token was not removed');
  }

  if (!sheetXml.includes('<drawing r:id="rId1"/>')) {
    throw new Error('sheet drawing reference was not added');
  }

  if (!sheetRelsXml.includes('Target="../drawings/drawing1.xml"')) {
    throw new Error('sheet relationship to drawing is missing');
  }

  if (!drawingRelsXml.includes('Target="../media/image1.png"')) {
    throw new Error('drawing relationship to image is missing');
  }

  if (!drawingXml.includes('name="logo"')) {
    throw new Error('image name was not preserved');
  }

  if (mediaBuffer.toString('base64') !== SAMPLE_PNG_BASE64) {
    throw new Error('embedded image content does not match input');
  }
}

async function verifySectionImagePlaceholderAcrossRecords() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Image Sheet');

  worksheet.pageSetup = {
    paperSize: 9,
    orientation: 'portrait',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    printArea: 'A1:D30',
  };

  worksheet.getCell('A2').value = '{{会社名}}';
  worksheet.mergeCells('B10:D10');
  worksheet.getRow(10).height = 90;
  worksheet.getCell('B10').value = '{{##請求.明細.image}}';

  const templateBase64 = Buffer.from(await workbook.xlsx.writeBuffer()).toString('base64');
  const generator = new ExcelGenerator();
  const result = await generator.generateExcel(templateBase64, {
    会社名: 'テスト株式会社',
    請求: {
      明細: [
        {
          image: {
            base64: SAMPLE_PNG_BASE64,
            contentType: 'image/png',
            name: 'youtube',
          },
        },
        {
          image: {
            base64: SAMPLE_JPEG_BASE64,
            contentType: 'image/jpeg',
            name: 'charactor',
          },
        },
      ],
    },
  });

  const zip = await JSZip.loadAsync(result);
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const sharedStringsXml = await zip.file('xl/sharedStrings.xml')?.async('string');
  const drawingXml = await zip.file('xl/drawings/drawing1.xml')?.async('string');
  const drawingRelsXml = await zip.file('xl/drawings/_rels/drawing1.xml.rels')?.async('string');

  if (!workbookXml || !sharedStringsXml || !drawingXml || !drawingRelsXml) {
    throw new Error('section multi-image output is invalid');
  }

  const mediaFiles = Object.keys(zip.files).filter((name) => /^xl\/media\/image\d+\.(png|jpe?g)$/i.test(name));
  if (mediaFiles.length !== 2) {
    throw new Error(`expected 2 embedded images, got ${mediaFiles.length}`);
  }

  if (sharedStringsXml.includes('__section_image_')) {
    throw new Error('generated section image tokens should not remain in shared strings');
  }

  const anchorCount = (drawingXml.match(/<xdr:twoCellAnchor\b/g) || []).length;
  if (anchorCount !== 2) {
    throw new Error(`expected 2 drawing anchors, got ${anchorCount}`);
  }

  if (!drawingXml.includes('name="youtube"') || !drawingXml.includes('name="charactor"')) {
    throw new Error('image names were not preserved in drawing XML');
  }

  const embeddedPayloads = await Promise.all(
    mediaFiles.map(async (filename) => (await zip.file(filename)?.async('nodebuffer'))?.toString('base64') || '')
  );
  const payloadSet = new Set(embeddedPayloads);
  if (!payloadSet.has(SAMPLE_PNG_BASE64) || !payloadSet.has(SAMPLE_JPEG_BASE64)) {
    throw new Error('embedded image content does not match section row input');
  }

  // テンプレートA1:D30に1行挿入されるため、単一範囲のままA1:D31へ拡張される
  const printAreaMatch = workbookXml.match(/name="_xlnm\.Print_Area"[^>]*>([^<]*)<\/definedName>/);
  const printArea = printAreaMatch ? printAreaMatch[1] : '';
  if (printArea.includes(',') || !printArea.includes('$A$1:$D$31')) {
    throw new Error(`unexpected print area after section image generation: ${printArea}`);
  }
}

async function verifyLegacyImagePlaceholderRejected() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Legacy Image Sheet');
  worksheet.getCell('B2').value = '{{%logo}}';

  const template = Buffer.from(await workbook.xlsx.writeBuffer());
  const replacer = new PlaceholderReplacer();

  let errorMessage = '';
  try {
    await replacer.replacePlaceholders(template, {});
  } catch (error) {
    errorMessage = error instanceof Error ? error.message : String(error);
  }

  if (!errorMessage.includes('Legacy image placeholder')) {
    throw new Error(`legacy image placeholder should be rejected, got: ${errorMessage || 'no error'}`);
  }
}

async function main() {
  await verifySingleSectionImagePlaceholder();
  await verifySectionImagePlaceholderAcrossRecords();
  await verifyLegacyImagePlaceholderRejected();
  console.log('verify-image-placeholder: OK');
}

main().catch((error) => {
  console.error('verify-image-placeholder: FAILED');
  console.error(error);
  process.exit(1);
});
