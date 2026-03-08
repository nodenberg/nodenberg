import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import * as fs from 'fs';
import * as path from 'path';
import { PlaceholderReplacer } from '../src/lib/placeholderReplacer';
import { ImagePlaceholderInput } from '../src/lib/imagePlaceholderReplacer';
import { ExcelGenerator } from '../src/lib/excelGenerator';

const SAMPLE_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAAAAAA6fptVAAAACklEQVR4nGNgAAAAAgABSK+kcQAAAABJRU5ErkJggg==';

async function verifySingleImagePlaceholder() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Image Sheet');

  worksheet.pageSetup = {
    paperSize: 9,
    orientation: 'portrait',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    printArea: 'A1:D69',
  };

  worksheet.mergeCells('B69:D69');
  worksheet.getRow(69).height = 121.7;
  worksheet.getCell('B69').value = '{{%logo}}';

  const template = Buffer.from(await workbook.xlsx.writeBuffer());
  const replacer = new PlaceholderReplacer();
  const images: Record<string, ImagePlaceholderInput> = {
    logo: {
      base64: SAMPLE_PNG_BASE64,
      contentType: 'image/png',
    },
  };

  const result = await replacer.replacePlaceholders(template, {}, { images });
  const zip = await JSZip.loadAsync(result);
  const sharedStringsXml = await zip.file('xl/sharedStrings.xml')?.async('string');
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')?.async('string');
  const sheetRelsXml = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')?.async('string');
  const drawingXml = await zip.file('xl/drawings/drawing1.xml')?.async('string');
  const drawingRelsXml = await zip.file('xl/drawings/_rels/drawing1.xml.rels')?.async('string');
  const mediaBuffer = await zip.file('xl/media/image1.png')?.async('nodebuffer');
  const contentTypesXml = await zip.file('[Content_Types].xml')?.async('string');

  if (!sharedStringsXml || !sheetXml || !sheetRelsXml || !drawingXml || !drawingRelsXml || !mediaBuffer || !contentTypesXml) {
    throw new Error('generated xlsx structure is invalid');
  }

  if (sharedStringsXml.includes('{{%logo}}')) {
    throw new Error('image placeholder was not removed from shared strings');
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

  if (!drawingXml.includes('<xdr:row>68</xdr:row>') || !drawingXml.includes('<xdr:row>69</xdr:row>')) {
    throw new Error('image anchor vertical span is incorrect');
  }

  if (
    drawingXml.includes('<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>') &&
    drawingXml.includes('<xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff>')
  ) {
    throw new Error('image should keep aspect ratio instead of stretching to the full width');
  }

  if (mediaBuffer.toString('base64') !== SAMPLE_PNG_BASE64) {
    throw new Error('embedded image content does not match input');
  }

  if (!contentTypesXml.includes('PartName="/xl/drawings/drawing1.xml"')) {
    throw new Error('drawing content type override is missing');
  }
}

async function verifyMultiImagePageBreakWithFixture() {
  const projectRoot = path.resolve(process.cwd(), '..');
  const templatePath = path.join(projectRoot, '06_test-assets', 'template_image_multi-image.xlsx');
  const requestPath = path.join(projectRoot, '06_test-assets', 'request_image.json');
  const youtubePath = path.join(projectRoot, '06_test-assets', 'youtube.png');
  const charactorPath = path.join(projectRoot, '06_test-assets', 'charactor.jpg');

  const templateBase64 = fs.readFileSync(templatePath).toString('base64');
  const data = JSON.parse(fs.readFileSync(requestPath, 'utf8'));
  const images = {
    youtube: {
      contentType: 'image/png',
      base64: fs.readFileSync(youtubePath).toString('base64'),
    },
    charactor: {
      contentType: 'image/jpeg',
      base64: fs.readFileSync(charactorPath).toString('base64'),
    },
  };

  const generator = new ExcelGenerator();
  const result = await generator.generateExcel(templateBase64, data, { images });
  const zip = await JSZip.loadAsync(result);
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const drawingXml = await zip.file('xl/drawings/drawing1.xml')?.async('string');
  const drawingRelsXml = await zip.file('xl/drawings/_rels/drawing1.xml.rels')?.async('string');

  if (!workbookXml || !drawingXml || !drawingRelsXml) {
    throw new Error('multi-image fixture output is invalid');
  }

  const mediaFiles = Object.keys(zip.files).filter((name) => /^xl\/media\/image\d+\.(png|jpe?g)$/i.test(name));
  if (mediaFiles.length !== 2) {
    throw new Error(`expected 2 embedded images, got ${mediaFiles.length}`);
  }

  if (!drawingRelsXml.includes('Target="../media/image1.png"') || !drawingRelsXml.includes('Target="../media/image2.jpeg"')) {
    throw new Error('multiple image relationships are missing');
  }

  const anchorCount = (drawingXml.match(/<xdr:twoCellAnchor\b/g) || []).length;
  if (anchorCount !== 2) {
    throw new Error(`expected 2 drawing anchors, got ${anchorCount}`);
  }

  if (!drawingXml.includes('name="youtube"') || !drawingXml.includes('name="charactor"')) {
    throw new Error('image names were not preserved in drawing XML');
  }

  const printAreaMatch = workbookXml.match(/name="_xlnm\.Print_Area"[^>]*>([^<]*)<\/definedName>/);
  const printArea = printAreaMatch ? printAreaMatch[1] : '';
  const ranges = printArea.split(',').map((range) => {
    const match = range.match(/\$[A-Z]+\$(\d+):\$[A-Z]+\$(\d+)/);
    if (!match) {
      throw new Error(`invalid print area range: ${range}`);
    }
    return { startRow: Number(match[1]), endRow: Number(match[2]) };
  });
  if (ranges.length !== 2 || ranges[0].startRow !== 1 || ranges[0].endRow !== 35 || ranges[1].startRow !== 36 || ranges[1].endRow !== 54) {
    throw new Error(`unexpected print area after multi-image generation: ${printArea}`);
  }

  const rowAnchors = Array.from(drawingXml.matchAll(/<xdr:row>(\d+)<\/xdr:row>/g)).map((match) => Number(match[1]));
  if (rowAnchors.length < 4) {
    throw new Error('image row anchors are missing');
  }

  const secondImageFromRow = rowAnchors[2];
  if (secondImageFromRow !== 35) {
    throw new Error(`second image was not moved to the next page boundary: fromRow=${secondImageFromRow}`);
  }
}

async function main() {
  await verifySingleImagePlaceholder();
  await verifyMultiImagePageBreakWithFixture();
  console.log('verify-image-placeholder: OK');
}

main().catch((error) => {
  console.error('verify-image-placeholder: FAILED');
  console.error(error);
  process.exit(1);
});
