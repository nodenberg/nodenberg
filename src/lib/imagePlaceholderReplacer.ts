import JSZip from 'jszip';

export interface ImagePlaceholderInput {
  base64: string;
  contentType?: string;
  extension?: string;
  name?: string;
}

export type PlaceholderImages = Record<string, ImagePlaceholderInput>;

type ImagePlacement = {
  imageKey: string;
  cellRef: string;
  range: CellRange;
};

type CellRange = {
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
};

type MediaInfo = {
  extension: string;
  contentType: string;
  path: string;
  buffer: Buffer;
};

type NativeAnchorPosition = {
  nativeCol: number;
  nativeColOff: number;
  nativeRow: number;
  nativeRowOff: number;
};

type SheetGeometry = {
  defaultColWidth: number;
  defaultRowHeight: number;
  columnWidths: Map<number, number>;
  rowHeights: Map<number, number>;
};

type AxisSegment = {
  pixels: number;
  native: number;
};

type PrintAreaRange = {
  startRow: number;
  endRow: number;
};

type WorkbookSheetEntry = {
  name: string;
  index: number;
  path: string;
};

const IMAGE_BOX_PADDING_PX = 2;
const DEBUG_IMAGE_LAYOUT = /^(1|true|yes|on)$/i.test(process.env.DEBUG_IMAGE_LAYOUT || '');
const EMU_PER_PIXEL = 9525;

type ContainedAnchorDebug = {
  imageWidth: number;
  imageHeight: number;
  boxWidth: number;
  boxHeight: number;
  innerBoxWidth: number;
  innerBoxHeight: number;
  scale: number;
  renderedWidth: number;
  renderedHeight: number;
  offsetX: number;
  offsetY: number;
};

type ContainedAnchorResult = {
  tl: NativeAnchorPosition;
  br: NativeAnchorPosition;
  debug: ContainedAnchorDebug;
};

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function decodeXml(value: string): string {
  return value
    .replace(/&apos;/g, '\'')
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, '>')
    .replace(/&lt;/g, '<')
    .replace(/&amp;/g, '&');
}

function getXmlAttribute(xml: string, name: string): string | null {
  const match = xml.match(new RegExp(`\\b${escapeRegExp(name)}="([^"]*)"`));
  return match ? match[1] : null;
}

function columnLettersToNumber(columnLetters: string): number {
  let result = 0;
  for (const char of columnLetters) {
    result = result * 26 + (char.charCodeAt(0) - 64);
  }
  return result;
}

function parseCellRef(cellRef: string): { col: number; row: number } {
  const match = cellRef.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid cell reference: ${cellRef}`);
  }

  return {
    col: columnLettersToNumber(match[1]),
    row: Number(match[2]),
  };
}

function parseCellRange(rangeRef: string): CellRange {
  const parts = rangeRef.split(':');
  const start = parseCellRef(parts[0]);
  const end = parseCellRef(parts[parts.length - 1]);

  return {
    startCol: start.col,
    startRow: start.row,
    endCol: end.col,
    endRow: end.row,
  };
}

function parsePrintAreaRanges(value: string): PrintAreaRange[] {
  return value
    .split(',')
    .map((part) => part.trim())
    .map((part): PrintAreaRange | null => {
      const bangIndex = part.indexOf('!');
      if (bangIndex === -1) return null;

      const rangePart = part.slice(bangIndex + 1);
      const match = rangePart.match(/\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/);
      if (!match) return null;

      return {
        startRow: Number(match[2]),
        endRow: Number(match[4]),
      };
    })
    .filter((range): range is PrintAreaRange => range !== null);
}

function normalizeSheetPrefix(prefix: string): string {
  const decoded = decodeXml(prefix).trim();
  if (decoded.startsWith('\'') && decoded.endsWith('\'')) {
    return decoded.slice(1, -1);
  }
  return decoded;
}

function parseWorkbookSheets(workbookXml: string, workbookRelsXml: string): WorkbookSheetEntry[] {
  const sheetsMatch = workbookXml.match(/<sheets>([\s\S]*?)<\/sheets>/i);
  if (!sheetsMatch) return [];

  const relationshipMap = new Map<string, string>();
  const relRegex = /<Relationship\b([^>]*)\/>/g;
  let relMatch: RegExpExecArray | null;

  while ((relMatch = relRegex.exec(workbookRelsXml)) !== null) {
    const attrs = relMatch[1];
    const idMatch = attrs.match(/\bId="([^"]+)"/);
    const targetMatch = attrs.match(/\bTarget="([^"]+)"/);
    if (!idMatch || !targetMatch) continue;

    let target = targetMatch[1];
    if (target.startsWith('/')) target = target.slice(1);
    if (!target.startsWith('xl/')) target = `xl/${target.replace(/^\.?\/?/, '')}`;
    relationshipMap.set(idMatch[1], target);
  }

  const sheetTags = sheetsMatch[1].match(/<sheet\b[^>]*\/>/g) || [];
  return sheetTags.map((tag, index) => {
    const nameMatch = tag.match(/\bname="([^"]*)"/);
    const relIdMatch = tag.match(/\br:id="([^"]+)"/);
    const relId = relIdMatch ? relIdMatch[1] : '';

    return {
      name: nameMatch ? decodeXml(nameMatch[1]) : `Sheet${index + 1}`,
      index,
      path: relationshipMap.get(relId) || '',
    };
  });
}

function parseManualRowBreakIds(sheetXml: string): number[] {
  const block = sheetXml.match(/<rowBreaks\b[^>]*>([\s\S]*?)<\/rowBreaks>/);
  if (!block) return [];

  const ids: number[] = [];
  for (const match of block[1].matchAll(/<(?:brk|rowBreak)\b[^>]*\bid="(\d+)"[^>]*\/>/g)) {
    ids.push(Number(match[1]));
  }
  return ids;
}

/**
 * 印刷範囲を手動改ページ（rowBreaks）の位置で分割し、ページ単位の行範囲に展開する。
 * brk id=N は「N行目の直後で改ページ」を意味する。
 */
function splitRangesByRowBreaks(ranges: PrintAreaRange[], breakIds: number[]): PrintAreaRange[] {
  if (breakIds.length === 0) return ranges;

  const sortedBreaks = Array.from(new Set(breakIds)).sort((a, b) => a - b);
  const result: PrintAreaRange[] = [];

  for (const range of ranges) {
    let currentStart = range.startRow;
    for (const breakId of sortedBreaks) {
      if (breakId < currentStart || breakId >= range.endRow) continue;
      result.push({ ...range, startRow: currentStart, endRow: breakId });
      currentStart = breakId + 1;
    }
    result.push({ ...range, startRow: currentStart, endRow: range.endRow });
  }

  return result;
}

async function getPrintAreaRangesBySheetPath(zip: JSZip): Promise<Map<string, PrintAreaRange[]>> {
  const workbookFile = zip.file('xl/workbook.xml');
  const workbookRelsFile = zip.file('xl/_rels/workbook.xml.rels');
  if (!workbookFile || !workbookRelsFile) return new Map();

  const workbookXml = await workbookFile.async('string');
  const workbookRelsXml = await workbookRelsFile.async('string');
  const sheetEntries = parseWorkbookSheets(workbookXml, workbookRelsXml).filter((entry) => !!entry.path);
  const result = new Map<string, PrintAreaRange[]>();
  const tags = workbookXml.match(/<definedName\b[\s\S]*?<\/definedName>/g) || [];

  for (const sheetEntry of sheetEntries) {
    for (const tag of tags) {
      if (!/name="_xlnm\.Print_Area"/.test(tag)) continue;

      const openTagMatch = tag.match(/<definedName\b[^>]*>/);
      const valueMatch = tag.match(/>([^<]*)<\/definedName>/);
      if (!openTagMatch || !valueMatch) continue;

      const openTag = openTagMatch[0];
      const value = valueMatch[1];
      const firstRangePart = value.split(',')[0];
      const bangIndex = firstRangePart.indexOf('!');
      if (bangIndex === -1) continue;

      const sheetNameFromValue = normalizeSheetPrefix(firstRangePart.slice(0, bangIndex));
      const localSheetIdMatch = openTag.match(/\blocalSheetId="(\d+)"/);
      const localSheetId = localSheetIdMatch ? Number(localSheetIdMatch[1]) : null;
      const isTarget =
        (localSheetId !== null && localSheetId === sheetEntry.index) ||
        sheetNameFromValue === sheetEntry.name;

      if (!isTarget) continue;

      const sheetFile = zip.file(sheetEntry.path);
      const sheetXml = sheetFile ? await sheetFile.async('string') : '';
      const pageRanges = splitRangesByRowBreaks(
        parsePrintAreaRanges(value),
        sheetXml ? parseManualRowBreakIds(sheetXml) : []
      );
      result.set(sheetEntry.path, pageRanges);
      break;
    }
  }

  return result;
}

function findImagePlaceholderIndices(
  sharedStringsXml: string,
  availableKeys: Set<string>
): Map<number, string> {
  const result = new Map<number, string>();
  const entries = parseSharedStringEntries(sharedStringsXml);

  entries.forEach((entry) => {
    const placeholderMatch = entry.text.match(/^\{\{%([^}]+)\}\}$/);
    if (!placeholderMatch) return;

    const imageKey = placeholderMatch[1];
    if (availableKeys.has(imageKey)) {
      result.set(entry.index, imageKey);
    }
  });

  return result;
}

function parseSharedStringEntries(sharedStringsXml: string): Array<{
  index: number;
  xml: string;
  text: string;
}> {
  const entries: Array<{ index: number; xml: string; text: string }> = [];
  const siRegex = /<si\b[^>]*>[\s\S]*?<\/si>/g;
  let match: RegExpExecArray | null;
  let index = 0;

  while ((match = siRegex.exec(sharedStringsXml)) !== null) {
    const xml = match[0];
    const textMatches = Array.from(xml.matchAll(/<t(?:\s+[^>]*)?>([\s\S]*?)<\/t>/g));
    const text = textMatches.map((textMatch) => decodeXmlText(textMatch[1])).join('');

    entries.push({
      index,
      xml,
      text,
    });

    index += 1;
  }

  return entries;
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, '\'')
    .replace(/&amp;/g, '&');
}

function clearImagePlaceholderSharedStrings(
  sharedStringsXml: string,
  imageKeys: Set<string>
): string {
  let updatedXml = sharedStringsXml;
  const entries = parseSharedStringEntries(sharedStringsXml);

  entries.forEach((entry) => {
    const placeholderMatch = entry.text.match(/^\{\{%([^}]+)\}\}$/);
    if (!placeholderMatch) return;
    if (!imageKeys.has(placeholderMatch[1])) return;

    updatedXml = updatedXml.replace(entry.xml, '<si><t></t></si>');
  });

  return updatedXml;
}

function listWorksheetPaths(zip: JSZip): string[] {
  return Object.keys(zip.files)
    .filter((filename) => /^xl\/worksheets\/sheet\d+\.xml$/.test(filename))
    .sort((a, b) => {
      const aNum = Number(a.match(/sheet(\d+)\.xml$/)?.[1] || 0);
      const bNum = Number(b.match(/sheet(\d+)\.xml$/)?.[1] || 0);
      return aNum - bNum;
    });
}

function parseMergedRanges(sheetXml: string): CellRange[] {
  const mergeBlock = sheetXml.match(/<mergeCells[^>]*>([\s\S]*?)<\/mergeCells>/);
  if (!mergeBlock) return [];

  const ranges: CellRange[] = [];
  const mergeCellRegex = /<mergeCell ref="([A-Z0-9:]+)"\/>/g;
  let match: RegExpExecArray | null;

  while ((match = mergeCellRegex.exec(mergeBlock[1])) !== null) {
    ranges.push(parseCellRange(match[1]));
  }

  return ranges;
}

function findAnchorRange(cellRef: string, mergedRanges: CellRange[]): CellRange {
  const cell = parseCellRef(cellRef);
  const merged = mergedRanges.find((range) => (
    cell.col >= range.startCol &&
    cell.col <= range.endCol &&
    cell.row >= range.startRow &&
    cell.row <= range.endRow
  ));

  return merged || {
    startCol: cell.col,
    startRow: cell.row,
    endCol: cell.col,
    endRow: cell.row,
  };
}

function findImagePlacements(sheetXml: string, indexToImageKey: Map<number, string>): ImagePlacement[] {
  if (indexToImageKey.size === 0) return [];

  const mergedRanges = parseMergedRanges(sheetXml);
  const placements: ImagePlacement[] = [];
  const cellRegex = /<c\b[^>]*r="([A-Z]+\d+)"[^>]*\bt="s"[^>]*>[\s\S]*?<v>(\d+)<\/v>[\s\S]*?<\/c>/g;
  let match: RegExpExecArray | null;

  while ((match = cellRegex.exec(sheetXml)) !== null) {
    const cellRef = match[1];
    const sharedStringIndex = Number(match[2]);
    const imageKey = indexToImageKey.get(sharedStringIndex);
    if (!imageKey) continue;

    placements.push({
      imageKey,
      cellRef,
      range: findAnchorRange(cellRef, mergedRanges),
    });
  }

  return placements;
}

function getNextNumericSuffix(paths: string[], regex: RegExp): number {
  const numbers = paths
    .map((path) => Number(path.match(regex)?.[1] || 0))
    .filter((value) => Number.isFinite(value));

  return numbers.length > 0 ? Math.max(...numbers) + 1 : 1;
}

function inferImageMeta(image: ImagePlaceholderInput): { extension: string; contentType: string } {
  const rawExtension = (image.extension || '').replace(/^\./, '').toLowerCase();
  const rawContentType = (image.contentType || '').toLowerCase();
  const fromContentType: Record<string, { extension: string; contentType: string }> = {
    'image/png': { extension: 'png', contentType: 'image/png' },
    'image/jpeg': { extension: 'jpeg', contentType: 'image/jpeg' },
    'image/jpg': { extension: 'jpeg', contentType: 'image/jpeg' },
  };

  if (rawExtension === 'png') return { extension: 'png', contentType: 'image/png' };
  if (rawExtension === 'jpg' || rawExtension === 'jpeg') {
    return { extension: 'jpeg', contentType: 'image/jpeg' };
  }

  if (rawContentType && fromContentType[rawContentType]) {
    return fromContentType[rawContentType];
  }

  return { extension: 'png', contentType: 'image/png' };
}

function parseSheetGeometry(sheetXml: string): SheetGeometry {
  const defaultColWidth = Number(sheetXml.match(/<sheetFormatPr[^>]*defaultColWidth="([^"]+)"/)?.[1] || 8.43);
  const defaultRowHeight = Number(sheetXml.match(/<sheetFormatPr[^>]*defaultRowHeight="([^"]+)"/)?.[1] || 15);
  const columnWidths = new Map<number, number>();
  const rowHeights = new Map<number, number>();

  const colRegex = /<col\b[^>]*\/?>/g;
  let colMatch: RegExpExecArray | null;
  while ((colMatch = colRegex.exec(sheetXml)) !== null) {
    const colXml = colMatch[0];
    const min = Number(getXmlAttribute(colXml, 'min'));
    const max = Number(getXmlAttribute(colXml, 'max'));
    const width = Number(getXmlAttribute(colXml, 'width'));
    if (!Number.isFinite(min) || !Number.isFinite(max) || !Number.isFinite(width)) continue;
    for (let col = min; col <= max; col++) {
      columnWidths.set(col, width);
    }
  }

  const rowRegex = /<row\b[^>]*>/g;
  let rowMatch: RegExpExecArray | null;
  while ((rowMatch = rowRegex.exec(sheetXml)) !== null) {
    const rowXml = rowMatch[0];
    const rowNumber = Number(getXmlAttribute(rowXml, 'r'));
    const height = Number(getXmlAttribute(rowXml, 'ht'));
    if (!Number.isFinite(rowNumber) || !Number.isFinite(height)) continue;
    rowHeights.set(rowNumber, height);
  }

  return {
    defaultColWidth,
    defaultRowHeight,
    columnWidths,
    rowHeights,
  };
}

function getColumnWidthChars(geometry: SheetGeometry, colNumber: number): number {
  return geometry.columnWidths.get(colNumber) ?? geometry.defaultColWidth;
}

function getRowHeightPoints(geometry: SheetGeometry, rowNumber: number): number {
  return geometry.rowHeights.get(rowNumber) ?? geometry.defaultRowHeight;
}

function getColumnNativeWidth(geometry: SheetGeometry, colNumber: number): number {
  return Math.round(columnWidthToPixels(getColumnWidthChars(geometry, colNumber)) * EMU_PER_PIXEL);
}

function getRowNativeHeight(geometry: SheetGeometry, rowNumber: number): number {
  return Math.round(rowHeightToPixels(getRowHeightPoints(geometry, rowNumber)) * EMU_PER_PIXEL);
}

function columnWidthToPixels(widthChars: number): number {
  if (widthChars <= 0) return 0;
  if (widthChars < 1) {
    return Math.floor(widthChars * 12 + 0.5);
  }
  return Math.floor(((256 * widthChars + Math.floor(128 / 7)) / 256) * 7);
}

function rowHeightToPixels(heightPoints: number): number {
  return Math.floor((heightPoints * 96) / 72);
}

function getImageDimensions(buffer: Buffer): { width: number; height: number } {
  if (buffer.length >= 24 && buffer.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))) {
    return {
      width: buffer.readUInt32BE(16),
      height: buffer.readUInt32BE(20),
    };
  }

  if (buffer.length >= 4 && buffer[0] === 0xff && buffer[1] === 0xd8) {
    let offset = 2;
    while (offset < buffer.length) {
      if (buffer[offset] !== 0xff) {
        offset += 1;
        continue;
      }

      const marker = buffer[offset + 1];
      if (marker === 0xd8 || marker === 0xd9) {
        offset += 2;
        continue;
      }

      const segmentLength = buffer.readUInt16BE(offset + 2);
      const isSofMarker = (
        (marker >= 0xc0 && marker <= 0xc3) ||
        (marker >= 0xc5 && marker <= 0xc7) ||
        (marker >= 0xc9 && marker <= 0xcb) ||
        (marker >= 0xcd && marker <= 0xcf)
      );

      if (isSofMarker) {
        return {
          width: buffer.readUInt16BE(offset + 7),
          height: buffer.readUInt16BE(offset + 5),
        };
      }

      offset += 2 + segmentLength;
    }
  }

  throw new Error('Unsupported image format for placeholder embedding');
}

function buildAxisSegments(params: {
  start: number;
  end: number;
  getPixels: (index: number) => number;
  getNative: (index: number) => number;
}): AxisSegment[] {
  const segments: AxisSegment[] = [];
  for (let index = params.start; index <= params.end; index++) {
    segments.push({
      pixels: params.getPixels(index),
      native: params.getNative(index),
    });
  }
  return segments;
}

function locateAxisPosition(
  segments: AxisSegment[],
  startNativeIndex: number,
  positionPixels: number
): { nativeIndex: number; nativeOffset: number } {
  const totalPixels = segments.reduce((sum, segment) => sum + segment.pixels, 0);
  const clampedPosition = Math.max(0, Math.min(totalPixels, positionPixels));

  if (clampedPosition === totalPixels) {
    return {
      nativeIndex: startNativeIndex + segments.length,
      nativeOffset: 0,
    };
  }

  let currentPixels = 0;
  for (let i = 0; i < segments.length; i++) {
    const segment = segments[i];
    const nextPixels = currentPixels + segment.pixels;
    if (clampedPosition < nextPixels || segment.pixels === 0) {
      const fraction = segment.pixels > 0 ? (clampedPosition - currentPixels) / segment.pixels : 0;
      return {
        nativeIndex: startNativeIndex + i,
        nativeOffset: Math.floor(fraction * segment.native),
      };
    }
    currentPixels = nextPixels;
  }

  return {
    nativeIndex: startNativeIndex + segments.length,
    nativeOffset: 0,
  };
}

function buildContainedAnchor(
  range: CellRange,
  geometry: SheetGeometry,
  imageBuffer: Buffer
): ContainedAnchorResult {
  const imageDimensions = getImageDimensions(imageBuffer);
  const colSegments = buildAxisSegments({
    start: range.startCol,
    end: range.endCol,
    getPixels: (col) => columnWidthToPixels(getColumnWidthChars(geometry, col)),
    getNative: (col) => getColumnNativeWidth(geometry, col),
  });
  const rowSegments = buildAxisSegments({
    start: range.startRow,
    end: range.endRow,
    getPixels: (row) => rowHeightToPixels(getRowHeightPoints(geometry, row)),
    getNative: (row) => getRowNativeHeight(geometry, row),
  });

  const boxWidth = Math.max(1, colSegments.reduce((sum, segment) => sum + segment.pixels, 0));
  const boxHeight = Math.max(1, rowSegments.reduce((sum, segment) => sum + segment.pixels, 0));
  const innerBoxWidth = Math.max(1, boxWidth - IMAGE_BOX_PADDING_PX * 2);
  const innerBoxHeight = Math.max(1, boxHeight - IMAGE_BOX_PADDING_PX * 2);
  const scale = Math.min(innerBoxWidth / imageDimensions.width, innerBoxHeight / imageDimensions.height);
  const renderedWidth = imageDimensions.width * scale;
  const renderedHeight = imageDimensions.height * scale;
  const offsetX = IMAGE_BOX_PADDING_PX + (innerBoxWidth - renderedWidth) / 2;
  const offsetY = IMAGE_BOX_PADDING_PX + (innerBoxHeight - renderedHeight) / 2;

  const fromCol = locateAxisPosition(colSegments, range.startCol - 1, offsetX);
  const toCol = locateAxisPosition(colSegments, range.startCol - 1, offsetX + renderedWidth);
  const fromRow = locateAxisPosition(rowSegments, range.startRow - 1, offsetY);
  const toRow = locateAxisPosition(rowSegments, range.startRow - 1, offsetY + renderedHeight);

  return {
    tl: {
      nativeCol: fromCol.nativeIndex,
      nativeColOff: fromCol.nativeOffset,
      nativeRow: fromRow.nativeIndex,
      nativeRowOff: fromRow.nativeOffset,
    },
    br: {
      nativeCol: toCol.nativeIndex,
      nativeColOff: toCol.nativeOffset,
      nativeRow: toRow.nativeIndex,
      nativeRowOff: toRow.nativeOffset,
    },
    debug: {
      imageWidth: imageDimensions.width,
      imageHeight: imageDimensions.height,
      boxWidth,
      boxHeight,
      innerBoxWidth,
      innerBoxHeight,
      scale,
      renderedWidth,
      renderedHeight,
      offsetX,
      offsetY,
    },
  };
}

function moveAnchorRows(
  anchor: { tl: NativeAnchorPosition; br: NativeAnchorPosition },
  rowDelta: number
): { tl: NativeAnchorPosition; br: NativeAnchorPosition } {
  if (rowDelta <= 0) return anchor;

  return {
    tl: {
      ...anchor.tl,
      nativeRow: anchor.tl.nativeRow + rowDelta,
    },
    br: {
      ...anchor.br,
      nativeRow: anchor.br.nativeRow + rowDelta,
    },
  };
}

function keepAnchorWithinSinglePrintPage(
  anchor: { tl: NativeAnchorPosition; br: NativeAnchorPosition },
  printAreaRanges: PrintAreaRange[]
): { tl: NativeAnchorPosition; br: NativeAnchorPosition } {
  if (printAreaRanges.length === 0) return anchor;

  const startRow = anchor.tl.nativeRow + 1;
  const endRow = anchor.br.nativeRow + (anchor.br.nativeRowOff > 0 ? 1 : 0);
  const crossingRange = printAreaRanges.find((range) =>
    startRow >= range.startRow &&
    startRow <= range.endRow &&
    endRow > range.endRow
  );

  if (!crossingRange) return anchor;

  const rowDelta = crossingRange.endRow + 1 - startRow;
  return moveAnchorRows(anchor, rowDelta);
}

function buildSheetRelsPath(sheetPath: string): string {
  const filename = sheetPath.split('/').pop();
  return `xl/worksheets/_rels/${filename}.rels`;
}

function createRelationshipsXml(): string {
  return '<?xml version="1.0" encoding="UTF-8"?>\n'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
}

function createDrawingXml(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
    + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>';
}

function appendRelationship(xml: string, relationship: string): string {
  return xml.replace('</Relationships>', `${relationship}</Relationships>`);
}

function ensureDrawingTag(sheetXml: string, relationshipId: string): string {
  if (/<drawing\b[^>]*r:id="[^"]+"\/>/.test(sheetXml)) {
    return sheetXml;
  }

  return sheetXml.replace('</worksheet>', `<drawing r:id="${relationshipId}"/></worksheet>`);
}

function ensureContentTypeDefault(contentTypesXml: string, extension: string, contentType: string): string {
  if (new RegExp(`<Default Extension="${escapeRegExp(extension)}"`).test(contentTypesXml)) {
    return contentTypesXml;
  }

  return contentTypesXml.replace(
    '</Types>',
    `<Default Extension="${extension}" ContentType="${contentType}"/></Types>`
  );
}

function ensureContentTypeOverride(contentTypesXml: string, partName: string, contentType: string): string {
  if (new RegExp(`PartName="${escapeRegExp(partName)}"`).test(contentTypesXml)) {
    return contentTypesXml;
  }

  return contentTypesXml.replace(
    '</Types>',
    `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`
  );
}

function getMaxRelationshipId(xml: string): number {
  const matches = Array.from(xml.matchAll(/Id="rId(\d+)"/g)).map((match) => Number(match[1]));
  return matches.length > 0 ? Math.max(...matches) : 0;
}

function getMaxPictureId(xml: string): number {
  const matches = Array.from(xml.matchAll(/<xdr:cNvPr id="(\d+)"/g)).map((match) => Number(match[1]));
  return matches.length > 0 ? Math.max(...matches) : 0;
}

function buildImageAnchorXml(params: {
  imageName: string;
  imageRelId: string;
  pictureId: number;
  tl: NativeAnchorPosition;
  br: NativeAnchorPosition;
}): string {
  return ''
    + '<xdr:twoCellAnchor editAs="oneCell">'
    + `<xdr:from><xdr:col>${params.tl.nativeCol}</xdr:col><xdr:colOff>${params.tl.nativeColOff}</xdr:colOff><xdr:row>${params.tl.nativeRow}</xdr:row><xdr:rowOff>${params.tl.nativeRowOff}</xdr:rowOff></xdr:from>`
    + `<xdr:to><xdr:col>${params.br.nativeCol}</xdr:col><xdr:colOff>${params.br.nativeColOff}</xdr:colOff><xdr:row>${params.br.nativeRow}</xdr:row><xdr:rowOff>${params.br.nativeRowOff}</xdr:rowOff></xdr:to>`
    + '<xdr:pic>'
    + `<xdr:nvPicPr><xdr:cNvPr id="${params.pictureId}" name="${params.imageName}"/>`
    + '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
    + '<xdr:blipFill>'
    + `<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${params.imageRelId}" cstate="print"/>`
    + '<a:stretch><a:fillRect/></a:stretch>'
    + '</xdr:blipFill>'
    + '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
    + '</xdr:pic>'
    + '<xdr:clientData/>'
    + '</xdr:twoCellAnchor>';
}

function appendDrawingAnchor(drawingXml: string, anchorXml: string): string {
  return drawingXml.replace('</xdr:wsDr>', `${anchorXml}</xdr:wsDr>`);
}

function adjustDrawingXmlToPrintPages(drawingXml: string, printAreaRanges: PrintAreaRange[]): string {
  if (printAreaRanges.length === 0) return drawingXml;

  return drawingXml.replace(/<xdr:twoCellAnchor[\s\S]*?<\/xdr:twoCellAnchor>/g, (anchorXml) => {
    const fromBlock = anchorXml.match(/<xdr:from>([\s\S]*?)<\/xdr:from>/);
    const toBlock = anchorXml.match(/<xdr:to>([\s\S]*?)<\/xdr:to>/);
    if (!fromBlock || !toBlock) return anchorXml;

    const fromRowMatch = fromBlock[1].match(/<xdr:row>(\d+)<\/xdr:row>/);
    const toRowMatch = toBlock[1].match(/<xdr:row>(\d+)<\/xdr:row>/);
    const toRowOffMatch = toBlock[1].match(/<xdr:rowOff>(\d+)<\/xdr:rowOff>/);
    if (!fromRowMatch || !toRowMatch) return anchorXml;

    const startRow = Number(fromRowMatch[1]) + 1;
    const endRow = Number(toRowMatch[1]) + (Number(toRowOffMatch?.[1] || 0) > 0 ? 1 : 0);
    const crossingRange = printAreaRanges.find((range) =>
      startRow >= range.startRow &&
      startRow <= range.endRow &&
      endRow > range.endRow
    );

    if (!crossingRange) return anchorXml;

    const rowDelta = crossingRange.endRow + 1 - startRow;
    const updatedFrom = fromBlock[1].replace(
      /<xdr:row>\d+<\/xdr:row>/,
      `<xdr:row>${Number(fromRowMatch[1]) + rowDelta}</xdr:row>`
    );
    const updatedTo = toBlock[1].replace(
      /<xdr:row>\d+<\/xdr:row>/,
      `<xdr:row>${Number(toRowMatch[1]) + rowDelta}</xdr:row>`
    );

    return anchorXml
      .replace(fromBlock[0], `<xdr:from>${updatedFrom}</xdr:from>`)
      .replace(toBlock[0], `<xdr:to>${updatedTo}</xdr:to>`);
  });
}

type EnsureDrawingResult = {
  sheetXml: string;
  sheetRelsXml: string;
  drawingPath: string;
  drawingRelsPath: string;
};

function ensureDrawingForSheet(
  zip: JSZip,
  sheetPath: string,
  sheetXml: string,
  sheetRelsXml: string
): EnsureDrawingResult {
  const existingDrawing = sheetRelsXml.match(
    /<Relationship Id="(rId\d+)" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/drawing" Target="\.\.\/drawings\/(drawing\d+\.xml)"\/>/
  );

  if (existingDrawing) {
    const drawingFilename = existingDrawing[2];
    return {
      sheetXml: ensureDrawingTag(sheetXml, existingDrawing[1]),
      sheetRelsXml,
      drawingPath: `xl/drawings/${drawingFilename}`,
      drawingRelsPath: `xl/drawings/_rels/${drawingFilename}.rels`,
    };
  }

  const nextDrawingNumber = getNextNumericSuffix(Object.keys(zip.files), /drawing(\d+)\.xml$/);
  const drawingFilename = `drawing${nextDrawingNumber}.xml`;
  const drawingPath = `xl/drawings/${drawingFilename}`;
  const drawingRelsPath = `xl/drawings/_rels/${drawingFilename}.rels`;
  const nextRelId = `rId${getMaxRelationshipId(sheetRelsXml) + 1}`;
  const relationship = `<Relationship Id="${nextRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/${drawingFilename}"/>`;

  zip.file(drawingPath, createDrawingXml());
  zip.file(drawingRelsPath, createRelationshipsXml());

  return {
    sheetXml: ensureDrawingTag(sheetXml, nextRelId),
    sheetRelsXml: appendRelationship(sheetRelsXml, relationship),
    drawingPath,
    drawingRelsPath,
  };
}

export async function embedImagePlaceholders(params: {
  zip: JSZip;
  sharedStringsXml: string;
  images?: PlaceholderImages;
}): Promise<string> {
  const { zip, images } = params;
  let sharedStringsXml = params.sharedStringsXml;

  if (!images || Object.keys(images).length === 0) {
    return sharedStringsXml;
  }

  const imageKeys = new Set(Object.keys(images));
  const indexToImageKey = findImagePlaceholderIndices(sharedStringsXml, imageKeys);
  if (indexToImageKey.size === 0) {
    return sharedStringsXml;
  }

  const worksheetPaths = listWorksheetPaths(zip);
  const sheetPlacements = new Map<string, ImagePlacement[]>();

  for (const sheetPath of worksheetPaths) {
    const sheetFile = zip.file(sheetPath);
    if (!sheetFile) continue;
    const sheetXml = await sheetFile.async('string');
    const placements = findImagePlacements(sheetXml, indexToImageKey);
    if (placements.length > 0) {
      sheetPlacements.set(sheetPath, placements);
    }
  }

  if (sheetPlacements.size === 0) {
    return sharedStringsXml;
  }

  const mediaInfos = new Map<string, MediaInfo>();
  let nextImageNumber = getNextNumericSuffix(Object.keys(zip.files), /image(\d+)\.[^.]+$/);

  for (const imageKey of new Set(Array.from(sheetPlacements.values()).flat().map((item) => item.imageKey))) {
    const image = images[imageKey];
    if (!image || typeof image.base64 !== 'string' || image.base64.trim() === '') {
      throw new Error(`Image placeholder "${imageKey}" requires a non-empty base64 string`);
    }

    const meta = inferImageMeta(image);
    const buffer = Buffer.from(image.base64, 'base64');
    const path = `xl/media/image${nextImageNumber}.${meta.extension}`;
    nextImageNumber += 1;

    mediaInfos.set(imageKey, {
      ...meta,
      path,
      buffer,
    });
  }

  let contentTypesXml = await zip.file('[Content_Types].xml')?.async('string');
  if (!contentTypesXml) {
    throw new Error('[Content_Types].xmlが見つかりません');
  }

  const printAreaRangesBySheet = await getPrintAreaRangesBySheetPath(zip);

  for (const mediaInfo of mediaInfos.values()) {
    zip.file(mediaInfo.path, mediaInfo.buffer);
    contentTypesXml = ensureContentTypeDefault(contentTypesXml, mediaInfo.extension, mediaInfo.contentType);
    contentTypesXml = ensureContentTypeOverride(contentTypesXml, `/${mediaInfo.path}`, mediaInfo.contentType);
  }

  for (const [sheetPath, placements] of sheetPlacements.entries()) {
    const sheetFile = zip.file(sheetPath);
    if (!sheetFile) continue;

    let sheetXml = await sheetFile.async('string');
    const geometry = parseSheetGeometry(sheetXml);
    const sheetRelsPath = buildSheetRelsPath(sheetPath);
    let sheetRelsXml = await zip.file(sheetRelsPath)?.async('string') || createRelationshipsXml();

    const drawingState = ensureDrawingForSheet(zip, sheetPath, sheetXml, sheetRelsXml);
    sheetXml = drawingState.sheetXml;
    sheetRelsXml = drawingState.sheetRelsXml;

    let drawingXml = await zip.file(drawingState.drawingPath)?.async('string') || createDrawingXml();
    let drawingRelsXml = await zip.file(drawingState.drawingRelsPath)?.async('string') || createRelationshipsXml();
    let nextRelId = getMaxRelationshipId(drawingRelsXml) + 1;
    let nextPictureId = getMaxPictureId(drawingXml) + 1;

    for (const placement of placements) {
      const mediaInfo = mediaInfos.get(placement.imageKey);
      if (!mediaInfo) continue;

      const imageRelId = `rId${nextRelId++}`;
      const pictureId = nextPictureId++;
      const imageName = images[placement.imageKey]?.name || placement.imageKey;
      const baseAnchor = buildContainedAnchor(placement.range, geometry, mediaInfo.buffer);
      const anchor = keepAnchorWithinSinglePrintPage(
        { tl: baseAnchor.tl, br: baseAnchor.br },
        printAreaRangesBySheet.get(sheetPath) || []
      );
      if (DEBUG_IMAGE_LAYOUT) {
        console.log('[image-layout-debug]', JSON.stringify({
          sheetPath,
          imageKey: placement.imageKey,
          imageName,
          cellRef: placement.cellRef,
          range: placement.range,
          paddingPx: IMAGE_BOX_PADDING_PX,
          ...baseAnchor.debug,
          baseAnchor: {
            tl: baseAnchor.tl,
            br: baseAnchor.br,
          },
          finalAnchor: anchor,
        }));
      }
      const relationship = `<Relationship Id="${imageRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaInfo.path.split('/').pop()}"/>`;
      drawingRelsXml = appendRelationship(drawingRelsXml, relationship);
      drawingXml = appendDrawingAnchor(drawingXml, buildImageAnchorXml({
        imageName,
        imageRelId,
        pictureId,
        tl: anchor.tl,
        br: anchor.br,
      }));
    }

    drawingXml = adjustDrawingXmlToPrintPages(
      drawingXml,
      printAreaRangesBySheet.get(sheetPath) || []
    );

    zip.file(sheetPath, sheetXml);
    zip.file(sheetRelsPath, sheetRelsXml);
    zip.file(drawingState.drawingPath, drawingXml);
    zip.file(drawingState.drawingRelsPath, drawingRelsXml);
    contentTypesXml = ensureContentTypeOverride(
      contentTypesXml,
      `/${drawingState.drawingPath}`,
      'application/vnd.openxmlformats-officedocument.drawing+xml'
    );
  }

  sharedStringsXml = clearImagePlaceholderSharedStrings(sharedStringsXml, new Set(mediaInfos.keys()));

  zip.file('[Content_Types].xml', contentTypesXml);
  return sharedStringsXml;
}
