import JSZip from 'jszip';
import {
  embedImagePlaceholders,
  ImagePlaceholderInput,
  PlaceholderImages,
} from './imagePlaceholderReplacer';

export type PlaceholderPrimitive = string | number | Date | null;
export type PlaceholderObject = Record<string, unknown>;
export type PlaceholderArray = Array<PlaceholderObject>;
export type PlaceholderValue = PlaceholderPrimitive | PlaceholderArray | PlaceholderObject;

export interface PlaceholderData {
  [key: string]: PlaceholderValue;
}

export interface PlaceholderReplaceOptions {
  /**
   * Deprecated: top-level image map is no longer supported.
   */
  images?: PlaceholderImages;
}
type LegacyArrayPlaceholder = {
  placeholderKey: string;
  arrayName: string;
  fieldPath: string;
};

type SectionTablePlaceholder = {
  placeholderKey: string;
  section: string;
  table: string;
  cellPath: string;
};

type TableBlock = {
  section: string;
  table: string;
  sheetPath: string;
  sheetName: string;
  sheetIndex: number;
  startRow: number;
  endRow: number;
  blockHeight: number;
  placeholders: SectionTablePlaceholder[];
};

type SheetEntry = {
  name: string;
  index: number;
  relId: string;
  path: string;
};

type ImagePlaceholderBlock = {
  imageKey: string;
  cellRef: string;
  startRow: number;
  endRow: number;
};

type SectionImageState = {
  nextId: number;
  images: PlaceholderImages;
};
/**
 * XML特殊文字をエスケープ（W3C準拠）
 */
function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * XMLエンティティをデコード
 */
function decodeXml(text: string): string {
  return text
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, '>')
    .replace(/&lt;/g, '<')
    .replace(/&amp;/g, '&');
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value) && !(value instanceof Date);
}

/**
 * 正規表現用に文字列をエスケープ
 */
function escapeRegExp(string: string): string {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * XML内のプレースホルダーを検出
 */
function detectPlaceholdersInXml(xml: string): Array<{
  placeholder: string;
  key: string;
  count: number;
}> {
  const placeholderMap = new Map<string, { placeholder: string; key: string; count: number }>();
  const regex = /\{\{([^}]+)\}\}/g;
  const sharedStrings = extractSharedStrings(xml);

  sharedStrings.forEach((text) => {
    regex.lastIndex = 0;
    let match: RegExpExecArray | null;
    while ((match = regex.exec(text)) !== null) {
      const placeholder = match[0];
      const key = match[1];

      const existing = placeholderMap.get(placeholder);
      if (existing) {
        existing.count += 1;
      } else {
        placeholderMap.set(placeholder, {
          placeholder,
          key,
          count: 1,
        });
      }
    }
  });

  return Array.from(placeholderMap.values());
}

function parseLegacyArrayPlaceholderKey(key: string): LegacyArrayPlaceholder | null {
  if (!key.startsWith('#') || key.startsWith('##')) return null;
  const cleanName = key.substring(1);
  const parts = cleanName.split('.');
  if (parts.length < 2) return null;
  return {
    placeholderKey: key,
    arrayName: parts[0],
    fieldPath: parts.slice(1).join('.'),
  };
}

function parseSectionTablePlaceholderKey(key: string): SectionTablePlaceholder | null {
  if (!key.startsWith('##')) return null;
  const cleanName = key.substring(2);
  const parts = cleanName.split('.');
  if (parts.length < 3) return null;
  return {
    placeholderKey: key,
    section: parts[0],
    table: parts[1],
    cellPath: parts.slice(2).join('.'),
  };
}

function extractSharedStrings(sharedStringsXml: string): string[] {
  const strings: string[] = [];
  const siRegex = /<si>(.*?)<\/si>/gs;
  let match: RegExpExecArray | null;

  while ((match = siRegex.exec(sharedStringsXml)) !== null) {
    const siContent = match[1];
    const tRegex = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
    const chunks: string[] = [];
    let tMatch: RegExpExecArray | null;
    while ((tMatch = tRegex.exec(siContent)) !== null) {
      chunks.push(tMatch[1]);
    }
    if (chunks.length > 0) {
      strings.push(decodeXml(chunks.join('')));
    }
  }

  return strings;
}

function extractSiText(siContent: string): string {
  const tRegex = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
  const chunks: string[] = [];
  let tMatch: RegExpExecArray | null;
  while ((tMatch = tRegex.exec(siContent)) !== null) {
    chunks.push(tMatch[1]);
  }
  return decodeXml(chunks.join(''));
}

function extractFirstRunProperties(siContent: string): string | null {
  const firstRunMatch = siContent.match(/<r>([\s\S]*?)<\/r>/);
  if (!firstRunMatch) return null;
  const rPrMatch = firstRunMatch[1].match(/<rPr>[\s\S]*?<\/rPr>/);
  return rPrMatch ? rPrMatch[0] : null;
}

function buildTextNode(text: string): string {
  const escapedText = escapeXml(text);
  const preserveSpace = /^\s|\s$/.test(text);
  return preserveSpace
    ? `<t xml:space="preserve">${escapedText}</t>`
    : `<t>${escapedText}</t>`;
}

function formatDateValue(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

function stringifyPrimitiveValue(value: unknown): string {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return formatDateValue(value);
  return String(value);
}

function isImagePlaceholderInput(value: unknown): value is ImagePlaceholderInput {
  if (!isPlainObject(value)) return false;
  if (typeof value.base64 !== 'string' || value.base64.trim() === '') return false;
  if (typeof value.contentType !== 'string') return false;
  return value.contentType.toLowerCase().startsWith('image/');
}

function buildGeneratedSectionImageToken(imageId: number): string {
  return `__section_image_${imageId}`;
}

function buildImagePlaceholderText(imageKey: string): string {
  return `{{%${imageKey}}}`;
}

function replacePrimitivePlaceholdersWithFirstRunStyle(
  sharedStringsXml: string,
  replacements: Map<string, string>
): string {
  if (replacements.size === 0) return sharedStringsXml;

  const siRegex = /<si>([\s\S]*?)<\/si>/g;
  return sharedStringsXml.replace(siRegex, (siWhole: string, siContent: string) => {
    const originalText = extractSiText(siContent);
    if (!originalText.includes('{{')) return siWhole;

    let replacedText = originalText;
    replacements.forEach((value, key) => {
      const tokenRegex = new RegExp(escapeRegExp(`{{${key}}}`), 'g');
      replacedText = replacedText.replace(tokenRegex, value);
    });

    if (replacedText === originalText) return siWhole;

    const firstRunProps = extractFirstRunProperties(siContent);
    const textNode = buildTextNode(replacedText);

    if (firstRunProps) {
      return `<si><r>${firstRunProps}${textNode}</r></si>`;
    }

    return `<si>${textNode}</si>`;
  });
}

function collectPlaceholderIndices(
  sharedStrings: string[],
  placeholderKeys: string[]
): Map<string, Set<number>> {
  const placeholderToIndices = new Map<string, Set<number>>();

  placeholderKeys.forEach((key) => {
    const token = `{{${key}}}`;
    const indices = new Set<number>();

    sharedStrings.forEach((str, idx) => {
      if (str.includes(token)) indices.add(idx);
    });

    if (indices.size > 0) {
      placeholderToIndices.set(key, indices);
    }
  });

  return placeholderToIndices;
}

function findRowsBySharedStringIndices(sheetXml: string, indices: Set<number>): Set<number> {
  const rows = new Set<number>();
  if (indices.size === 0) return rows;

  const cellRegex = /<c\b[^>]*\br="[A-Z]+(\d+)"[^>]*>([\s\S]*?)<\/c>/g;
  let match: RegExpExecArray | null;

  while ((match = cellRegex.exec(sheetXml)) !== null) {
    const row = Number(match[1]);
    const cellBody = match[2];
    const valueMatch = cellBody.match(/<v>(\d+)<\/v>/);
    if (!valueMatch) continue;

    const sharedStringIndex = Number(valueMatch[1]);
    if (indices.has(sharedStringIndex)) {
      rows.add(row);
    }
  }

  return rows;
}

function collectImagePlaceholderIndices(
  sharedStrings: string[],
  imageKeys: Set<string>
): Map<number, string> {
  const result = new Map<number, string>();

  sharedStrings.forEach((text, index) => {
    const match = text.match(/^\{\{%([^}]+)\}\}$/);
    if (!match) return;

    const imageKey = match[1];
    if (imageKeys.has(imageKey)) {
      result.set(index, imageKey);
    }
  });

  return result;
}

function findImagePlaceholderBlocks(
  sheetXml: string,
  sharedStrings: string[],
  imageKeys: Set<string>
): ImagePlaceholderBlock[] {
  const placeholderIndices = collectImagePlaceholderIndices(sharedStrings, imageKeys);
  if (placeholderIndices.size === 0) return [];

  const mergeSpans = parseMergeSpans(sheetXml);
  const blocks: ImagePlaceholderBlock[] = [];
  const cellRegex = /<c\b[^>]*\br="([A-Z]+\d+)"[^>]*>([\s\S]*?)<\/c>/g;
  let match: RegExpExecArray | null;

  while ((match = cellRegex.exec(sheetXml)) !== null) {
    const cellRef = match[1];
    const valueMatch = match[2].match(/<v>(\d+)<\/v>/);
    if (!valueMatch) continue;

    const sharedStringIndex = Number(valueMatch[1]);
    const imageKey = placeholderIndices.get(sharedStringIndex);
    if (!imageKey) continue;

    const cellMatch = cellRef.match(/^([A-Z]+)(\d+)$/);
    if (!cellMatch) continue;

    const colIndex = columnNameToIndex(cellMatch[1]);
    const rowIndex = Number(cellMatch[2]);
    const mergeSpan = mergeSpans.find((span) =>
      rowIndex >= span.startRow &&
      rowIndex <= span.endRow &&
      colIndex >= span.startColIndex &&
      colIndex <= span.endColIndex
    );

    blocks.push({
      imageKey,
      cellRef,
      startRow: mergeSpan ? mergeSpan.startRow : rowIndex,
      endRow: mergeSpan ? mergeSpan.endRow : rowIndex,
    });
  }

  return blocks.sort((a, b) => a.startRow - b.startRow);
}

function extractRowXml(sheetXml: string, rowNumber: number): string | null {
  const rowRegex = new RegExp(`<row[^>]*r="${rowNumber}"[^>]*>.*?<\\/row>`, 's');
  const match = sheetXml.match(rowRegex);
  return match ? match[0] : null;
}

function updateRowNumber(rowXml: string, newRowNumber: number): string {
  let updated = rowXml.replace(/(<row[^>]*r=")(\d+)(")/g, `$1${newRowNumber}$3`);
  updated = updated.replace(/(<c r="[A-Z]+)(\d+)(")/g, `$1${newRowNumber}$3`);
  return updated;
}

function shiftA1ReferencesInFormula(formula: string, rowDelta: number): string {
  if (rowDelta === 0) return formula;

  const cellRefRegex = /((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]{1,3})(\$?)(\d+)/g;
  return formula.replace(
    cellRefRegex,
    (_whole, sheetPrefix: string | undefined, colAbs: string, col: string, rowAbs: string, rowNum: string) => {
      if (rowAbs === '$') {
        return `${sheetPrefix || ''}${colAbs}${col}${rowAbs}${rowNum}`;
      }

      const shifted = Math.max(1, Number(rowNum) + rowDelta);
      return `${sheetPrefix || ''}${colAbs}${col}${shifted}`;
    }
  );
}

function shiftFormulaRowsInRowXml(rowXml: string, rowDelta: number): string {
  if (rowDelta === 0) return rowXml;
  return rowXml.replace(/(<f\b[^>]*>)([\s\S]*?)(<\/f>)/g, (_whole, openTag: string, formula: string, closeTag: string) => {
    return `${openTag}${shiftA1ReferencesInFormula(formula, rowDelta)}${closeTag}`;
  });
}

function copyRowXmlWithShiftedFormulas(templateRowXml: string, sourceRowNumber: number, targetRowNumber: number): string {
  const rowDelta = targetRowNumber - sourceRowNumber;
  const renumbered = updateRowNumber(templateRowXml, targetRowNumber);
  return shiftFormulaRowsInRowXml(renumbered, rowDelta);
}

function shiftRowsDown(sheetXml: string, fromRow: number, shiftAmount: number): string {
  const rowRegex = /<row[^>]*r="(\d+)"[^>]*>.*?<\/row>/gs;
  let result = sheetXml;
  const matches: Array<{ rowNum: number; xml: string }> = [];

  let match: RegExpExecArray | null;
  while ((match = rowRegex.exec(sheetXml)) !== null) {
    const rowNum = Number(match[1]);
    if (rowNum >= fromRow) {
      matches.push({ rowNum, xml: match[0] });
    }
  }

  matches.sort((a, b) => b.rowNum - a.rowNum);
  matches.forEach(({ rowNum, xml }) => {
    const updatedXml = updateRowNumber(xml, rowNum + shiftAmount);
    result = result.replace(xml, updatedXml);
  });

  return result;
}

function shiftMergeCellsDown(sheetXml: string, fromRow: number, shiftAmount: number): string {
  const mergeCellsRegex = /<mergeCells[^>]*>(.*?)<\/mergeCells>/s;
  const mergeCellsMatch = sheetXml.match(mergeCellsRegex);
  if (!mergeCellsMatch) return sheetXml;

  const mergeCellsContent = mergeCellsMatch[1];
  const mergeCellRegex = /<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/g;
  const updates: Array<{ oldRef: string; newRef: string; row: number }> = [];
  let match: RegExpExecArray | null;

  while ((match = mergeCellRegex.exec(mergeCellsContent)) !== null) {
    const startCol = match[1];
    const startRow = Number(match[2]);
    const endCol = match[3];
    const endRow = Number(match[4]);

    if (startRow >= fromRow || endRow >= fromRow) {
      const newStartRow = startRow >= fromRow ? startRow + shiftAmount : startRow;
      const newEndRow = endRow >= fromRow ? endRow + shiftAmount : endRow;
      updates.push({
        oldRef: `${startCol}${startRow}:${endCol}${endRow}`,
        newRef: `${startCol}${newStartRow}:${endCol}${newEndRow}`,
        row: startRow,
      });
    }
  }

  updates.sort((a, b) => b.row - a.row);
  let result = sheetXml;
  updates.forEach(({ oldRef, newRef }) => {
    result = result.replace(`<mergeCell ref="${oldRef}"/>`, `<mergeCell ref="${newRef}"/>`);
  });

  return result;
}

function shiftRowBreaksDown(sheetXml: string, fromRow: number, shiftAmount: number): string {
  if (shiftAmount <= 0) return sheetXml;

  return sheetXml.replace(/<rowBreaks\b([^>]*)>([\s\S]*?)<\/rowBreaks>/g, (_full, attrs: string, content: string) => {
    const updatedContent = content.replace(
      /<(brk|rowBreak)\b([^>]*)\/>/g,
      (_tagFull, tagName: string, tagAttrs: string) => {
        const idMatch = tagAttrs.match(/\bid="(\d+)"/);
        if (!idMatch) return `<${tagName}${tagAttrs}/>`;

        const id = Number(idMatch[1]);
        if (id < fromRow) return `<${tagName}${tagAttrs}/>`;

        const updatedAttrs = tagAttrs.replace(/\bid="\d+"/, `id="${id + shiftAmount}"`);
        return `<${tagName}${updatedAttrs}/>`;
      }
    );

    return `<rowBreaks${attrs}>${updatedContent}</rowBreaks>`;
  });
}

function insertMergeCellsForTemplateBlock(
  sheetXml: string,
  templateStartRow: number,
  templateEndRow: number,
  blockHeight: number,
  repeatCount: number
): string {
  if (repeatCount <= 0) return sheetXml;

  const mergeCellsRegex = /<mergeCells[^>]*>(.*?)<\/mergeCells>/s;
  const mergeCellsMatch = sheetXml.match(mergeCellsRegex);
  if (!mergeCellsMatch) return sheetXml;

  const mergeCellsContent = mergeCellsMatch[1];
  const mergeCellRegex = /<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/g;
  let match: RegExpExecArray | null;

  const templateMerges: Array<{ startCol: string; startRow: number; endCol: string; endRow: number }> = [];
  while ((match = mergeCellRegex.exec(mergeCellsContent)) !== null) {
    const startRow = Number(match[2]);
    const endRow = Number(match[4]);
    if (startRow >= templateStartRow && endRow <= templateEndRow) {
      templateMerges.push({
        startCol: match[1],
        startRow,
        endCol: match[3],
        endRow,
      });
    }
  }

  if (templateMerges.length === 0) return sheetXml;

  const newMergeCells: string[] = [];
  for (let repeat = 1; repeat <= repeatCount; repeat++) {
    const delta = repeat * blockHeight;
    templateMerges.forEach((m) => {
      newMergeCells.push(
        `<mergeCell ref="${m.startCol}${m.startRow + delta}:${m.endCol}${m.endRow + delta}"/>`
      );
    });
  }

  let result = sheetXml.replace('</mergeCells>', `${newMergeCells.join('')}</mergeCells>`);

  const countMatch = result.match(/<mergeCells count="(\d+)"/);
  if (countMatch) {
    const currentCount = Number(countMatch[1]);
    result = result.replace(/(<mergeCells count=")(\d+)(")/, `$1${currentCount + newMergeCells.length}$3`);
  }

  return result;
}

function addSharedString(
  sharedStringsXml: string,
  newString: string,
  firstRunProps?: string | null
): { updatedXml: string; newIndex: number } {
  const countMatch = sharedStringsXml.match(/<sst[^>]*count="(\d+)"/);
  const uniqueCountMatch = sharedStringsXml.match(/<sst[^>]*uniqueCount="(\d+)"/);
  const currentCount = countMatch ? Number(countMatch[1]) : 0;
  const currentUniqueCount = uniqueCountMatch ? Number(uniqueCountMatch[1]) : 0;
  const newIndex = currentUniqueCount;
  const textNode = buildTextNode(newString);
  const siNode = firstRunProps
    ? `<si><r>${firstRunProps}${textNode}</r></si>`
    : `<si>${textNode}</si>`;

  let updatedXml = sharedStringsXml.replace(/count="\d+"/, `count="${currentCount + 1}"`);
  updatedXml = updatedXml.replace(/uniqueCount="\d+"/, `uniqueCount="${currentUniqueCount + 1}"`);
  updatedXml = updatedXml.replace('</sst>', `${siNode}</sst>`);

  return { updatedXml, newIndex };
}

function collectFirstRunPropertiesBySharedStringIndex(
  sharedStringsXml: string,
  indices: Set<number>
): Map<number, string | null> {
  const result = new Map<number, string | null>();
  if (indices.size === 0) return result;

  const siRegex = /<si>([\s\S]*?)<\/si>/g;
  let index = 0;
  let match: RegExpExecArray | null;
  while ((match = siRegex.exec(sharedStringsXml)) !== null) {
    if (indices.has(index)) {
      result.set(index, extractFirstRunProperties(match[1]));
      if (result.size === indices.size) break;
    }
    index += 1;
  }

  return result;
}

function parseFirstPrintArea(printAreaValue: string): {
  sheetPrefix: string;
  startCol: string;
  startRow: number;
  endCol: string;
  endRow: number;
} | null {
  const first = printAreaValue.split(',')[0];
  const bangIndex = first.indexOf('!');
  if (bangIndex === -1) return null;

  const sheetPrefix = first.slice(0, bangIndex);
  const rangePart = first.slice(bangIndex + 1);

  const m = rangePart.match(/\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/);
  if (!m) return null;

  return {
    sheetPrefix,
    startCol: m[1],
    startRow: Number(m[2]),
    endCol: m[3],
    endRow: Number(m[4]),
  };
}

type PrintAreaRange = {
  sheetPrefix: string;
  startCol: string;
  startRow: number;
  endCol: string;
  endRow: number;
};

type CellStyleInfo = {
  fontSize: number;
  wrapText: boolean;
};

type StyleCatalog = {
  styles: Map<number, CellStyleInfo>;
  defaultFontSize: number;
};

type MergeSpan = {
  startColIndex: number;
  endColIndex: number;
  startRow: number;
  endRow: number;
};

type PageLayoutInfo = {
  defaultRowHeight: number;
  defaultColWidth: number;
  pageHeightPoints: number;
  pageWidthPoints: number;
  printableHeightPoints: number;
  printableWidthPoints: number;
  effectiveScale: number;
  pageHeightCapacityPoints: number;
  columnWidthByIndex: Map<number, number>;
};

function parsePrintAreaRanges(printAreaValue: string): PrintAreaRange[] {
  return printAreaValue
    .split(',')
    .map((part) => part.trim())
    .map((part): PrintAreaRange | null => {
      const bangIndex = part.indexOf('!');
      if (bangIndex === -1) return null;

      const sheetPrefix = part.slice(0, bangIndex);
      const rangePart = part.slice(bangIndex + 1);
      const m = rangePart.match(/\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/);
      if (!m) return null;

      return {
        sheetPrefix,
        startCol: m[1],
        startRow: Number(m[2]),
        endCol: m[3],
        endRow: Number(m[4]),
      };
    })
    .filter((r): r is PrintAreaRange => r !== null);
}

function parseDefaultRowHeight(sheetXml: string): number {
  const m = sheetXml.match(/<sheetFormatPr\b[^>]*\bdefaultRowHeight="([\d.]+)"/);
  return m ? Number(m[1]) : 15;
}

function parseDefaultColWidth(sheetXml: string): number {
  const m = sheetXml.match(/<sheetFormatPr\b[^>]*\bdefaultColWidth="([\d.]+)"/);
  return m ? Number(m[1]) : 8.43;
}

function parseColumnWidthMap(sheetXml: string): Map<number, number> {
  const result = new Map<number, number>();
  const colRegex = /<col\b([^>]*)\/>/g;
  let match: RegExpExecArray | null;

  while ((match = colRegex.exec(sheetXml)) !== null) {
    const attrs = match[1];
    const minMatch = attrs.match(/\bmin="(\d+)"/);
    const maxMatch = attrs.match(/\bmax="(\d+)"/);
    const widthMatch = attrs.match(/\bwidth="([\d.]+)"/);
    if (!minMatch || !maxMatch || !widthMatch) continue;

    const min = Number(minMatch[1]);
    const max = Number(maxMatch[1]);
    const width = Number(widthMatch[1]);
    for (let index = min; index <= max; index++) {
      result.set(index, width);
    }
  }

  return result;
}

function columnNameToIndex(columnName: string): number {
  let index = 0;
  for (const ch of columnName) {
    index = index * 26 + (ch.charCodeAt(0) - 64);
  }
  return index;
}

function columnWidthToPoints(width: number): number {
  return width * 5.25;
}

function getColumnWidthPoints(
  columnIndex: number,
  columnWidthByIndex: Map<number, number>,
  defaultColWidth: number
): number {
  return columnWidthToPoints(columnWidthByIndex.get(columnIndex) ?? defaultColWidth);
}

function getPaperSizePoints(paperSize: number | null): { width: number; height: number } {
  switch (paperSize) {
    case 1:
      return { width: 612, height: 792 };
    case 5:
      return { width: 612, height: 1008 };
    case 8:
      return { width: 841.89, height: 1190.55 };
    case 9:
      return { width: 595.28, height: 841.89 };
    default:
      return { width: 595.28, height: 841.89 };
  }
}

function parsePageLayoutInfo(sheetXml: string, range: PrintAreaRange): PageLayoutInfo {
  const defaultRowHeight = parseDefaultRowHeight(sheetXml);
  const defaultColWidth = parseDefaultColWidth(sheetXml);
  const columnWidthByIndex = parseColumnWidthMap(sheetXml);
  const pageSetupTag = sheetXml.match(/<pageSetup\b([^>]*)\/>/)?.[1] || '';
  const pageMarginsTag = sheetXml.match(/<pageMargins\b([^>]*)\/>/)?.[1] || '';

  const paperSize = pageSetupTag.match(/\bpaperSize="(\d+)"/)?.[1];
  const orientation = pageSetupTag.match(/\borientation="([^"]+)"/)?.[1] || 'portrait';
  const scaleAttr = pageSetupTag.match(/\bscale="([\d.]+)"/)?.[1];
  const fitToWidthAttr = pageSetupTag.match(/\bfitToWidth="(\d+)"/)?.[1];
  const fitToHeightAttr = pageSetupTag.match(/\bfitToHeight="(\d+)"/)?.[1];

  const basePaper = getPaperSizePoints(paperSize ? Number(paperSize) : null);
  const isLandscape = orientation === 'landscape';
  const pageWidthPoints = isLandscape ? basePaper.height : basePaper.width;
  const pageHeightPoints = isLandscape ? basePaper.width : basePaper.height;

  const left = Number(pageMarginsTag.match(/\bleft="([\d.]+)"/)?.[1] || '0.7') * 72;
  const right = Number(pageMarginsTag.match(/\bright="([\d.]+)"/)?.[1] || '0.7') * 72;
  const top = Number(pageMarginsTag.match(/\btop="([\d.]+)"/)?.[1] || '0.75') * 72;
  const bottom = Number(pageMarginsTag.match(/\bbottom="([\d.]+)"/)?.[1] || '0.75') * 72;

  // ヘッダー/フッターはtop/bottomマージンの内側に描画されるため、本文領域の計算には含めない
  const printableWidthPoints = Math.max(1, pageWidthPoints - left - right);
  const printableHeightPoints = Math.max(1, pageHeightPoints - top - bottom);

  const startColIndex = columnNameToIndex(range.startCol);
  const endColIndex = columnNameToIndex(range.endCol);
  let contentWidthPoints = 0;
  for (let colIndex = startColIndex; colIndex <= endColIndex; colIndex++) {
    contentWidthPoints += getColumnWidthPoints(colIndex, columnWidthByIndex, defaultColWidth);
  }

  const fitModeEnabled =
    (fitToWidthAttr && Number(fitToWidthAttr) > 0) ||
    (fitToHeightAttr && Number(fitToHeightAttr) > 0);

  const explicitScale = scaleAttr ? Number(scaleAttr) / 100 : 1;
  let effectiveScale = fitModeEnabled ? 1 : explicitScale > 0 ? explicitScale : 1;
  if (fitToWidthAttr && Number(fitToWidthAttr) > 0 && contentWidthPoints > 0) {
    const fitWidthScale = Math.min(1, printableWidthPoints / contentWidthPoints);
    effectiveScale = fitWidthScale;
  }

  return {
    defaultRowHeight,
    defaultColWidth,
    pageHeightPoints,
    pageWidthPoints,
    printableHeightPoints,
    printableWidthPoints,
    effectiveScale,
    pageHeightCapacityPoints: printableHeightPoints / Math.max(effectiveScale, 0.01),
    columnWidthByIndex,
  };
}

function parseStyleCatalog(stylesXml: string): StyleCatalog {
  const fontSizes: number[] = [];
  const fontsBlock = stylesXml.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/);
  if (fontsBlock) {
    const fontRegex = /<font\b[^>]*>([\s\S]*?)<\/font>/g;
    let fontMatch: RegExpExecArray | null;
    while ((fontMatch = fontRegex.exec(fontsBlock[1])) !== null) {
      const sizeMatch = fontMatch[1].match(/<sz\b[^>]*val="([\d.]+)"/);
      fontSizes.push(sizeMatch ? Number(sizeMatch[1]) : 11);
    }
  }

  const defaultFontSize = fontSizes[0] || 11;
  const styles = new Map<number, CellStyleInfo>();
  const cellXfsBlock = stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
  if (cellXfsBlock) {
    const xfRegex = /<xf\b([^>]*?)(?:\/>|>([\s\S]*?)<\/xf>)/g;
    let xfMatch: RegExpExecArray | null;
    let styleIndex = 0;
    while ((xfMatch = xfRegex.exec(cellXfsBlock[1])) !== null) {
      const attrs = xfMatch[1];
      const body = xfMatch[2] || '';
      const fontId = Number(attrs.match(/\bfontId="(\d+)"/)?.[1] || '0');
      const alignmentTag = body.match(/<alignment\b([^>]*)\/>/)?.[1] || '';
      const wrapText = /\bwrapText="1"/.test(attrs) || /\bwrapText="1"/.test(alignmentTag);
      styles.set(styleIndex, {
        fontSize: fontSizes[fontId] || defaultFontSize,
        wrapText,
      });
      styleIndex += 1;
    }
  }

  return { styles, defaultFontSize };
}

function parseMergeSpans(sheetXml: string): MergeSpan[] {
  const mergeSpans: MergeSpan[] = [];
  const mergeCellRegex = /<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/g;
  let match: RegExpExecArray | null;

  while ((match = mergeCellRegex.exec(sheetXml)) !== null) {
    mergeSpans.push({
      startColIndex: columnNameToIndex(match[1]),
      endColIndex: columnNameToIndex(match[3]),
      startRow: Number(match[2]),
      endRow: Number(match[4]),
    });
  }

  return mergeSpans;
}

function getMergedCellWidthPoints(
  cellRef: string,
  layoutInfo: PageLayoutInfo,
  mergeSpans: MergeSpan[]
): number {
  const cellMatch = cellRef.match(/^([A-Z]+)(\d+)$/);
  if (!cellMatch) return 0;

  const colIndex = columnNameToIndex(cellMatch[1]);
  const rowIndex = Number(cellMatch[2]);
  const mergeSpan = mergeSpans.find((span) =>
    span.startColIndex === colIndex &&
    span.startRow === rowIndex
  );

  const start = mergeSpan ? mergeSpan.startColIndex : colIndex;
  const end = mergeSpan ? mergeSpan.endColIndex : colIndex;
  let widthPoints = 0;
  for (let index = start; index <= end; index++) {
    widthPoints += getColumnWidthPoints(index, layoutInfo.columnWidthByIndex, layoutInfo.defaultColWidth);
  }
  return widthPoints;
}

function estimateTextWidthPoints(text: string, fontSize: number): number {
  let widthPoints = 0;
  for (const ch of text) {
    if (ch === '\n') continue;
    if (/\s/.test(ch)) {
      widthPoints += fontSize * 0.35;
      continue;
    }
    if (/[\u0000-\u00ff]/.test(ch)) {
      widthPoints += fontSize * 0.55;
      continue;
    }
    widthPoints += fontSize;
  }
  return widthPoints;
}

function estimateWrappedTextHeight(text: string, fontSize: number, widthPoints: number, wrapText: boolean): number {
  const paragraphs = text.split(/\r?\n/);
  let lines = 0;

  for (const paragraph of paragraphs) {
    if (paragraph.length === 0) {
      lines += 1;
      continue;
    }

    if (!wrapText || widthPoints <= 0) {
      lines += 1;
      continue;
    }

    const paragraphWidth = estimateTextWidthPoints(paragraph, fontSize);
    lines += Math.max(1, Math.ceil(paragraphWidth / widthPoints));
  }

  return Math.max(0, lines) * (fontSize * 1.25) + 4;
}

function getCellTextFromXml(cellXml: string, sharedStrings: string[]): string {
  const typeMatch = cellXml.match(/\bt="([^"]+)"/);
  const cellType = typeMatch ? typeMatch[1] : '';

  if (cellType === 's') {
    const valueMatch = cellXml.match(/<v>(\d+)<\/v>/);
    if (!valueMatch) return '';
    return sharedStrings[Number(valueMatch[1])] || '';
  }

  if (cellType === 'inlineStr') {
    const inlineText = cellXml.match(/<t\b[^>]*>([\s\S]*?)<\/t>/);
    return inlineText ? decodeXml(inlineText[1]) : '';
  }

  const valueMatch = cellXml.match(/<v>([\s\S]*?)<\/v>/);
  return valueMatch ? decodeXml(valueMatch[1]) : '';
}

function estimateRowRenderedHeight(
  rowXml: string,
  sharedStrings: string[],
  styleCatalog: StyleCatalog,
  layoutInfo: PageLayoutInfo,
  mergeSpans: MergeSpan[]
): number {
  const rowTagMatch = rowXml.match(/<row\b[^>]*>/);
  const baseHeight = rowTagMatch?.[0].match(/\bht="([\d.]+)"/)?.[1];
  let maxHeight = baseHeight ? Number(baseHeight) : layoutInfo.defaultRowHeight;

  const cellRegex = /<c\b[^>]*\br="([A-Z]+\d+)"([^>]*)>([\s\S]*?)<\/c>/g;
  let match: RegExpExecArray | null;
  while ((match = cellRegex.exec(rowXml)) !== null) {
    const cellRef = match[1];
    const cellAttrs = match[2];
    const cellBody = match[3];
    const styleIndex = Number(cellAttrs.match(/\bs="(\d+)"/)?.[1] || '0');
    const style = styleCatalog.styles.get(styleIndex) || {
      fontSize: styleCatalog.defaultFontSize,
      wrapText: false,
    };
    if (!style.wrapText) continue;
    const text = getCellTextFromXml(`${match[0]}`, sharedStrings);
    if (!text) continue;

    const widthPoints = getMergedCellWidthPoints(cellRef, layoutInfo, mergeSpans);
    const estimatedHeight = estimateWrappedTextHeight(text, style.fontSize, widthPoints, style.wrapText);
    maxHeight = Math.max(maxHeight, estimatedHeight);
  }

  return maxHeight;
}

function buildEffectiveRowHeightMap(
  sheetXml: string,
  sharedStrings: string[],
  styleCatalog: StyleCatalog,
  layoutInfo: PageLayoutInfo
): Map<number, number> {
  const result = new Map<number, number>();
  const rowRegex = /<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g;
  const mergeSpans = parseMergeSpans(sheetXml);
  let match: RegExpExecArray | null;

  while ((match = rowRegex.exec(sheetXml)) !== null) {
    const rowNumber = Number(match[1]);
    result.set(
      rowNumber,
      estimateRowRenderedHeight(match[0], sharedStrings, styleCatalog, layoutInfo, mergeSpans)
    );
  }

  return result;
}

type KeepTogetherUnit = {
  startRow: number;
  endRow: number;
};

function mergeKeepTogetherUnits(units: KeepTogetherUnit[]): KeepTogetherUnit[] {
  const sorted = [...units].sort((a, b) => a.startRow - b.startRow);
  const merged: KeepTogetherUnit[] = [];

  for (const unit of sorted) {
    const last = merged[merged.length - 1];
    if (last && unit.startRow <= last.endRow) {
      last.endRow = Math.max(last.endRow, unit.endRow);
    } else {
      merged.push({ ...unit });
    }
  }

  return merged;
}

/**
 * 印刷範囲内の手動改ページ位置を計算する。
 * 返り値はOOXMLのrowBreaks仕様（id=N で「N行目の直後で改ページ」）に合わせた行番号。
 * - units: レコードブロックなどページをまたいではいけない行塊
 * - sections: セクション全体の行範囲。前のセクションがページをまたいだ場合、次のセクションは新しいページから開始する
 * - forcedBreakIds: テンプレート由来の既存手動改ページ
 */
function computeManualRowBreaks(params: {
  startRow: number;
  endRow: number;
  pageCapacity: number;
  rowHeights: Map<number, number>;
  defaultRowHeight: number;
  units: KeepTogetherUnit[];
  sections: KeepTogetherUnit[];
  forcedBreakIds: Set<number>;
}): number[] {
  const EPS = 0.000001;
  if (params.pageCapacity <= 0 || params.endRow < params.startRow) return [];

  const unitByRow = new Map<number, KeepTogetherUnit>();
  mergeKeepTogetherUnits(params.units).forEach((unit) => {
    for (let row = unit.startRow; row <= unit.endRow; row++) {
      unitByRow.set(row, unit);
    }
  });

  const sections = [...params.sections].sort((a, b) => a.startRow - b.startRow);
  let sectionIndex = 0;
  let activeSection: KeepTogetherUnit | null = null;
  let activeSectionCrossed = false;
  let prevSectionCrossed = false;

  const breaks: number[] = [];
  let pageStart = params.startRow;
  let used = 0;

  const heightOf = (row: number): number => {
    const height = params.rowHeights.get(row) ?? params.defaultRowHeight;
    return height > 0 ? height : 1;
  };
  const markSectionCrossing = (breakRow: number): void => {
    if (activeSection && breakRow >= activeSection.startRow && breakRow < activeSection.endRow) {
      activeSectionCrossed = true;
    }
  };

  for (let row = params.startRow; row <= params.endRow; row++) {
    if (activeSection && row > activeSection.endRow) {
      prevSectionCrossed = activeSectionCrossed;
      activeSection = null;
      activeSectionCrossed = false;
    }

    if (row > params.startRow && params.forcedBreakIds.has(row - 1)) {
      markSectionCrossing(row - 1);
      pageStart = row;
      used = 0;
    }

    if (sectionIndex < sections.length && row === sections[sectionIndex].startRow) {
      if (prevSectionCrossed && row !== pageStart) {
        breaks.push(row - 1);
        pageStart = row;
        used = 0;
      }
      activeSection = sections[sectionIndex];
      activeSectionCrossed = false;
      sectionIndex += 1;
    }

    const rowHeight = heightOf(row);
    if (used > 0 && used + rowHeight > params.pageCapacity + EPS) {
      const unit = unitByRow.get(row);
      if (unit && unit.startRow <= pageStart) {
        // ページ先頭から始まる塊がページ容量を超える場合は分割を諦めて流し込む
        used += rowHeight;
      } else if (unit && unit.startRow < row) {
        // 塊の途中で溢れた場合は塊全体を次ページへ送る
        breaks.push(unit.startRow - 1);
        markSectionCrossing(unit.startRow - 1);
        pageStart = unit.startRow;
        used = 0;
        for (let unitRow = unit.startRow; unitRow <= row; unitRow++) {
          used += heightOf(unitRow);
        }
      } else {
        breaks.push(row - 1);
        markSectionCrossing(row - 1);
        pageStart = row;
        used = rowHeight;
      }
    } else {
      used += rowHeight;
    }
  }

  return breaks;
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

function upsertRowBreaks(sheetXml: string, breakIds: number[]): string {
  const ids = Array.from(new Set(breakIds))
    .filter((id) => id >= 1)
    .sort((a, b) => a - b);
  if (ids.length === 0) return sheetXml;

  const content = ids.map((id) => `<brk id="${id}" max="16383" man="1"/>`).join('');
  const element = `<rowBreaks count="${ids.length}" manualBreakCount="${ids.length}">${content}</rowBreaks>`;

  if (/<rowBreaks\b[^>]*>[\s\S]*?<\/rowBreaks>/.test(sheetXml)) {
    return sheetXml.replace(/<rowBreaks\b[^>]*>[\s\S]*?<\/rowBreaks>/, element);
  }
  if (/<rowBreaks\b[^>]*\/>/.test(sheetXml)) {
    return sheetXml.replace(/<rowBreaks\b[^>]*\/>/, element);
  }

  // CT_Worksheetの要素順（... pageMargins, pageSetup, headerFooter, rowBreaks ...）に従って挿入
  const insertAfterPatterns = [
    /<\/headerFooter>/,
    /<headerFooter\b[^>]*\/>/,
    /<pageSetup\b[^>]*\/>/,
    /<pageMargins\b[^>]*\/>/,
  ];
  for (const pattern of insertAfterPatterns) {
    const match = sheetXml.match(pattern);
    if (match) {
      return sheetXml.replace(match[0], `${match[0]}${element}`);
    }
  }

  if (/<colBreaks\b/.test(sheetXml)) {
    return sheetXml.replace(/<colBreaks\b/, `${element}<colBreaks`);
  }
  if (/<drawing\b/.test(sheetXml)) {
    return sheetXml.replace(/<drawing\b/, `${element}<drawing`);
  }
  return sheetXml.replace('</worksheet>', `${element}</worksheet>`);
}

function disableFitToHeight(sheetXml: string): string {
  return sheetXml.replace(/<pageSetup\b[^>]*\/>/, (tag) => {
    const fitToHeightMatch = tag.match(/\bfitToHeight="(\d+)"/);
    if (!fitToHeightMatch || Number(fitToHeightMatch[1]) === 0) return tag;
    // 手動改ページと縦方向のページ収め縮小は両立しないため無効化する
    return tag.replace(/\bfitToHeight="\d+"/, 'fitToHeight="0"');
  });
}

function getFirstPrintAreaForSheet(
  workbookXml: string,
  targetSheetName: string,
  targetSheetIndex: number
): PrintAreaRange | null {
  const ranges = getPrintAreaRangesForSheet(workbookXml, targetSheetName, targetSheetIndex);
  return ranges[0] || null;
}

function getPrintAreaRangesForSheet(
  workbookXml: string,
  targetSheetName: string,
  targetSheetIndex: number
): PrintAreaRange[] {
  const tags = workbookXml.match(/<definedName\b[\s\S]*?<\/definedName>/g) || [];
  for (const tag of tags) {
    if (!/name="_xlnm\.Print_Area"/.test(tag)) continue;

    const openTagMatch = tag.match(/<definedName\b[^>]*>/);
    const valueMatch = tag.match(/>([^<]*)<\/definedName>/);
    if (!openTagMatch || !valueMatch) continue;

    const openTag = openTagMatch[0];
    const value = valueMatch[1];
    const parsed = parseFirstPrintArea(value);
    if (!parsed) continue;

    const localSheetIdMatch = openTag.match(/\blocalSheetId="(\d+)"/);
    const localSheetId = localSheetIdMatch ? Number(localSheetIdMatch[1]) : null;
    const sheetNameFromValue = normalizeSheetPrefix(parsed.sheetPrefix);
    const isTarget =
      (localSheetId !== null && localSheetId === targetSheetIndex) ||
      sheetNameFromValue === targetSheetName;

    if (!isTarget) continue;
    return parsePrintAreaRanges(value);
  }

  return [];
}

function normalizeSheetPrefix(prefix: string): string {
  const decoded = decodeXml(prefix).trim();
  if (decoded.startsWith("'") && decoded.endsWith("'")) {
    return decoded.slice(1, -1);
  }
  return decoded;
}

function parseWorkbookSheets(workbookXml: string, workbookRelsXml: string): SheetEntry[] {
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
      relId,
      path: relationshipMap.get(relId) || '',
    };
  });
}

/**
 * 印刷範囲を単一範囲に保ったまま、手動改ページ（rowBreaks）でページ割りを表現する。
 * - Print_Area: テンプレートの先頭範囲を行挿入数ぶん下に伸ばした単一範囲に更新
 * - rowBreaks: 行高さの見積もりから改ページ位置を計算してシートXMLへ挿入
 */
function applyPaginationToSheet(params: {
  workbookXml: string;
  sheetXml: string;
  targetSheetName: string;
  targetSheetIndex: number;
  insertedRows: number;
  sharedStrings: string[];
  styleCatalog: StyleCatalog;
  units: KeepTogetherUnit[];
  sections: KeepTogetherUnit[];
}): { workbookXml: string; sheetXml: string } {
  const firstRange = getFirstPrintAreaForSheet(
    params.workbookXml,
    params.targetSheetName,
    params.targetSheetIndex
  );
  if (!firstRange) {
    return { workbookXml: params.workbookXml, sheetXml: params.sheetXml };
  }

  const layoutInfo = parsePageLayoutInfo(params.sheetXml, firstRange);
  const pageCapacity = layoutInfo.pageHeightCapacityPoints;
  if (pageCapacity <= 0) {
    return { workbookXml: params.workbookXml, sheetXml: params.sheetXml };
  }

  const endRow = Math.max(firstRange.endRow + Math.max(0, params.insertedRows), firstRange.startRow);
  const rowHeights = buildEffectiveRowHeightMap(
    params.sheetXml,
    params.sharedStrings,
    params.styleCatalog,
    layoutInfo
  );
  const existingBreakIds = parseManualRowBreakIds(params.sheetXml);
  const forcedBreakIds = new Set(
    existingBreakIds.filter((id) => id >= firstRange.startRow && id < endRow)
  );

  const computedBreaks = computeManualRowBreaks({
    startRow: firstRange.startRow,
    endRow,
    pageCapacity,
    rowHeights,
    defaultRowHeight: layoutInfo.defaultRowHeight,
    units: params.units,
    sections: params.sections,
    forcedBreakIds,
  });

  let sheetXml = params.sheetXml;
  if (computedBreaks.length > 0) {
    sheetXml = upsertRowBreaks(sheetXml, [...existingBreakIds, ...computedBreaks]);
    sheetXml = disableFitToHeight(sheetXml);
  }

  const newValue =
    `${firstRange.sheetPrefix}!$${firstRange.startCol}$${firstRange.startRow}:$${firstRange.endCol}$${endRow}`;
  const workbookXml = params.workbookXml.replace(
    /(<definedName\b[^>]*name="_xlnm\.Print_Area"[^>]*>)([^<]*)(<\/definedName>)/g,
    (full, openTag: string, value: string, closeTag: string) => {
      const localSheetIdMatch = openTag.match(/\blocalSheetId="(\d+)"/);
      const localSheetId = localSheetIdMatch ? Number(localSheetIdMatch[1]) : null;

      const parsed = parseFirstPrintArea(value);
      if (!parsed) return full;

      const sheetNameFromValue = normalizeSheetPrefix(parsed.sheetPrefix);
      const isTarget =
        (localSheetId !== null && localSheetId === params.targetSheetIndex) ||
        sheetNameFromValue === params.targetSheetName;

      if (!isTarget) return full;

      return `${openTag}${newValue}${closeTag}`;
    }
  );

  return { workbookXml, sheetXml };
}

function getNestedValue(source: unknown, path: string): unknown {
  if (!isPlainObject(source)) return undefined;

  const parts = path.split('.');
  let current: unknown = source;
  for (const part of parts) {
    if (!isPlainObject(current)) return undefined;
    current = current[part];
  }
  return current;
}

function getTableData(data: PlaceholderData, section: string, table: string): Array<Record<string, unknown>> {
  const sectionObj = data[section];
  if (isPlainObject(sectionObj)) {
    const nested = sectionObj[table];
    if (Array.isArray(nested)) return nested.filter((v): v is Record<string, unknown> => isPlainObject(v));
  }

  const legacy = data[table];
  if (Array.isArray(legacy)) {
    return legacy.filter((v): v is Record<string, unknown> => isPlainObject(v));
  }

  return [];
}

function replaceSharedStringIndexInRow(
  rowXml: string,
  rowNumber: number,
  oldIndex: number,
  newIndex: number
): string {
  const regex = new RegExp(
    `(<c r="[A-Z]+${rowNumber}"[^>]*>[\\s\\S]*?<v>)${oldIndex}(</v>[\\s\\S]*?<\\/c>)`,
    'g'
  );
  return rowXml.replace(regex, `$1${newIndex}$2`);
}

function renderRecordRowXml(params: {
  rowXml: string;
  rowNumber: number;
  item: Record<string, unknown>;
  placeholders: SectionTablePlaceholder[];
  placeholderToIndices: Map<string, Set<number>>;
  sharedStringsXml: string;
  firstRunPropsByIndex: Map<number, string | null>;
  sectionImageState: SectionImageState;
  recordImageTokens: Map<string, string>;
}): { rowXml: string; sharedStringsXml: string } {
  let rowXml = params.rowXml;
  let sharedStringsXml = params.sharedStringsXml;

  params.placeholders.forEach((ph) => {
    const rawValue = getNestedValue(params.item, ph.cellPath);
    let stringValue = stringifyPrimitiveValue(rawValue);
    if (isImagePlaceholderInput(rawValue)) {
      let imageToken = params.recordImageTokens.get(ph.placeholderKey);
      if (!imageToken) {
        imageToken = buildGeneratedSectionImageToken(params.sectionImageState.nextId++);
        params.recordImageTokens.set(ph.placeholderKey, imageToken);
        params.sectionImageState.images[imageToken] = rawValue;
      }
      stringValue = buildImagePlaceholderText(imageToken);
    }
    const indices = params.placeholderToIndices.get(ph.placeholderKey);
    if (!indices) return;

    indices.forEach((oldIndex) => {
      const added = addSharedString(sharedStringsXml, stringValue, params.firstRunPropsByIndex.get(oldIndex));
      sharedStringsXml = added.updatedXml;
      rowXml = replaceSharedStringIndexInRow(rowXml, params.rowNumber, oldIndex, added.newIndex);
    });
  });

  return { rowXml, sharedStringsXml };
}

function applyLegacyArrayExpansion(
  zip: JSZip,
  sharedStringsXml: string,
  data: PlaceholderData,
  placeholderInfo: Array<{ placeholder: string; key: string; count: number }>
): Promise<{ sharedStringsXml: string; insertedRows: number }> {
  const legacyArrayPlaceholders = placeholderInfo
    .map((p) => parseLegacyArrayPlaceholderKey(p.key))
    .filter((p): p is LegacyArrayPlaceholder => p !== null);

  if (legacyArrayPlaceholders.length === 0) {
    return Promise.resolve({ sharedStringsXml, insertedRows: 0 });
  }

  return (async () => {
    const sheetFile = zip.file('xl/worksheets/sheet1.xml');
    if (!sheetFile) {
      throw new Error('sheet1.xmlが見つかりません（配列展開に必要）');
    }

    let sheetXml = await sheetFile.async('string');
    const sharedStrings = extractSharedStrings(sharedStringsXml);
    const placeholderToIndices = collectPlaceholderIndices(
      sharedStrings,
      legacyArrayPlaceholders.map((p) => p.placeholderKey)
    );
    const styleSourceIndices = new Set<number>();
    placeholderToIndices.forEach((indices) => {
      indices.forEach((idx) => styleSourceIndices.add(idx));
    });
    const firstRunPropsByIndex = collectFirstRunPropertiesBySharedStringIndex(sharedStringsXml, styleSourceIndices);

    const arrayNameToIndices = new Map<string, Set<number>>();
    legacyArrayPlaceholders.forEach((ph) => {
      const indices = placeholderToIndices.get(ph.placeholderKey);
      if (!indices || indices.size === 0) return;
      const current = arrayNameToIndices.get(ph.arrayName) || new Set<number>();
      indices.forEach((idx) => current.add(idx));
      arrayNameToIndices.set(ph.arrayName, current);
    });

    const rowsByArrayName = new Map<string, Set<number>>();
    arrayNameToIndices.forEach((indices, arrayName) => {
      rowsByArrayName.set(arrayName, findRowsBySharedStringIndices(sheetXml, indices));
    });

    const firstArrayName = Array.from(rowsByArrayName.keys())[0];
    const templateRows = Array.from(rowsByArrayName.get(firstArrayName) || []).sort((a, b) => a - b);
    const arrayData = data[firstArrayName];

    if (!Array.isArray(arrayData) || templateRows.length === 0) {
      return { sharedStringsXml, insertedRows: 0 };
    }

    const normalizedData = arrayData.filter((item): item is Record<string, unknown> => isPlainObject(item));
    const templateCapacity = templateRows.length;
    const insertCount = Math.max(0, normalizedData.length - templateCapacity);

    if (insertCount > 0) {
      const lastTemplateRow = templateRows[templateRows.length - 1];
      const insertStartRow = lastTemplateRow + 1;

      sheetXml = shiftRowsDown(sheetXml, insertStartRow, insertCount);
      sheetXml = shiftMergeCellsDown(sheetXml, insertStartRow, insertCount);
      sheetXml = shiftRowBreaksDown(sheetXml, insertStartRow, insertCount);

      const templateRowXml = extractRowXml(sheetXml, lastTemplateRow);
      if (!templateRowXml) {
        throw new Error(`テンプレート行のXMLが取得できません: row=${lastTemplateRow}`);
      }

      const newRowsXml: string[] = [];
      for (let i = 0; i < insertCount; i++) {
        const targetRow = insertStartRow + i;
        newRowsXml.push(copyRowXmlWithShiftedFormulas(templateRowXml, lastTemplateRow, targetRow));
      }

      const lastTemplateRowXml = extractRowXml(sheetXml, lastTemplateRow);
      if (!lastTemplateRowXml) {
        throw new Error(`テンプレート行のXMLが取得できません: row=${lastTemplateRow}`);
      }

      sheetXml = sheetXml.replace(lastTemplateRowXml, `${lastTemplateRowXml}${newRowsXml.join('')}`);
      sheetXml = insertMergeCellsForTemplateBlock(sheetXml, lastTemplateRow, lastTemplateRow, 1, insertCount);
    }

    const totalRows = Math.max(normalizedData.length, templateCapacity);
    for (let i = 0; i < totalRows; i++) {
      const rowNumber = templateRows[0] + i;
      const originalRowXml = extractRowXml(sheetXml, rowNumber);
      if (!originalRowXml) continue;

      let rowXml = originalRowXml;
      legacyArrayPlaceholders.forEach((ph) => {
        if (ph.arrayName !== firstArrayName) return;

        const indices = placeholderToIndices.get(ph.placeholderKey);
        if (!indices) return;

        const item = i < normalizedData.length ? normalizedData[i] : {};
        const rawValue = getNestedValue(item, ph.fieldPath);
        const stringValue = stringifyPrimitiveValue(rawValue);

        indices.forEach((oldIndex) => {
          const added = addSharedString(sharedStringsXml, stringValue, firstRunPropsByIndex.get(oldIndex));
          sharedStringsXml = added.updatedXml;
          rowXml = replaceSharedStringIndexInRow(rowXml, rowNumber, oldIndex, added.newIndex);
        });
      });

      sheetXml = sheetXml.replace(originalRowXml, rowXml);
    }

    zip.file('xl/worksheets/sheet1.xml', sheetXml);
    return { sharedStringsXml, insertedRows: insertCount };
  })();
}

export class PlaceholderReplacer {
  /**
   * Excelファイル（Buffer）内のプレースホルダーを置換
   * sharedStrings.xmlを直接編集することで印刷設定を保持しつつ、
   * 配列プレースホルダーの行追加とPrint_Area拡張を行う
   */
  async replacePlaceholders(
    excelBuffer: Buffer,
    data: PlaceholderData,
    options: PlaceholderReplaceOptions = {}
  ): Promise<Buffer> {
    if (options.images && Object.keys(options.images).length > 0) {
      throw new Error('Top-level "images" is no longer supported. Put image objects inside data section rows.');
    }

    const zip = await JSZip.loadAsync(excelBuffer);

    const sharedStringsFile = zip.file('xl/sharedStrings.xml');
    if (!sharedStringsFile) {
      throw new Error('sharedStrings.xmlが見つかりません');
    }
    const stylesFile = zip.file('xl/styles.xml');

    let sharedStringsXml = await sharedStringsFile.async('string');
    const styleCatalog = stylesFile
      ? parseStyleCatalog(await stylesFile.async('string'))
      : { styles: new Map<number, CellStyleInfo>(), defaultFontSize: 11 };
    const placeholderInfo = detectPlaceholdersInXml(sharedStringsXml);

    const legacyArrayPlaceholders = placeholderInfo
      .map((p) => parseLegacyArrayPlaceholderKey(p.key))
      .filter((p): p is LegacyArrayPlaceholder => p !== null);

    const sectionTablePlaceholders = placeholderInfo
      .map((p) => parseSectionTablePlaceholderKey(p.key))
      .filter((p): p is SectionTablePlaceholder => p !== null);

    if (placeholderInfo.some((p) => p.key.startsWith('%'))) {
      throw new Error('Legacy image placeholder "{{%...}}" is no longer supported. Use image objects in section data instead.');
    }

    if (legacyArrayPlaceholders.length > 0 && sectionTablePlaceholders.length > 0) {
      throw new Error('テンプレート内で旧配列記法（{{#...}}）と新記法（{{##section.table.cell}}）は混在できません');
    }

    let workbookXmlForPrintArea: string | null = null;
    let sheetEntries: SheetEntry[] = [];
    const sheetOffsetByPath = new Map<string, number>();
    const unitsBySheetPath = new Map<string, KeepTogetherUnit[]>();
    const sectionsBySheetPath = new Map<string, KeepTogetherUnit[]>();
    const sectionImageState: SectionImageState = {
      nextId: 1,
      images: {},
    };

    // Stage A-1: 新記法 ##section.table.cell
    if (sectionTablePlaceholders.length > 0) {
      const workbookFile = zip.file('xl/workbook.xml');
      const workbookRelsFile = zip.file('xl/_rels/workbook.xml.rels');
      if (!workbookFile || !workbookRelsFile) {
        throw new Error('workbook.xml または workbook.xml.rels が見つかりません');
      }

      workbookXmlForPrintArea = await workbookFile.async('string');
      const workbookRelsXml = await workbookRelsFile.async('string');
      sheetEntries = parseWorkbookSheets(workbookXmlForPrintArea, workbookRelsXml).filter((s) => !!s.path);

      const sharedStrings = extractSharedStrings(sharedStringsXml);
      const placeholderToIndices = collectPlaceholderIndices(
        sharedStrings,
        sectionTablePlaceholders.map((p) => p.placeholderKey)
      );
      const styleSourceIndices = new Set<number>();
      placeholderToIndices.forEach((indices) => {
        indices.forEach((idx) => styleSourceIndices.add(idx));
      });
      const firstRunPropsByIndex = collectFirstRunPropertiesBySharedStringIndex(sharedStringsXml, styleSourceIndices);

      type Group = {
        section: string;
        table: string;
        placeholders: SectionTablePlaceholder[];
        indexSet: Set<number>;
      };

      const groups = new Map<string, Group>();
      sectionTablePlaceholders.forEach((ph) => {
        const key = `${ph.section}.${ph.table}`;
        const existing = groups.get(key);
        if (existing) {
          existing.placeholders.push(ph);
        } else {
          groups.set(key, {
            section: ph.section,
            table: ph.table,
            placeholders: [ph],
            indexSet: new Set<number>(),
          });
        }

        const indices = placeholderToIndices.get(ph.placeholderKey);
        if (indices) {
          indices.forEach((idx) => groups.get(key)?.indexSet.add(idx));
        }
      });

      const blocks: TableBlock[] = [];
      const templateSheetXmlByPath = new Map<string, string>();
      for (const group of groups.values()) {
        let foundBlock: TableBlock | null = null;

        for (const sheetEntry of sheetEntries) {
          const sheetFile = zip.file(sheetEntry.path);
          if (!sheetFile) continue;

          const sheetXml = await sheetFile.async('string');
          if (!templateSheetXmlByPath.has(sheetEntry.path)) {
            templateSheetXmlByPath.set(sheetEntry.path, sheetXml);
          }
          const rows = Array.from(findRowsBySharedStringIndices(sheetXml, group.indexSet)).sort((a, b) => a - b);
          if (rows.length === 0) continue;

          const minRow = rows[0];
          const maxRow = rows[rows.length - 1];
          if (rows.length !== maxRow - minRow + 1) {
            throw new Error(
              `section=${group.section}, table=${group.table} の行塊は連続行で配置してください（sheet=${sheetEntry.name}）`
            );
          }

          if (foundBlock) {
            throw new Error(`section=${group.section}, table=${group.table} のブロックが複数箇所に配置されています`);
          }

          foundBlock = {
            section: group.section,
            table: group.table,
            sheetPath: sheetEntry.path,
            sheetName: sheetEntry.name,
            sheetIndex: sheetEntry.index,
            startRow: minRow,
            endRow: maxRow,
            blockHeight: maxRow - minRow + 1,
            placeholders: group.placeholders,
          };
        }

        if (!foundBlock) {
          throw new Error(`section=${group.section}, table=${group.table} のプレースホルダー配置を検出できませんでした`);
        }

        blocks.push(foundBlock);
      }

      // section重複禁止
      const seenSection = new Set<string>();
      blocks.forEach((block) => {
        if (seenSection.has(block.section)) {
          throw new Error(`section=${block.section} がテンプレート内で重複しています（sectionは一意にしてください）`);
        }
        seenSection.add(block.section);
      });

      blocks.sort((a, b) => {
        if (a.sheetIndex !== b.sheetIndex) return a.sheetIndex - b.sheetIndex;
        return a.startRow - b.startRow;
      });

      for (const block of blocks) {
        const sheetFile = zip.file(block.sheetPath);
        if (!sheetFile) {
          throw new Error(`対象シートが見つかりません: ${block.sheetPath}`);
        }

        let sheetXml = await sheetFile.async('string');
        const tableData = getTableData(data, block.section, block.table);
        const recordCount = tableData.length;

        const currentOffset = sheetOffsetByPath.get(block.sheetPath) || 0;
        const currentStartRow = block.startRow + currentOffset;
        const currentEndRow = block.endRow + currentOffset;

        const repeatCount = Math.max(0, recordCount - 1);
        const insertedRows = repeatCount * block.blockHeight;

        if (insertedRows > 0) {
          const insertStartRow = currentEndRow + 1;

          sheetXml = shiftRowsDown(sheetXml, insertStartRow, insertedRows);
          sheetXml = shiftMergeCellsDown(sheetXml, insertStartRow, insertedRows);
          sheetXml = shiftRowBreaksDown(sheetXml, insertStartRow, insertedRows);

          const templateRowsXml: string[] = [];
          for (let offset = 0; offset < block.blockHeight; offset++) {
            const templateRowNumber = currentStartRow + offset;
            const rowXml = extractRowXml(sheetXml, templateRowNumber);
            if (!rowXml) {
              throw new Error(`テンプレート行のXMLが取得できません: row=${templateRowNumber}`);
            }
            templateRowsXml.push(rowXml);
          }

          const repeatedRowsXml: string[] = [];
          for (let repeat = 1; repeat <= repeatCount; repeat++) {
            for (let offset = 0; offset < block.blockHeight; offset++) {
              const sourceRow = currentStartRow + offset;
              const targetRow = currentStartRow + repeat * block.blockHeight + offset;
              repeatedRowsXml.push(copyRowXmlWithShiftedFormulas(templateRowsXml[offset], sourceRow, targetRow));
            }
          }

          const endRowXml = extractRowXml(sheetXml, currentEndRow);
          if (!endRowXml) {
            throw new Error(`テンプレート行のXMLが取得できません: row=${currentEndRow}`);
          }

          sheetXml = sheetXml.replace(endRowXml, `${endRowXml}${repeatedRowsXml.join('')}`);
          sheetXml = insertMergeCellsForTemplateBlock(
            sheetXml,
            currentStartRow,
            currentEndRow,
            block.blockHeight,
            repeatCount
          );
        }

        const totalBlocks = Math.max(recordCount, 1);
        let currentRecordStartRow = currentStartRow;
        const blockUnits = unitsBySheetPath.get(block.sheetPath) || [];

        for (let blockIndex = 0; blockIndex < totalBlocks; blockIndex++) {
          const recordImageTokens = new Map<string, string>();
          const item = blockIndex < recordCount ? tableData[blockIndex] : {};

          for (let offset = 0; offset < block.blockHeight; offset++) {
            const rowNumber = currentRecordStartRow + offset;
            const originalRowXml = extractRowXml(sheetXml, rowNumber);
            if (!originalRowXml) continue;

            const rendered = renderRecordRowXml({
              rowXml: originalRowXml,
              rowNumber,
              item,
              placeholders: block.placeholders,
              placeholderToIndices,
              sharedStringsXml,
              firstRunPropsByIndex,
              sectionImageState,
              recordImageTokens,
            });

            sharedStringsXml = rendered.sharedStringsXml;
            sheetXml = sheetXml.replace(originalRowXml, rendered.rowXml);
          }

          blockUnits.push({
            startRow: currentRecordStartRow,
            endRow: currentRecordStartRow + block.blockHeight - 1,
          });
          currentRecordStartRow += block.blockHeight;
        }

        const finalEndRow = currentEndRow + insertedRows;
        sheetOffsetByPath.set(block.sheetPath, currentOffset + insertedRows);
        unitsBySheetPath.set(block.sheetPath, blockUnits);

        const blockSections = sectionsBySheetPath.get(block.sheetPath) || [];
        blockSections.push({
          startRow: currentStartRow,
          endRow: finalEndRow,
        });
        sectionsBySheetPath.set(block.sheetPath, blockSections);

        zip.file(block.sheetPath, sheetXml);
      }
    }

    // Stage A-2: 旧記法 #array.field
    let legacyInsertedRows = 0;
    if (sectionTablePlaceholders.length === 0 && legacyArrayPlaceholders.length > 0) {
      const legacy = await applyLegacyArrayExpansion(zip, sharedStringsXml, data, placeholderInfo);
      sharedStringsXml = legacy.sharedStringsXml;
      legacyInsertedRows = legacy.insertedRows;
    }

    // Stage B: 通常プレースホルダー（先頭run書式を維持して sharedStrings を置換）
    const primitiveReplacements = new Map<string, string>();
    Object.entries(data).forEach(([key, value]) => {
      if (Array.isArray(value) || isPlainObject(value)) return;
      primitiveReplacements.set(key, stringifyPrimitiveValue(value));
    });

    let replacedSharedStrings = replacePrimitivePlaceholdersWithFirstRunStyle(
      sharedStringsXml,
      primitiveReplacements
    );

    // Stage C: ページ割り更新（単一Print_Area + 手動rowBreaks）
    // 置換後の文字列で行高さを見積もるため、Stage Bの後に実行する
    const paginationSharedStrings = extractSharedStrings(replacedSharedStrings);
    const imageKeys = new Set(Object.keys(sectionImageState.images));

    if (workbookXmlForPrintArea) {
      for (const sheetEntry of sheetEntries) {
        if (!sheetOffsetByPath.has(sheetEntry.path)) continue;

        const sheetFile = zip.file(sheetEntry.path);
        if (!sheetFile) continue;

        const sheetXml = await sheetFile.async('string');
        const units = [...(unitsBySheetPath.get(sheetEntry.path) || [])];
        if (imageKeys.size > 0) {
          findImagePlaceholderBlocks(sheetXml, paginationSharedStrings, imageKeys).forEach((imageBlock) => {
            units.push({ startRow: imageBlock.startRow, endRow: imageBlock.endRow });
          });
        }

        const paginated = applyPaginationToSheet({
          workbookXml: workbookXmlForPrintArea,
          sheetXml,
          targetSheetName: sheetEntry.name,
          targetSheetIndex: sheetEntry.index,
          insertedRows: sheetOffsetByPath.get(sheetEntry.path) || 0,
          sharedStrings: paginationSharedStrings,
          styleCatalog,
          units,
          sections: sectionsBySheetPath.get(sheetEntry.path) || [],
        });

        workbookXmlForPrintArea = paginated.workbookXml;
        if (paginated.sheetXml !== sheetXml) {
          zip.file(sheetEntry.path, paginated.sheetXml);
        }
      }

      zip.file('xl/workbook.xml', workbookXmlForPrintArea);
    }

    // 旧記法（sheet1想定の既存互換）のページ割り更新
    if (legacyInsertedRows > 0 && !workbookXmlForPrintArea) {
      const workbookFile = zip.file('xl/workbook.xml');
      const sheet1File = zip.file('xl/worksheets/sheet1.xml');

      if (workbookFile && sheet1File) {
        const workbookXml = await workbookFile.async('string');
        const sheet1Xml = await sheet1File.async('string');
        const firstSheetNameMatch = workbookXml.match(/<sheet\b[^>]*name="([^"]*)"/);
        const firstSheetName = firstSheetNameMatch ? decodeXml(firstSheetNameMatch[1]) : '';

        const paginated = applyPaginationToSheet({
          workbookXml,
          sheetXml: sheet1Xml,
          targetSheetName: firstSheetName,
          targetSheetIndex: 0,
          insertedRows: legacyInsertedRows,
          sharedStrings: paginationSharedStrings,
          styleCatalog,
          units: [],
          sections: [],
        });

        zip.file('xl/workbook.xml', paginated.workbookXml);
        if (paginated.sheetXml !== sheet1Xml) {
          zip.file('xl/worksheets/sheet1.xml', paginated.sheetXml);
        }
      }
    }

    replacedSharedStrings = await embedImagePlaceholders({
      zip,
      sharedStringsXml: replacedSharedStrings,
      images: sectionImageState.images,
    });

    zip.file('xl/sharedStrings.xml', replacedSharedStrings);

    return zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 9 },
    });
  }

  /**
   * Excelファイル（Buffer）内のプレースホルダーを検出
   */
  async findPlaceholders(excelBuffer: Buffer): Promise<string[]> {
    const zip = await JSZip.loadAsync(excelBuffer);
    const sharedStringsFile = zip.file('xl/sharedStrings.xml');
    if (!sharedStringsFile) return [];

    const sharedStringsXml = await sharedStringsFile.async('string');
    const placeholders = detectPlaceholdersInXml(sharedStringsXml);
    const keys = placeholders.map((p) => p.key);
    return Array.from(new Set(keys)).sort();
  }

  /**
   * プレースホルダーの詳細情報を取得
   */
  async getPlaceholderInfo(excelBuffer: Buffer): Promise<Array<{
    placeholder: string;
    key: string;
    count: number;
  }>> {
    const zip = await JSZip.loadAsync(excelBuffer);
    const sharedStringsFile = zip.file('xl/sharedStrings.xml');
    if (!sharedStringsFile) return [];

    const sharedStringsXml = await sharedStringsFile.async('string');
    return detectPlaceholdersInXml(sharedStringsXml);
  }

}
