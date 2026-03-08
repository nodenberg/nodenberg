import JSZip from 'jszip';
import { embedImagePlaceholders, PlaceholderImages } from './imagePlaceholderReplacer';

export type PlaceholderPrimitive = string | number | Date | null;
export type PlaceholderObject = Record<string, unknown>;
export type PlaceholderArray = Array<PlaceholderObject>;
export type PlaceholderValue = PlaceholderPrimitive | PlaceholderArray | PlaceholderObject;

export interface PlaceholderData {
  [key: string]: PlaceholderValue;
}

export interface PlaceholderReplaceOptions {
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

function getPrintAreaFromWorkbookXml(workbookXml: string): string | null {
  const m = workbookXml.match(/name="_xlnm\.Print_Area"[^>]*>([^<]+)<\/definedName>/);
  return m ? m[1] : null;
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

type PrintAreaTemplateInfo = {
  firstRange: PrintAreaRange;
  basePageHeight: number;
  defaultRowHeight: number;
  layoutInfo: PageLayoutInfo;
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

function buildRowHeightMap(sheetXml: string): Map<number, number> {
  const result = new Map<number, number>();
  const rowRegex = /<row\b[^>]*\br="(\d+)"[^>]*>/g;
  let match: RegExpExecArray | null;

  while ((match = rowRegex.exec(sheetXml)) !== null) {
    const rowNumber = Number(match[1]);
    const rowTag = match[0];
    const htMatch = rowTag.match(/\bht="([\d.]+)"/);
    if (!htMatch) continue;
    result.set(rowNumber, Number(htMatch[1]));
  }

  return result;
}

function getMaxRowNumber(sheetXml: string): number {
  const rowRegex = /<row\b[^>]*\br="(\d+)"[^>]*>/g;
  let maxRow = 0;
  let match: RegExpExecArray | null;

  while ((match = rowRegex.exec(sheetXml)) !== null) {
    maxRow = Math.max(maxRow, Number(match[1]));
  }

  return maxRow;
}

function buildExistingRowSet(sheetXml: string): Set<number> {
  const rows = new Set<number>();
  const rowRegex = /<row\b[^>]*\br="(\d+)"[^>]*>/g;
  let match: RegExpExecArray | null;

  while ((match = rowRegex.exec(sheetXml)) !== null) {
    rows.add(Number(match[1]));
  }

  return rows;
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
  const header = Number(pageMarginsTag.match(/\bheader="([\d.]+)"/)?.[1] || '0.3') * 72;
  const footer = Number(pageMarginsTag.match(/\bfooter="([\d.]+)"/)?.[1] || '0.3') * 72;

  const printableWidthPoints = Math.max(1, pageWidthPoints - left - right);
  const printableHeightPoints = Math.max(1, pageHeightPoints - top - bottom - header - footer);

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

function sumRowHeights(
  startRow: number,
  endRow: number,
  rowHeights: Map<number, number>,
  defaultRowHeight: number
): number {
  if (endRow < startRow) return 0;

  let sum = 0;
  for (let row = startRow; row <= endRow; row++) {
    sum += rowHeights.get(row) ?? defaultRowHeight;
  }
  return sum;
}

function buildPrintAreasByHeight(params: {
  sheetPrefix: string;
  startCol: string;
  startRow: number;
  endCol: string;
  contentEndRow: number;
  basePageHeight: number;
  rowHeights: Map<number, number>;
  defaultRowHeight: number;
  existingRows?: Set<number>;
}): string {
  if (params.contentEndRow < params.startRow) {
    return `${params.sheetPrefix}!$${params.startCol}$${params.startRow}:$${params.endCol}$${params.startRow}`;
  }

  if (params.basePageHeight <= 0) {
    return `${params.sheetPrefix}!$${params.startCol}$${params.startRow}:$${params.endCol}$${params.contentEndRow}`;
  }

  const ranges: string[] = [];
  let currentStart = params.startRow;

  while (currentStart <= params.contentEndRow) {
    let currentHeight = 0;
    let currentEnd = currentStart;

    while (currentEnd <= params.contentEndRow) {
      const rawRowHeight = params.rowHeights.get(currentEnd) ?? params.defaultRowHeight;
      const rowHeight = rawRowHeight > 0 ? rawRowHeight : 1;
      if (currentEnd > currentStart && currentHeight + rowHeight > params.basePageHeight) {
        currentEnd -= 1;
        break;
      }
      currentHeight += rowHeight;
      currentEnd += 1;

      if (currentHeight >= params.basePageHeight) {
        currentEnd -= 1;
        break;
      }
    }

    if (currentEnd > params.contentEndRow) {
      currentEnd = params.contentEndRow;
    }

    if (currentEnd < currentStart) {
      currentEnd = currentStart;
    }

    if (params.existingRows && params.existingRows.size > 0) {
      while (currentEnd < params.contentEndRow && !params.existingRows.has(currentEnd + 1)) {
        currentEnd += 1;
      }
    }

    ranges.push(`${params.sheetPrefix}!$${params.startCol}$${currentStart}:$${params.endCol}$${currentEnd}`);
    currentStart = currentEnd + 1;
  }

  return ranges.join(',');
}

function sumRowHeightsBeforeRow(
  rowExclusive: number,
  startRow: number,
  rowHeights: Map<number, number>,
  defaultRowHeight: number
): number {
  if (rowExclusive <= startRow) return 0;

  let sum = 0;
  for (let row = startRow; row < rowExclusive; row++) {
    sum += rowHeights.get(row) ?? defaultRowHeight;
  }
  return sum;
}

function doesSectionOverflowPages(params: {
  sectionStartRow: number;
  sectionEndRow: number;
  pageStartRow: number;
  basePageHeight: number;
  rowHeights: Map<number, number>;
  defaultRowHeight: number;
}): boolean {
  if (params.sectionEndRow < params.sectionStartRow) return false;
  if (params.basePageHeight <= 0) return false;

  const startHeight = sumRowHeightsBeforeRow(
    params.sectionStartRow,
    params.pageStartRow,
    params.rowHeights,
    params.defaultRowHeight
  );
  const endHeightExclusive = sumRowHeightsBeforeRow(
    params.sectionEndRow + 1,
    params.pageStartRow,
    params.rowHeights,
    params.defaultRowHeight
  );
  const startPage = Math.floor(startHeight / params.basePageHeight);
  const endPage = Math.floor(Math.max(0, endHeightExclusive - 0.000001) / params.basePageHeight);
  return endPage > startPage;
}

function calculateBlankRowsToNextPageStart(params: {
  nextSectionStartRow: number;
  pageStartRow: number;
  basePageHeight: number;
  rowHeights: Map<number, number>;
  defaultRowHeight: number;
}): number {
  if (params.basePageHeight <= 0 || params.defaultRowHeight <= 0) return 0;

  const usedHeight = sumRowHeightsBeforeRow(
    params.nextSectionStartRow,
    params.pageStartRow,
    params.rowHeights,
    params.defaultRowHeight
  );
  const remainder = usedHeight % params.basePageHeight;
  if (remainder === 0) return 0;

  const remaining = params.basePageHeight - remainder;
  return Math.max(1, Math.ceil(remaining / params.defaultRowHeight));
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

function updatePrintAreaForSheet(
  workbookXml: string,
  targetSheetName: string,
  targetSheetIndex: number,
  sheetXml: string,
  sharedStrings: string[],
  styleCatalog: StyleCatalog
): string {
  return workbookXml.replace(
    /(<definedName\b[^>]*name="_xlnm\.Print_Area"[^>]*>)([^<]*)(<\/definedName>)/g,
    (full, openTag: string, value: string, closeTag: string) => {
      const localSheetIdMatch = openTag.match(/\blocalSheetId="(\d+)"/);
      const localSheetId = localSheetIdMatch ? Number(localSheetIdMatch[1]) : null;

      const parsed = parseFirstPrintArea(value);
      if (!parsed) return full;

      const sheetNameFromValue = normalizeSheetPrefix(parsed.sheetPrefix);
      const isTarget =
        (localSheetId !== null && localSheetId === targetSheetIndex) ||
        sheetNameFromValue === targetSheetName;

      if (!isTarget) return full;

      const parsedRanges = parsePrintAreaRanges(value);
      if (parsedRanges.length === 0) return full;

      const firstRange = parsedRanges[0];
      const contentEndRow = Math.max(getMaxRowNumber(sheetXml), firstRange.endRow);
      const existingRows = buildExistingRowSet(sheetXml);
      const layoutInfo = parsePageLayoutInfo(sheetXml, firstRange);
      const rowHeights = buildEffectiveRowHeightMap(sheetXml, sharedStrings, styleCatalog, layoutInfo);
      const basePageHeight = layoutInfo.pageHeightCapacityPoints;

      if (basePageHeight <= 0 || contentEndRow < firstRange.startRow) {
        return full;
      }

      const newValue = buildPrintAreasByHeight({
        sheetPrefix: firstRange.sheetPrefix,
        startCol: firstRange.startCol,
        startRow: firstRange.startRow,
        endCol: firstRange.endCol,
        contentEndRow,
        basePageHeight,
        rowHeights,
        defaultRowHeight: layoutInfo.defaultRowHeight,
        existingRows,
      });

      return `${openTag}${newValue}${closeTag}`;
    }
  );
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
}): { rowXml: string; sharedStringsXml: string } {
  let rowXml = params.rowXml;
  let sharedStringsXml = params.sharedStringsXml;

  params.placeholders.forEach((ph) => {
    const rawValue = getNestedValue(params.item, ph.cellPath);
    const stringValue = stringifyPrimitiveValue(rawValue);
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

function renderRecordRowXmlForEstimate(params: {
  rowXml: string;
  rowNumber: number;
  item: Record<string, unknown>;
  placeholders: SectionTablePlaceholder[];
  placeholderToIndices: Map<string, Set<number>>;
  sharedStrings: string[];
}): { rowXml: string; sharedStrings: string[] } {
  let rowXml = params.rowXml;
  const sharedStrings = [...params.sharedStrings];

  params.placeholders.forEach((ph) => {
    const rawValue = getNestedValue(params.item, ph.cellPath);
    const stringValue = stringifyPrimitiveValue(rawValue);
    const indices = params.placeholderToIndices.get(ph.placeholderKey);
    if (!indices) return;

    indices.forEach((oldIndex) => {
      const newIndex = sharedStrings.length;
      sharedStrings.push(stringValue);
      rowXml = replaceSharedStringIndexInRow(rowXml, params.rowNumber, oldIndex, newIndex);
    });
  });

  return { rowXml, sharedStrings };
}

function estimateRecordBlockHeight(params: {
  sheetXml: string;
  startRow: number;
  blockHeight: number;
  item: Record<string, unknown>;
  placeholders: SectionTablePlaceholder[];
  placeholderToIndices: Map<string, Set<number>>;
  sharedStrings: string[];
  styleCatalog: StyleCatalog;
  layoutInfo: PageLayoutInfo;
}): number {
  const mergeSpans = parseMergeSpans(params.sheetXml);
  let previewSharedStrings = [...params.sharedStrings];
  let totalHeight = 0;

  for (let offset = 0; offset < params.blockHeight; offset++) {
    const rowNumber = params.startRow + offset;
    const rowXml = extractRowXml(params.sheetXml, rowNumber);
    if (!rowXml) continue;

    const rendered = renderRecordRowXmlForEstimate({
      rowXml,
      rowNumber,
      item: params.item,
      placeholders: params.placeholders,
      placeholderToIndices: params.placeholderToIndices,
      sharedStrings: previewSharedStrings,
    });

    previewSharedStrings = rendered.sharedStrings;
    totalHeight += estimateRowRenderedHeight(
      rendered.rowXml,
      previewSharedStrings,
      params.styleCatalog,
      params.layoutInfo,
      mergeSpans
    );
  }

  return totalHeight;
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

    if (legacyArrayPlaceholders.length > 0 && sectionTablePlaceholders.length > 0) {
      throw new Error('テンプレート内で旧配列記法（{{#...}}）と新記法（{{##section.table.cell}}）は混在できません');
    }

    let workbookXmlForPrintArea: string | null = null;
    let sheetEntries: SheetEntry[] = [];

    // Stage A-1: 新記法 ##section.table.cell
    if (sectionTablePlaceholders.length > 0) {
      const workbookFile = zip.file('xl/workbook.xml');
      const workbookRelsFile = zip.file('xl/_rels/workbook.xml.rels');
      if (!workbookFile || !workbookRelsFile) {
        throw new Error('workbook.xml または workbook.xml.rels が見つかりません');
      }

      workbookXmlForPrintArea = await workbookFile.async('string');
      const templateWorkbookXmlForPrintArea = workbookXmlForPrintArea;
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

      const printAreaTemplateInfoBySheetPath = new Map<string, PrintAreaTemplateInfo>();
      for (const block of blocks) {
        if (printAreaTemplateInfoBySheetPath.has(block.sheetPath)) continue;

        const templateSheetXml = templateSheetXmlByPath.get(block.sheetPath);
        if (!templateSheetXml) continue;

        const firstRange = getFirstPrintAreaForSheet(
          templateWorkbookXmlForPrintArea,
          block.sheetName,
          block.sheetIndex
        );
        if (!firstRange) continue;

        const layoutInfo = parsePageLayoutInfo(templateSheetXml, firstRange);
        const defaultRowHeight = layoutInfo.defaultRowHeight;
        const basePageHeight = layoutInfo.pageHeightCapacityPoints;

        if (basePageHeight <= 0) continue;

        printAreaTemplateInfoBySheetPath.set(block.sheetPath, {
          firstRange,
          basePageHeight,
          defaultRowHeight,
          layoutInfo,
        });
      }

      const sheetOffsetByPath = new Map<string, number>();
      const prevSectionSpanBySheetPath = new Map<string, { startRow: number; endRow: number }>();
      const pageUsedHeightBySheetPath = new Map<string, number>();

      for (const block of blocks) {
        const sheetFile = zip.file(block.sheetPath);
        if (!sheetFile) {
          throw new Error(`対象シートが見つかりません: ${block.sheetPath}`);
        }

        let sheetXml = await sheetFile.async('string');
        const tableData = getTableData(data, block.section, block.table);
        const recordCount = tableData.length;

        let totalInsertedRowsForPrintArea = 0;
        const currentOffset = sheetOffsetByPath.get(block.sheetPath) || 0;
        let currentStartRow = block.startRow + currentOffset;
        let currentEndRow = block.endRow + currentOffset;
        const templateInfo = printAreaTemplateInfoBySheetPath.get(block.sheetPath);
        let currentSharedStrings = extractSharedStrings(sharedStringsXml);
        let currentRowHeights = templateInfo
          ? buildEffectiveRowHeightMap(sheetXml, currentSharedStrings, styleCatalog, templateInfo.layoutInfo)
          : buildRowHeightMap(sheetXml);
        let currentPageUsedHeight = templateInfo
          ? pageUsedHeightBySheetPath.get(block.sheetPath) ??
            sumRowHeights(
              templateInfo.firstRange.startRow,
              currentStartRow - 1,
              currentRowHeights,
              templateInfo.defaultRowHeight
            )
          : 0;
        const prevSectionSpan = prevSectionSpanBySheetPath.get(block.sheetPath);
        if (templateInfo && prevSectionSpan) {
          const sectionOverflowed = doesSectionOverflowPages({
            sectionStartRow: prevSectionSpan.startRow,
            sectionEndRow: prevSectionSpan.endRow,
            pageStartRow: templateInfo.firstRange.startRow,
            basePageHeight: templateInfo.basePageHeight,
            rowHeights: currentRowHeights,
            defaultRowHeight: templateInfo.defaultRowHeight,
          });

          if (sectionOverflowed) {
            const blankRows = calculateBlankRowsToNextPageStart({
              nextSectionStartRow: currentStartRow,
              pageStartRow: templateInfo.firstRange.startRow,
              basePageHeight: templateInfo.basePageHeight,
              rowHeights: currentRowHeights,
              defaultRowHeight: templateInfo.defaultRowHeight,
            });

            if (blankRows > 0) {
              sheetXml = shiftRowsDown(sheetXml, currentStartRow, blankRows);
              sheetXml = shiftMergeCellsDown(sheetXml, currentStartRow, blankRows);
              sheetXml = shiftRowBreaksDown(sheetXml, currentStartRow, blankRows);
              currentRowHeights = buildEffectiveRowHeightMap(
                sheetXml,
                currentSharedStrings,
                styleCatalog,
                templateInfo.layoutInfo
              );
              currentPageUsedHeight = 0;
              currentStartRow += blankRows;
              currentEndRow += blankRows;
              totalInsertedRowsForPrintArea += blankRows;
            }
          }
        }

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
          totalInsertedRowsForPrintArea += insertedRows;
          currentRowHeights = templateInfo
            ? buildEffectiveRowHeightMap(sheetXml, currentSharedStrings, styleCatalog, templateInfo.layoutInfo)
            : buildRowHeightMap(sheetXml);
        }

        const totalBlocks = Math.max(recordCount, 1);
        let currentRecordStartRow = currentStartRow;
        let pageBreakInsertedRowsInBlock = 0;

        for (let blockIndex = 0; blockIndex < totalBlocks; blockIndex++) {
          if (blockIndex > 0 && templateInfo) {
            const itemForEstimate = blockIndex < recordCount ? tableData[blockIndex] : {};
            const estimatedRecordHeight = estimateRecordBlockHeight({
              sheetXml,
              startRow: currentRecordStartRow,
              blockHeight: block.blockHeight,
              item: itemForEstimate,
              placeholders: block.placeholders,
              placeholderToIndices,
              sharedStrings: currentSharedStrings,
              styleCatalog,
              layoutInfo: templateInfo.layoutInfo,
            });
            const currentRecordWouldOverflow =
              currentPageUsedHeight > 0 &&
              currentPageUsedHeight + estimatedRecordHeight > templateInfo.basePageHeight + 0.000001;

            if (currentRecordWouldOverflow) {
              const remainingHeight = templateInfo.basePageHeight - currentPageUsedHeight;
              const blankRows = Math.max(1, Math.ceil(remainingHeight / templateInfo.defaultRowHeight));

              if (blankRows > 0) {
                sheetXml = shiftRowsDown(sheetXml, currentRecordStartRow, blankRows);
                sheetXml = shiftMergeCellsDown(sheetXml, currentRecordStartRow, blankRows);
                sheetXml = shiftRowBreaksDown(sheetXml, currentRecordStartRow, blankRows);
                currentRowHeights = buildEffectiveRowHeightMap(
                  sheetXml,
                  currentSharedStrings,
                  styleCatalog,
                  templateInfo.layoutInfo
                );
                currentPageUsedHeight = 0;
                currentRecordStartRow += blankRows;
                totalInsertedRowsForPrintArea += blankRows;
                pageBreakInsertedRowsInBlock += blankRows;
              }
            }
          }

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
            });

            sharedStringsXml = rendered.sharedStringsXml;
            sheetXml = sheetXml.replace(originalRowXml, rendered.rowXml);
          }

          if (templateInfo) {
            currentSharedStrings = extractSharedStrings(sharedStringsXml);
            currentRowHeights = buildEffectiveRowHeightMap(
              sheetXml,
              currentSharedStrings,
              styleCatalog,
              templateInfo.layoutInfo
            );
            currentPageUsedHeight += sumRowHeights(
              currentRecordStartRow,
              currentRecordStartRow + block.blockHeight - 1,
              currentRowHeights,
              templateInfo.defaultRowHeight
            );
          }

          currentRecordStartRow += block.blockHeight;
        }

        const finalEndRow = currentEndRow + insertedRows + pageBreakInsertedRowsInBlock;
        const newOffset = currentOffset + totalInsertedRowsForPrintArea;
        sheetOffsetByPath.set(block.sheetPath, newOffset);
        prevSectionSpanBySheetPath.set(block.sheetPath, {
          startRow: currentStartRow,
          endRow: finalEndRow,
        });
        if (templateInfo) {
          pageUsedHeightBySheetPath.set(block.sheetPath, currentPageUsedHeight);
        }

        zip.file(block.sheetPath, sheetXml);

        if (workbookXmlForPrintArea) {
          workbookXmlForPrintArea = updatePrintAreaForSheet(
            workbookXmlForPrintArea,
            block.sheetName,
            block.sheetIndex,
            sheetXml,
            extractSharedStrings(sharedStringsXml),
            styleCatalog
          );
        }
      }
    }

    // Stage A-2: 旧記法 #array.field
    let legacyInsertedRows = 0;
    if (sectionTablePlaceholders.length === 0 && legacyArrayPlaceholders.length > 0) {
      const legacy = await applyLegacyArrayExpansion(zip, sharedStringsXml, data, placeholderInfo);
      sharedStringsXml = legacy.sharedStringsXml;
      legacyInsertedRows = legacy.insertedRows;
    }

    const imageKeys = new Set(Object.keys(options.images || {}));
    if (imageKeys.size > 0) {
      const workbookFile = zip.file('xl/workbook.xml');
      const workbookRelsFile = zip.file('xl/_rels/workbook.xml.rels');

      if (workbookFile && workbookRelsFile) {
        if (!workbookXmlForPrintArea) {
          workbookXmlForPrintArea = await workbookFile.async('string');
        }

        const workbookRelsXml = await workbookRelsFile.async('string');
        const targetSheetEntries = sheetEntries.length > 0
          ? sheetEntries
          : parseWorkbookSheets(workbookXmlForPrintArea, workbookRelsXml).filter((s) => !!s.path);
        const currentSharedStrings = extractSharedStrings(sharedStringsXml);

        for (const sheetEntry of targetSheetEntries) {
          const sheetFile = zip.file(sheetEntry.path);
          if (!sheetFile) continue;

          let sheetXml = await sheetFile.async('string');
          const imageBlocks = findImagePlaceholderBlocks(sheetXml, currentSharedStrings, imageKeys);
          if (imageBlocks.length === 0) continue;

          const firstRange = getFirstPrintAreaForSheet(
            workbookXmlForPrintArea,
            sheetEntry.name,
            sheetEntry.index
          );
          if (!firstRange) continue;

          const layoutInfo = parsePageLayoutInfo(sheetXml, firstRange);
          if (layoutInfo.pageHeightCapacityPoints <= 0) continue;

          let rowHeights = buildEffectiveRowHeightMap(sheetXml, currentSharedStrings, styleCatalog, layoutInfo);
          let printAreaRanges = getPrintAreaRangesForSheet(
            workbookXmlForPrintArea,
            sheetEntry.name,
            sheetEntry.index
          );
          let insertedRows = 0;

          for (const block of imageBlocks) {
            const startRow = block.startRow + insertedRows;
            const endRow = block.endRow + insertedRows;

            const crossingRange = printAreaRanges.find((range) =>
              startRow >= range.startRow &&
              startRow <= range.endRow &&
              endRow > range.endRow
            );

            let blankRows = 0;
            if (crossingRange) {
              blankRows = crossingRange.endRow + 1 - startRow;
            } else {
              const imageOverflowed = doesSectionOverflowPages({
                sectionStartRow: startRow,
                sectionEndRow: endRow,
                pageStartRow: firstRange.startRow,
                basePageHeight: layoutInfo.pageHeightCapacityPoints,
                rowHeights,
                defaultRowHeight: layoutInfo.defaultRowHeight,
              });

              if (imageOverflowed) {
                blankRows = calculateBlankRowsToNextPageStart({
                  nextSectionStartRow: startRow,
                  pageStartRow: firstRange.startRow,
                  basePageHeight: layoutInfo.pageHeightCapacityPoints,
                  rowHeights,
                  defaultRowHeight: layoutInfo.defaultRowHeight,
                });
              }
            }

            if (blankRows <= 0) continue;

            sheetXml = shiftRowsDown(sheetXml, startRow, blankRows);
            sheetXml = shiftMergeCellsDown(sheetXml, startRow, blankRows);
            sheetXml = shiftRowBreaksDown(sheetXml, startRow, blankRows);
            insertedRows += blankRows;
            rowHeights = buildEffectiveRowHeightMap(sheetXml, currentSharedStrings, styleCatalog, layoutInfo);
            workbookXmlForPrintArea = updatePrintAreaForSheet(
              workbookXmlForPrintArea,
              sheetEntry.name,
              sheetEntry.index,
              sheetXml,
              currentSharedStrings,
              styleCatalog
            );
            printAreaRanges = getPrintAreaRangesForSheet(
              workbookXmlForPrintArea,
              sheetEntry.name,
              sheetEntry.index
            );
          }

          if (insertedRows > 0) {
            zip.file(sheetEntry.path, sheetXml);
          }
        }
      }
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

    if (workbookXmlForPrintArea) {
      zip.file('xl/workbook.xml', workbookXmlForPrintArea);
    }

    replacedSharedStrings = await embedImagePlaceholders({
      zip,
      sharedStringsXml: replacedSharedStrings,
      images: options.images,
    });
    // Stage C: 旧記法でのPrint_Area更新（sheet1想定の既存互換）
    if (legacyInsertedRows > 0 && !workbookXmlForPrintArea) {
      const workbookFile = zip.file('xl/workbook.xml');
      if (workbookFile) {
        const workbookXml = await workbookFile.async('string');
        const currentPrintArea = getPrintAreaFromWorkbookXml(workbookXml);
        const parsed = currentPrintArea ? parseFirstPrintArea(currentPrintArea) : null;

        if (parsed) {
          const sheet1File = zip.file('xl/worksheets/sheet1.xml');
          const sheet1Xml = sheet1File ? await sheet1File.async('string') : null;
          if (sheet1Xml) {
            const firstSheetNameMatch = workbookXml.match(/<sheet\b[^>]*name="([^"]*)"/);
            const firstSheetName = firstSheetNameMatch ? decodeXml(firstSheetNameMatch[1]) : '';
            const updatedWorkbookXml = updatePrintAreaForSheet(
              workbookXml,
              firstSheetName,
              0,
              sheet1Xml,
              extractSharedStrings(sharedStringsXml),
              styleCatalog
            );
            zip.file('xl/workbook.xml', updatedWorkbookXml);
          }
        }
      }
    }

    if (workbookXmlForPrintArea) {
      zip.file('xl/workbook.xml', workbookXmlForPrintArea);
    }

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
