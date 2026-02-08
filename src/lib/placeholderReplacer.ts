import JSZip from 'jszip';

export type PlaceholderPrimitive = string | number | Date | null;
export type PlaceholderObject = Record<string, unknown>;
export type PlaceholderArray = Array<PlaceholderObject>;
export type PlaceholderValue = PlaceholderPrimitive | PlaceholderArray | PlaceholderObject;

export interface PlaceholderData {
  [key: string]: PlaceholderValue;
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
      strings.push(chunks.join(''));
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

function buildPagedPrintAreas(params: {
  sheetPrefix: string;
  startCol: string;
  startRow: number;
  endCol: string;
  pageHeight: number;
  pages: number;
}): string {
  const ranges: string[] = [];
  for (let pageIndex = 0; pageIndex < params.pages; pageIndex++) {
    const pageStartRow = params.startRow + pageIndex * params.pageHeight;
    const pageEndRow = pageStartRow + params.pageHeight - 1;
    ranges.push(`${params.sheetPrefix}!$${params.startCol}$${pageStartRow}:$${params.endCol}$${pageEndRow}`);
  }
  return ranges.join(',');
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
  insertedRows: number
): string {
  if (insertedRows <= 0) return workbookXml;

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

      const baseRows = parsed.endRow - parsed.startRow + 1;
      if (baseRows <= 0) return full;

      const existingPages = Math.max(1, value.split(',').length);
      const requiredPages = Math.ceil((baseRows + insertedRows) / baseRows);
      const pages = Math.max(existingPages, requiredPages);

      if (pages <= existingPages && existingPages > 1) {
        return full;
      }

      const newValue = buildPagedPrintAreas({
        sheetPrefix: parsed.sheetPrefix,
        startCol: parsed.startCol,
        startRow: parsed.startRow,
        endCol: parsed.endCol,
        pageHeight: baseRows,
        pages,
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
  async replacePlaceholders(excelBuffer: Buffer, data: PlaceholderData): Promise<Buffer> {
    const zip = await JSZip.loadAsync(excelBuffer);

    const sharedStringsFile = zip.file('xl/sharedStrings.xml');
    if (!sharedStringsFile) {
      throw new Error('sharedStrings.xmlが見つかりません');
    }

    let sharedStringsXml = await sharedStringsFile.async('string');
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
      for (const group of groups.values()) {
        let foundBlock: TableBlock | null = null;

        for (const sheetEntry of sheetEntries) {
          const sheetFile = zip.file(sheetEntry.path);
          if (!sheetFile) continue;

          const sheetXml = await sheetFile.async('string');
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

      for (const block of blocks) {
        const sheetFile = zip.file(block.sheetPath);
        if (!sheetFile) {
          throw new Error(`対象シートが見つかりません: ${block.sheetPath}`);
        }

        let sheetXml = await sheetFile.async('string');
        const tableData = getTableData(data, block.section, block.table);
        const recordCount = tableData.length;

        const repeatCount = Math.max(0, recordCount - 1);
        const insertedRows = repeatCount * block.blockHeight;

        if (insertedRows > 0) {
          const insertStartRow = block.endRow + 1;

          sheetXml = shiftRowsDown(sheetXml, insertStartRow, insertedRows);
          sheetXml = shiftMergeCellsDown(sheetXml, insertStartRow, insertedRows);

          const templateRowsXml: string[] = [];
          for (let offset = 0; offset < block.blockHeight; offset++) {
            const templateRowNumber = block.startRow + offset;
            const rowXml = extractRowXml(sheetXml, templateRowNumber);
            if (!rowXml) {
              throw new Error(`テンプレート行のXMLが取得できません: row=${templateRowNumber}`);
            }
            templateRowsXml.push(rowXml);
          }

          const repeatedRowsXml: string[] = [];
          for (let repeat = 1; repeat <= repeatCount; repeat++) {
            for (let offset = 0; offset < block.blockHeight; offset++) {
              const sourceRow = block.startRow + offset;
              const targetRow = block.startRow + repeat * block.blockHeight + offset;
              repeatedRowsXml.push(copyRowXmlWithShiftedFormulas(templateRowsXml[offset], sourceRow, targetRow));
            }
          }

          const endRowXml = extractRowXml(sheetXml, block.endRow);
          if (!endRowXml) {
            throw new Error(`テンプレート行のXMLが取得できません: row=${block.endRow}`);
          }

          sheetXml = sheetXml.replace(endRowXml, `${endRowXml}${repeatedRowsXml.join('')}`);
          sheetXml = insertMergeCellsForTemplateBlock(
            sheetXml,
            block.startRow,
            block.endRow,
            block.blockHeight,
            repeatCount
          );
        }

        const totalBlocks = Math.max(recordCount, 1);

        for (let blockIndex = 0; blockIndex < totalBlocks; blockIndex++) {
          const item = blockIndex < recordCount ? tableData[blockIndex] : {};

          for (let offset = 0; offset < block.blockHeight; offset++) {
            const rowNumber = block.startRow + blockIndex * block.blockHeight + offset;
            const originalRowXml = extractRowXml(sheetXml, rowNumber);
            if (!originalRowXml) continue;

            let rowXml = originalRowXml;

            block.placeholders.forEach((ph) => {
              const rawValue = getNestedValue(item, ph.cellPath);
              const stringValue = stringifyPrimitiveValue(rawValue);
              const indices = placeholderToIndices.get(ph.placeholderKey);
              if (!indices) return;

              indices.forEach((oldIndex) => {
                const added = addSharedString(sharedStringsXml, stringValue, firstRunPropsByIndex.get(oldIndex));
                sharedStringsXml = added.updatedXml;
                rowXml = replaceSharedStringIndexInRow(rowXml, rowNumber, oldIndex, added.newIndex);
              });
            });

            sheetXml = sheetXml.replace(originalRowXml, rowXml);
          }
        }

        zip.file(block.sheetPath, sheetXml);

        if (insertedRows > 0 && workbookXmlForPrintArea) {
          workbookXmlForPrintArea = updatePrintAreaForSheet(
            workbookXmlForPrintArea,
            block.sheetName,
            block.sheetIndex,
            insertedRows
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

    // Stage B: 通常プレースホルダー（先頭run書式を維持して sharedStrings を置換）
    const primitiveReplacements = new Map<string, string>();
    Object.entries(data).forEach(([key, value]) => {
      if (Array.isArray(value) || isPlainObject(value)) return;
      primitiveReplacements.set(key, stringifyPrimitiveValue(value));
    });

    const replacedSharedStrings = replacePrimitivePlaceholdersWithFirstRunStyle(
      sharedStringsXml,
      primitiveReplacements
    );

    // Stage C: 旧記法でのPrint_Area更新（sheet1想定の既存互換）
    if (legacyInsertedRows > 0 && !workbookXmlForPrintArea) {
      const workbookFile = zip.file('xl/workbook.xml');
      if (workbookFile) {
        const workbookXml = await workbookFile.async('string');
        const currentPrintArea = getPrintAreaFromWorkbookXml(workbookXml);
        const parsed = currentPrintArea ? parseFirstPrintArea(currentPrintArea) : null;

        if (parsed) {
          const baseRows = parsed.endRow - parsed.startRow + 1;
          const pages = Math.ceil((baseRows + legacyInsertedRows) / baseRows);
          if (pages > 1) {
            const updatedWorkbookXml = workbookXml.replace(
              /(<definedName[^>]*name="_xlnm\.Print_Area"[^>]*>)([^<]*)(<\/definedName>)/g,
              (_m, p1, _p2, p3) => `${p1}${buildPagedPrintAreas({
                sheetPrefix: parsed.sheetPrefix,
                startCol: parsed.startCol,
                startRow: parsed.startRow,
                endCol: parsed.endCol,
                pageHeight: baseRows,
                pages,
              })}${p3}`
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
