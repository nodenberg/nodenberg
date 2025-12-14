import JSZip from 'jszip';

export interface PlaceholderData {
  [key: string]:
    | string
    | number
    | Date
    | null
    | Array<Record<string, unknown>>;
}

/**
 * XML特殊文字をエスケープ（W3C準拠）
 */
function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')    // & を最初に処理（重要）
    .replace(/</g, '&lt;')     // < を変換
    .replace(/>/g, '&gt;')     // > を変換
    .replace(/"/g, '&quot;')   // " を変換
    .replace(/'/g, '&apos;');  // ' を変換
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
  let match;

  while ((match = regex.exec(xml)) !== null) {
    const placeholder = match[0];
    const key = match[1];

    if (placeholderMap.has(placeholder)) {
      placeholderMap.get(placeholder)!.count++;
    } else {
      placeholderMap.set(placeholder, {
        placeholder: placeholder,
        key: key,
        count: 1
      });
    }
  }

  return Array.from(placeholderMap.values());
}

type ArrayPlaceholder = {
  placeholderKey: string; // 例: "#明細.番号"
  arrayName: string;      // 例: "明細"
  fieldPath: string;      // 例: "番号"
};

function parseArrayPlaceholderKey(key: string): ArrayPlaceholder | null {
  if (!key.startsWith('#')) return null;
  const cleanName = key.substring(1);
  const parts = cleanName.split('.');
  if (parts.length < 2) return null;
  return {
    placeholderKey: key,
    arrayName: parts[0],
    fieldPath: parts.slice(1).join('.'),
  };
}

function extractSharedStrings(sharedStringsXml: string): string[] {
  const strings: string[] = [];
  const siRegex = /<si>(.*?)<\/si>/gs;
  let match: RegExpExecArray | null;
  while ((match = siRegex.exec(sharedStringsXml)) !== null) {
    const siContent = match[1];
    const tMatch = siContent.match(/<t[^>]*>(.*?)<\/t>/s);
    if (tMatch) strings.push(tMatch[1]);
  }
  return strings;
}

function findCellsWithArrayPlaceholderIndices(
  sheetXml: string,
  arrayNameToIndices: Map<string, Set<number>>
): Map<string, Set<number>> {
  const result = new Map<string, Set<number>>();
  const cellRegex = /<c r="([A-Z]+)(\d+)"[^>]*><v>(\d+)<\/v><\/c>/g;
  let match: RegExpExecArray | null;
  while ((match = cellRegex.exec(sheetXml)) !== null) {
    const row = parseInt(match[2], 10);
    const sharedStringIndex = parseInt(match[3], 10);
    arrayNameToIndices.forEach((indices, arrayName) => {
      if (indices.has(sharedStringIndex)) {
        if (!result.has(arrayName)) result.set(arrayName, new Set());
        result.get(arrayName)!.add(row);
      }
    });
  }
  return result;
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

function shiftRowsDown(sheetXml: string, fromRow: number, shiftAmount: number): string {
  const rowRegex = /<row[^>]*r="(\d+)"[^>]*>.*?<\/row>/gs;
  let result = sheetXml;
  const matches: Array<{ rowNum: number; xml: string }> = [];

  let match: RegExpExecArray | null;
  while ((match = rowRegex.exec(sheetXml)) !== null) {
    const rowNum = parseInt(match[1], 10);
    if (rowNum >= fromRow) matches.push({ rowNum, xml: match[0] });
  }

  matches.sort((a, b) => b.rowNum - a.rowNum);
  matches.forEach(({ rowNum, xml }) => {
    const newRowNum = rowNum + shiftAmount;
    const updatedXml = updateRowNumber(xml, newRowNum);
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
  let match: RegExpExecArray | null;
  const updates: Array<{ oldRef: string; newRef: string }> = [];

  while ((match = mergeCellRegex.exec(mergeCellsContent)) !== null) {
    const startCol = match[1];
    const startRow = parseInt(match[2], 10);
    const endCol = match[3];
    const endRow = parseInt(match[4], 10);

    if (startRow >= fromRow || endRow >= fromRow) {
      const newStartRow = startRow >= fromRow ? startRow + shiftAmount : startRow;
      const newEndRow = endRow >= fromRow ? endRow + shiftAmount : endRow;
      const oldRef = `${startCol}${startRow}:${endCol}${endRow}`;
      const newRef = `${startCol}${newStartRow}:${endCol}${newEndRow}`;
      updates.push({ oldRef, newRef });
    }
  }

  updates.sort((a, b) => {
    const aRow = parseInt(a.oldRef.match(/\d+/)![0], 10);
    const bRow = parseInt(b.oldRef.match(/\d+/)![0], 10);
    return bRow - aRow;
  });

  let result = sheetXml;
  updates.forEach(({ oldRef, newRef }) => {
    result = result.replace(`<mergeCell ref="${oldRef}"/>`, `<mergeCell ref="${newRef}"/>`);
  });
  return result;
}

function insertMergeCellsForNewRows(sheetXml: string, templateRow: number, newRows: number[]): string {
  const mergeCellsRegex = /<mergeCells[^>]*>(.*?)<\/mergeCells>/s;
  const mergeCellsMatch = sheetXml.match(mergeCellsRegex);
  if (!mergeCellsMatch) return sheetXml;

  const mergeCellsContent = mergeCellsMatch[1];
  const mergeCellRegex = /<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/g;
  let match: RegExpExecArray | null;
  const templateMergeCells: Array<{ startCol: string; endCol: string }> = [];

  while ((match = mergeCellRegex.exec(mergeCellsContent)) !== null) {
    const startCol = match[1];
    const startRow = parseInt(match[2], 10);
    const endCol = match[3];
    const endRow = parseInt(match[4], 10);

    if (startRow === templateRow && endRow === templateRow) {
      templateMergeCells.push({ startCol, endCol });
    }
  }

  if (templateMergeCells.length === 0) return sheetXml;

  const newMergeCells: string[] = [];
  newRows.forEach(rowNumber => {
    templateMergeCells.forEach(({ startCol, endCol }) => {
      newMergeCells.push(`<mergeCell ref="${startCol}${rowNumber}:${endCol}${rowNumber}"/>`);
    });
  });

  const result = sheetXml.replace('</mergeCells>', newMergeCells.join('') + '</mergeCells>');
  const countMatch = sheetXml.match(/<mergeCells count="(\d+)"/);
  if (!countMatch) return result;

  const currentCount = parseInt(countMatch[1], 10);
  const newCount = currentCount + newMergeCells.length;
  return result.replace(/(<mergeCells count=")(\d+)(")/, `$1${newCount}$3`);
}

function addSharedString(sharedStringsXml: string, newString: string) {
  const escapedString = escapeXml(newString);
  const countMatch = sharedStringsXml.match(/<sst[^>]*count="(\d+)"/);
  const uniqueCountMatch = sharedStringsXml.match(/<sst[^>]*uniqueCount="(\d+)"/);
  const currentCount = countMatch ? parseInt(countMatch[1], 10) : 0;
  const currentUniqueCount = uniqueCountMatch ? parseInt(uniqueCountMatch[1], 10) : 0;
  const newIndex = currentUniqueCount;

  let updatedXml = sharedStringsXml.replace(/count="\d+"/, `count="${currentCount + 1}"`);
  updatedXml = updatedXml.replace(/uniqueCount="\d+"/, `uniqueCount="${currentUniqueCount + 1}"`);
  const newSi = `<si><t>${escapedString}</t></si>`;
  updatedXml = updatedXml.replace('</sst>', `${newSi}</sst>`);

  return { updatedXml, newIndex };
}

function getPrintAreaFromWorkbookXml(workbookXml: string): string | null {
  const m = workbookXml.match(/name="_xlnm\.Print_Area"[^>]*>([^<]+)<\/definedName>/);
  return m ? m[1] : null;
}

function parseFirstPrintArea(printAreaValue: string) {
  const first = printAreaValue.split(',')[0];
  const bangIndex = first.indexOf('!');
  if (bangIndex === -1) return null;

  const sheetPrefix = first.slice(0, bangIndex); // &apos;Sheet&apos;
  const rangePart = first.slice(bangIndex + 1); // $A$1:$Q$40

  const m = rangePart.match(/\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)/);
  if (!m) return null;

  return {
    sheetPrefix,
    startCol: m[1],
    startRow: parseInt(m[2], 10),
    endCol: m[3],
    endRow: parseInt(m[4], 10),
  };
}

function buildPagedPrintAreas(params: {
  sheetPrefix: string;
  startCol: string;
  startRow: number;
  endCol: string;
  pageHeight: number;
  pages: number;
}) {
  const ranges: string[] = [];
  for (let pageIndex = 0; pageIndex < params.pages; pageIndex++) {
    const pageStartRow = params.startRow + pageIndex * params.pageHeight;
    const pageEndRow = pageStartRow + params.pageHeight - 1;
    ranges.push(`${params.sheetPrefix}!$${params.startCol}$${pageStartRow}:$${params.endCol}$${pageEndRow}`);
  }
  return ranges.join(',');
}

function updatePrintAreaToPaged(workbookXml: string, pages: number): string {
  const current = getPrintAreaFromWorkbookXml(workbookXml);
  if (!current) return workbookXml;

  const first = parseFirstPrintArea(current);
  if (!first) return workbookXml;

  const pageHeight = first.endRow - first.startRow + 1;
  if (pageHeight <= 0) return workbookXml;

  const newValue = buildPagedPrintAreas({
    sheetPrefix: first.sheetPrefix,
    startCol: first.startCol,
    startRow: first.startRow,
    endCol: first.endCol,
    pageHeight,
    pages,
  });

  return workbookXml.replace(/(<definedName[^>]*name="_xlnm\.Print_Area"[^>]*>)([^<]*)(<\/definedName>)/g, (_m, p1, _p2, p3) => {
    return `${p1}${newValue}${p3}`;
  });
}

export class PlaceholderReplacer {
  private placeholderPattern = /\{\{([^}]+)\}\}/g;

  /**
   * Excelファイル（Buffer）内のプレースホルダーを置換
   * sharedStrings.xmlを直接編集することで印刷設定を保持しつつ、
   * 配列プレースホルダーがある場合は明細行の増加に合わせて行挿入 + Print_Area 追加（test13方式）も行う
   */
  async replacePlaceholders(
    excelBuffer: Buffer,
    data: PlaceholderData
  ): Promise<Buffer> {
    // 1. ZIPとして読み込み
    const zip = await JSZip.loadAsync(excelBuffer);

    // 2. sharedStrings.xmlを取得（プレースホルダーはここにある）
    const sharedStringsFile = zip.file('xl/sharedStrings.xml');
    if (!sharedStringsFile) {
      throw new Error('sharedStrings.xmlが見つかりません');
    }

    let sharedStringsXml = await sharedStringsFile.async('string');

    // ===== Stage A: 配列プレースホルダーがあれば、行挿入 + セル参照差し替え =====
    const placeholderInfo = detectPlaceholdersInXml(sharedStringsXml);
    const arrayPlaceholders = placeholderInfo
      .map(p => parseArrayPlaceholderKey(p.key))
      .filter((p): p is ArrayPlaceholder => p !== null);

    let insertCount = 0;

    if (arrayPlaceholders.length > 0) {
      // 現状はテンプレート（請求書）想定で sheet1 を対象にする
      const sheetFile = zip.file('xl/worksheets/sheet1.xml');
      if (!sheetFile) throw new Error('sheet1.xmlが見つかりません（配列展開に必要）');
      let sheetXml = await sheetFile.async('string');

      // sharedStrings の index を特定（placeholderKey -> sharedStringIndex）
      const sharedStrings = extractSharedStrings(sharedStringsXml);
      const placeholderToIndex = new Map<string, number>();
      const arrayNameToIndices = new Map<string, Set<number>>();

      arrayPlaceholders.forEach(ph => {
        const placeholderToken = `{{${ph.placeholderKey}}}`;
        sharedStrings.forEach((str, idx) => {
          if (str.includes(placeholderToken)) {
            placeholderToIndex.set(ph.placeholderKey, idx);
            if (!arrayNameToIndices.has(ph.arrayName)) arrayNameToIndices.set(ph.arrayName, new Set());
            arrayNameToIndices.get(ph.arrayName)!.add(idx);
          }
        });
      });

      // テンプレート上の行位置を取得（配列ごと）
      const arrayRows = findCellsWithArrayPlaceholderIndices(sheetXml, arrayNameToIndices);
      const firstArrayName = Array.from(arrayRows.keys())[0];
      const templateRowNumbers = Array.from(arrayRows.get(firstArrayName) || []).sort((a, b) => a - b);

      const arrayData = Array.isArray((data as Record<string, unknown>)[firstArrayName])
        ? ((data as Record<string, unknown>)[firstArrayName] as Array<Record<string, unknown>>)
        : null;

      if (arrayData && templateRowNumbers.length > 0) {
        const templateCapacity = templateRowNumbers.length;
        insertCount = Math.max(0, arrayData.length - templateCapacity);

        if (insertCount > 0) {
          const lastTemplateRow = templateRowNumbers[templateRowNumbers.length - 1];
          const insertStartRow = lastTemplateRow + 1;

          sheetXml = shiftRowsDown(sheetXml, insertStartRow, insertCount);
          sheetXml = shiftMergeCellsDown(sheetXml, insertStartRow, insertCount);

          const templateRowXml = extractRowXml(sheetXml, lastTemplateRow);
          if (!templateRowXml) throw new Error(`テンプレート行のXMLが取得できません: row=${lastTemplateRow}`);

          const newRowsXml: string[] = [];
          for (let i = 0; i < insertCount; i++) {
            const newRowNumber = insertStartRow + i;
            newRowsXml.push(updateRowNumber(templateRowXml, newRowNumber));
          }

          const lastTemplateRowXml = extractRowXml(sheetXml, lastTemplateRow);
          if (!lastTemplateRowXml) throw new Error(`テンプレート行のXMLが取得できません: row=${lastTemplateRow}`);
          sheetXml = sheetXml.replace(lastTemplateRowXml, lastTemplateRowXml + newRowsXml.join(''));

          const newRowNumbers = Array.from({ length: insertCount }, (_, i) => insertStartRow + i);
          sheetXml = insertMergeCellsForNewRows(sheetXml, lastTemplateRow, newRowNumbers);
        }

        // 配列データをセルに反映（sharedStrings を追加して index を差し替える）
        // データがテンプレート行数より少ない場合も、残りのテンプレート行は空文字に置換して
        // プレースホルダーが表示されないようにする（#VALUE! の原因にもなる）
        const totalRows = Math.max(arrayData.length, templateCapacity);
        for (let i = 0; i < totalRows; i++) {
          const rowNumber = templateRowNumbers[0] + i;
          const originalRowXml = extractRowXml(sheetXml, rowNumber);
          if (!originalRowXml) continue;
          let rowXml = originalRowXml;

          for (const ph of arrayPlaceholders) {
            if (ph.arrayName !== firstArrayName) continue;
            const oldIndex = placeholderToIndex.get(ph.placeholderKey);
            if (oldIndex === undefined) continue;

            const item = i < arrayData.length ? (arrayData[i] || {}) : {};
            const rawValue = (item as Record<string, unknown>)[ph.fieldPath];
            const stringValue = rawValue === null || rawValue === undefined ? '' : String(rawValue);

            const added = addSharedString(sharedStringsXml, stringValue);
            sharedStringsXml = added.updatedXml;

            const cellRegex = new RegExp(`<c r="[A-Z]+${rowNumber}"[^>]*><v>${oldIndex}</v><\\/c>`, 'g');
            const updatedRowXml = rowXml.replace(cellRegex, (match) => {
              return match.replace(`<v>${oldIndex}</v>`, `<v>${added.newIndex}</v>`);
            });
            rowXml = updatedRowXml;
          }

          sheetXml = sheetXml.replace(originalRowXml, rowXml);
        }

        zip.file('xl/worksheets/sheet1.xml', sheetXml);
      }
    }

    // ===== Stage B: 通常プレースホルダー（sharedStrings の置換） =====
    let replacedSharedStrings = sharedStringsXml;

    Object.entries(data).forEach(([key, value]) => {
      if (Array.isArray(value)) return;

      const placeholder = `{{${key}}}`;

      let stringValue: string;
      if (value === null || value === undefined) {
        stringValue = '';
      } else if (value instanceof Date) {
        stringValue = this.formatDate(value);
      } else {
        stringValue = String(value);
      }

      const escapedValue = escapeXml(stringValue);
      const regex = new RegExp(escapeRegExp(placeholder), 'g');
      replacedSharedStrings = replacedSharedStrings.replace(regex, escapedValue);
    });

    // ===== Stage C: はみ出しがあれば Print_Area を追加（test13方式） =====
    if (insertCount > 0) {
      const workbookFile = zip.file('xl/workbook.xml');
      if (workbookFile) {
        const workbookXml = await workbookFile.async('string');
        const currentPrintArea = getPrintAreaFromWorkbookXml(workbookXml);
        const parsed = currentPrintArea ? parseFirstPrintArea(currentPrintArea) : null;
        if (parsed) {
          const baseEndRow = parsed.endRow;
          const estimatedFinalEndRow = baseEndRow + insertCount;
          const pageHeight = baseEndRow - parsed.startRow + 1;
          const rowsNeeded = estimatedFinalEndRow - parsed.startRow + 1;
          const pages = Math.ceil(rowsNeeded / pageHeight);

          if (pages > 1) {
            const updatedWorkbookXml = updatePrintAreaToPaged(workbookXml, pages);
            zip.file('xl/workbook.xml', updatedWorkbookXml);
          }
        }
      }
    }

    // ZIPに書き戻し（sharedStrings）
    zip.file('xl/sharedStrings.xml', replacedSharedStrings);

    // 5. Bufferとして返す
    const outputBuffer = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 9 }
    });

    return outputBuffer;
  }

  /**
   * Excelファイル（Buffer）内のプレースホルダーを検出
   */
  async findPlaceholders(excelBuffer: Buffer): Promise<string[]> {
    // 1. ZIPとして読み込み
    const zip = await JSZip.loadAsync(excelBuffer);

    // 2. sharedStrings.xmlを取得
    const sharedStringsFile = zip.file('xl/sharedStrings.xml');
    if (!sharedStringsFile) {
      return [];
    }

    const sharedStringsXml = await sharedStringsFile.async('string');

    // 3. プレースホルダー検出
    const placeholders = detectPlaceholdersInXml(sharedStringsXml);

    // 4. キーのみを返す（重複を除去してソート）
    const keys = placeholders.map(p => p.key);
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
    // 1. ZIPとして読み込み
    const zip = await JSZip.loadAsync(excelBuffer);

    // 2. sharedStrings.xmlを取得
    const sharedStringsFile = zip.file('xl/sharedStrings.xml');
    if (!sharedStringsFile) {
      return [];
    }

    const sharedStringsXml = await sharedStringsFile.async('string');

    // 3. プレースホルダー検出
    return detectPlaceholdersInXml(sharedStringsXml);
  }

  /**
   * 日付をフォーマット（yyyy/MM/dd形式）
   */
  private formatDate(date: Date): string {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}/${month}/${day}`;
  }
}
