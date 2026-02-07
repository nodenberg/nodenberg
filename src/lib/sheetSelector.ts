import JSZip from 'jszip';

export interface SheetSelectionOptions {
  sheetName?: string;
  sheetId?: number;
}

type SheetEntry = {
  raw: string;
  name: string;
  sheetId: number;
  relId: string;
  index: number;
};

function getAttr(tag: string, attrName: string): string | null {
  const escaped = attrName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const m = tag.match(new RegExp(`${escaped}="([^"]*)"`, 'i'));
  return m ? m[1] : null;
}

function parseSheets(workbookXml: string): SheetEntry[] {
  const sheetsMatch = workbookXml.match(/<sheets>([\s\S]*?)<\/sheets>/i);
  if (!sheetsMatch) return [];

  const sheetTags = sheetsMatch[1].match(/<sheet\b[^>]*\/>/gi) || [];
  const sheets: SheetEntry[] = [];

  sheetTags.forEach((tag, index) => {
    const name = getAttr(tag, 'name');
    const sheetId = getAttr(tag, 'sheetId');
    const relId = getAttr(tag, 'r:id');
    if (!name || !sheetId || !relId) return;
    sheets.push({
      raw: tag,
      name,
      sheetId: Number(sheetId),
      relId,
      index,
    });
  });

  return sheets;
}

function findWorksheetTargetByRelId(workbookRelsXml: string, relId: string): string | null {
  const relTags = workbookRelsXml.match(/<Relationship\b[^>]*\/>/gi) || [];
  for (const relTag of relTags) {
    const id = getAttr(relTag, 'Id');
    if (id !== relId) continue;
    const target = getAttr(relTag, 'Target');
    if (!target) return null;
    return target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\/+/, '')}`;
  }
  return null;
}

function normalizeWorksheetPath(target: string): string {
  if (target.startsWith('xl/')) return target;
  if (target.startsWith('/xl/')) return target.slice(1);
  if (target.startsWith('worksheets/')) return `xl/${target}`;
  return `xl/${target.replace(/^\/+/, '')}`;
}

function keepOnlyTargetSheetInWorkbookXml(workbookXml: string, target: SheetEntry): string {
  const sheetsMatch = workbookXml.match(/<sheets>[\s\S]*?<\/sheets>/i);
  if (!sheetsMatch) throw new Error('Invalid workbook.xml: <sheets> not found');

  const newSheetsXml = `<sheets>${target.raw}</sheets>`;
  let updated = workbookXml.replace(sheetsMatch[0], newSheetsXml);

  // Single-sheet workbook should point to the first sheet tab.
  updated = updated.replace(/<workbookView\b([^>]*)\/>/i, (_whole, attrs: string) => {
    let next = attrs;
    if (/\bactiveTab="/i.test(next)) {
      next = next.replace(/\bactiveTab="[^"]*"/i, 'activeTab="0"');
    } else {
      next += ' activeTab="0"';
    }
    if (/\bfirstSheet="/i.test(next)) {
      next = next.replace(/\bfirstSheet="[^"]*"/i, 'firstSheet="0"');
    } else {
      next += ' firstSheet="0"';
    }
    return `<workbookView${next}/>`;
  });

  // Keep only defined names that are global or belong to the selected sheet.
  const definedNamesMatch = updated.match(/<definedNames>([\s\S]*?)<\/definedNames>/i);
  if (definedNamesMatch) {
    const definedNameTags = definedNamesMatch[1].match(/<definedName\b[\s\S]*?<\/definedName>/gi) || [];
    const kept = definedNameTags.filter((tag) => {
      const localSheetId = getAttr(tag, 'localSheetId');
      if (localSheetId === null) return true; // global name
      return Number(localSheetId) === target.index;
    }).map((tag) => {
      if (/\blocalSheetId="/i.test(tag)) {
        return tag.replace(/\blocalSheetId="[^"]*"/i, 'localSheetId="0"');
      }
      return tag;
    });

    if (kept.length === 0) {
      updated = updated.replace(definedNamesMatch[0], '');
    } else {
      updated = updated.replace(definedNamesMatch[0], `<definedNames>${kept.join('')}</definedNames>`);
    }
  }

  return updated;
}

function keepOnlyTargetSheetInWorkbookRelsXml(workbookRelsXml: string, targetRelId: string): string {
  const relationshipTags = workbookRelsXml.match(/<Relationship\b[^>]*\/>/gi) || [];
  const kept = relationshipTags.filter((tag) => {
    const type = getAttr(tag, 'Type') || '';
    const id = getAttr(tag, 'Id') || '';
    const isWorksheetRel = /\/worksheet$/i.test(type);
    if (!isWorksheetRel) return true;
    return id === targetRelId;
  });

  return workbookRelsXml.replace(
    /<Relationships[^>]*>[\s\S]*<\/Relationships>/i,
    (whole) => {
      const openTag = whole.match(/<Relationships[^>]*>/i)?.[0] || '<Relationships>';
      return `${openTag}${kept.join('')}</Relationships>`;
    }
  );
}

function keepOnlyTargetSheetInContentTypesXml(contentTypesXml: string, keptWorksheetPath: string): string {
  const normalizedKept = `/${keptWorksheetPath.replace(/^\/+/, '')}`;
  const overrideTags = contentTypesXml.match(/<Override\b[^>]*\/>/gi) || [];
  const keptOverrides = overrideTags.filter((tag) => {
    const partName = getAttr(tag, 'PartName');
    if (!partName) return true;
    const normalizedPart = partName.startsWith('/') ? partName : `/${partName}`;
    if (!normalizedPart.startsWith('/xl/worksheets/')) return true;
    return normalizedPart === normalizedKept;
  });

  return contentTypesXml.replace(
    /<Types[^>]*>[\s\S]*<\/Types>/i,
    (whole) => {
      const openTag = whole.match(/<Types[^>]*>/i)?.[0] || '<Types>';
      const defaultTags = whole.match(/<Default\b[^>]*\/>/gi) || [];
      return `${openTag}${defaultTags.join('')}${keptOverrides.join('')}</Types>`;
    }
  );
}

export async function selectSingleSheetFromWorkbookBuffer(
  workbookBuffer: Buffer,
  options: SheetSelectionOptions
): Promise<Buffer> {
  if (!options.sheetName && options.sheetId === undefined) return workbookBuffer;

  const zip = await JSZip.loadAsync(workbookBuffer);
  const workbookXmlFile = zip.file('xl/workbook.xml');
  const workbookRelsFile = zip.file('xl/_rels/workbook.xml.rels');
  if (!workbookXmlFile || !workbookRelsFile) {
    throw new Error('Invalid workbook: workbook.xml or workbook.xml.rels not found');
  }

  const workbookXml = await workbookXmlFile.async('string');
  const workbookRelsXml = await workbookRelsFile.async('string');

  const sheets = parseSheets(workbookXml);
  if (sheets.length === 0) throw new Error('No worksheets found in workbook');

  const targetSheet = options.sheetId !== undefined
    ? sheets.find((s) => s.sheetId === options.sheetId)
    : sheets.find((s) => s.name === options.sheetName);

  if (!targetSheet) {
    const selector = options.sheetId !== undefined ? `id=${options.sheetId}` : `name=${options.sheetName}`;
    throw new Error(`Worksheet not found (${selector})`);
  }

  const targetPathRaw = findWorksheetTargetByRelId(workbookRelsXml, targetSheet.relId);
  if (!targetPathRaw) {
    throw new Error(`Worksheet relationship not found (r:id=${targetSheet.relId})`);
  }
  const keptWorksheetPath = normalizeWorksheetPath(targetPathRaw);

  const updatedWorkbookXml = keepOnlyTargetSheetInWorkbookXml(workbookXml, targetSheet);
  const updatedWorkbookRelsXml = keepOnlyTargetSheetInWorkbookRelsXml(workbookRelsXml, targetSheet.relId);
  zip.file('xl/workbook.xml', updatedWorkbookXml);
  zip.file('xl/_rels/workbook.xml.rels', updatedWorkbookRelsXml);

  const contentTypesFile = zip.file('[Content_Types].xml');
  if (contentTypesFile) {
    const contentTypesXml = await contentTypesFile.async('string');
    const updatedContentTypesXml = keepOnlyTargetSheetInContentTypesXml(contentTypesXml, keptWorksheetPath);
    zip.file('[Content_Types].xml', updatedContentTypesXml);
  }

  // Remove worksheet XML files other than the selected sheet, while keeping target XML untouched.
  const worksheetFiles = Object.keys(zip.files).filter((p) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(p));
  worksheetFiles.forEach((path) => {
    if (path !== keptWorksheetPath) zip.remove(path);
  });

  // Remove worksheet rel files that correspond to removed sheets.
  const worksheetRelFiles = Object.keys(zip.files).filter((p) => /^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/i.test(p));
  worksheetRelFiles.forEach((path) => {
    const baseSheetPath = path
      .replace(/^xl\/worksheets\/_rels\//, 'xl/worksheets/')
      .replace(/\.rels$/, '');
    if (baseSheetPath !== keptWorksheetPath) zip.remove(path);
  });

  const output = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 },
  });
  return Buffer.from(output);
}
