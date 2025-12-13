import JSZip from 'jszip';

export interface PlaceholderData {
  [key: string]: string | number | Date | null;
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

export class PlaceholderReplacer {
  private placeholderPattern = /\{\{([^}]+)\}\}/g;

  /**
   * Excelファイル（Buffer）内のプレースホルダーを置換
   * sharedStrings.xmlを直接編集することで印刷設定を完全保持
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

    // 3. プレースホルダー置換（自動エスケープ）
    let replacedSharedStrings = sharedStringsXml;

    Object.entries(data).forEach(([key, value]) => {
      const placeholder = `{{${key}}}`;

      // 値を文字列に変換
      let stringValue: string;
      if (value === null || value === undefined) {
        stringValue = '';
      } else if (value instanceof Date) {
        stringValue = this.formatDate(value);
      } else {
        stringValue = String(value);
      }

      // XMLエスケープ
      const escapedValue = escapeXml(stringValue);

      // プレースホルダーを検索して置換
      const regex = new RegExp(escapeRegExp(placeholder), 'g');
      replacedSharedStrings = replacedSharedStrings.replace(regex, escapedValue);
    });

    // 4. ZIPに書き戻し
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
