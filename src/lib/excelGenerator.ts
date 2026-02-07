import { PlaceholderReplacer, PlaceholderData } from './placeholderReplacer';

export interface ExcelGenerationOptions {
  /**
   * 特定のシートのみを残す（指定しない場合は全シート）
   * NOTE: シート削除にはExcelJSを使用するため、一部の設定が変化する可能性があります。
   */
  sheetName?: string;
  sheetId?: number;
}

export class ExcelGenerator {
  private placeholderReplacer: PlaceholderReplacer;

  constructor() {
    this.placeholderReplacer = new PlaceholderReplacer();
  }

  /**
   * エクセル帳票を生成（test9方式）
   * sharedStrings.xmlを直接編集することで印刷設定を完全保持
   */
  async generateExcel(
    templateBase64: string,
    data: PlaceholderData,
    options: ExcelGenerationOptions = {}
  ): Promise<Buffer> {
    // Base64をBufferに変換
    const templateBuffer = Buffer.from(templateBase64, 'base64');

    // プレースホルダーを置換（印刷設定は完全保持される）
    const resultBuffer = await this.placeholderReplacer.replacePlaceholders(
      templateBuffer,
      data
    );

    // 特定のシートのみを残す場合（ExcelJSが必要）
    if (options.sheetName || options.sheetId !== undefined) {
      const ExcelJS = require('exceljs');
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(resultBuffer as any);

      const targetSheet =
        options.sheetId !== undefined
          ? workbook.getWorksheet(options.sheetId)
          : workbook.getWorksheet(options.sheetName);

      if (!targetSheet) {
        const selector = options.sheetId !== undefined ? `id=${options.sheetId}` : `name=${options.sheetName}`;
        throw new Error(`Worksheet not found (${selector})`);
      }

      const sheetsToRemove = workbook.worksheets.filter(
        (sheet: any) => sheet.id !== targetSheet.id
      );
      sheetsToRemove.forEach((sheet: any) => {
        workbook.removeWorksheet(sheet.id);
      });

      const filteredBuffer = await workbook.xlsx.writeBuffer();
      return Buffer.from(filteredBuffer);
    }

    return resultBuffer;
  }

  /**
   * エクセル帳票をBase64形式で生成
   */
  async generateExcelAsBase64(
    templateBase64: string,
    data: PlaceholderData,
    options: ExcelGenerationOptions = {}
  ): Promise<string> {
    const buffer = await this.generateExcel(templateBase64, data, options);
    return buffer.toString('base64');
  }

  /**
   * テンプレート内のプレースホルダーを検出
   */
  async findPlaceholders(templateBase64: string): Promise<string[]> {
    const templateBuffer = Buffer.from(templateBase64, 'base64');
    return this.placeholderReplacer.findPlaceholders(templateBuffer);
  }

  /**
   * テンプレート内のプレースホルダー詳細情報を取得
   */
  async getPlaceholderInfo(templateBase64: string) {
    const templateBuffer = Buffer.from(templateBase64, 'base64');
    return this.placeholderReplacer.getPlaceholderInfo(templateBuffer);
  }

  /**
   * テンプレートの基本情報を取得
   * （簡易版 - ZIP構造から取得）
   */
  async getTemplateInfo(templateBase64: string) {
    const JSZip = require('jszip');
    const templateBuffer = Buffer.from(templateBase64, 'base64');
    const zip = await JSZip.loadAsync(templateBuffer);

    // ワークシート数を取得
    const worksheetFiles = Object.keys(zip.files).filter(
      (filename) => filename.startsWith('xl/worksheets/sheet') && filename.endsWith('.xml')
    );

    const sheets = await Promise.all(
      worksheetFiles.map(async (filename, index) => {
        const content = await zip.file(filename)?.async('string');

        // シート名を取得（workbook.xmlから）
        const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
        const sheetNameMatch = workbookXml?.match(
          new RegExp(`<sheet[^>]*sheetId="${index + 1}"[^>]*name="([^"]*)"`)
        );
        const sheetName = sheetNameMatch?.[1] || `Sheet${index + 1}`;

        // 行数・列数の概算
        const dimensionMatch = content?.match(/<dimension[^>]*ref="([^"]*)"/)
        const dimension = dimensionMatch?.[1] || 'A1';

        return {
          id: index + 1,
          name: sheetName,
          rowCount: 0, // 簡易版では省略
          columnCount: 0, // 簡易版では省略
        };
      })
    );

    return {
      sheetCount: worksheetFiles.length,
      sheets,
    };
  }

  /**
   * エクセル帳票を生成し、ストリームで返す
   */
  async generateExcelStream(
    templateBase64: string,
    data: PlaceholderData
  ): Promise<NodeJS.ReadableStream> {
    const buffer = await this.generateExcel(templateBase64, data);

    const { PassThrough } = require('stream');
    const stream = new PassThrough();
    stream.end(buffer);

    return stream;
  }
}
