import { PlaceholderReplacer, PlaceholderData } from './placeholderReplacer';
import { SofficeConverter, SofficeConversionOptions } from './sofficeConverter';

export interface PDFGenerationOptions {
  /**
   * LibreOfficeのsofficeコマンドのパス
   */
  sofficeCommand?: string;

  /**
   * 一時ファイルを保存するディレクトリ
   */
  tempDir?: string;

  /**
   * タイムアウト時間（ミリ秒）
   */
  timeout?: number;

  /**
   * 特定のシートのみをPDFに変換（指定しない場合は全シート）
   */
  sheetName?: string;
}

export class PDFGenerator {
  private placeholderReplacer: PlaceholderReplacer;
  private sofficeConverter: SofficeConverter;

  constructor(sofficeOptions?: SofficeConversionOptions) {
    this.placeholderReplacer = new PlaceholderReplacer();
    this.sofficeConverter = new SofficeConverter(sofficeOptions);
  }

  /**
   * エクセル帳票からPDFを生成（test9方式）
   * LibreOfficeのsoffice headlessを使用してExcelの見た目そのままPDF化
   */
  async generatePDF(
    templateBase64: string,
    data: PlaceholderData,
    options: PDFGenerationOptions = {}
  ): Promise<Buffer> {
    // sofficeコマンドのパスが指定されている場合は設定
    if (options.sofficeCommand) {
      this.sofficeConverter.setSofficeCommand(options.sofficeCommand);
    }

    // Base64をBufferに変換
    const templateBuffer = Buffer.from(templateBase64, 'base64');

    // プレースホルダーを置換（test9方式 - 印刷設定完全保持）
    const excelBuffer = await this.placeholderReplacer.replacePlaceholders(
      templateBuffer,
      data
    );

    // 特定のシートのみを処理する場合
    // 注意: test9方式ではシート削除にExcelJSが必要
    if (options.sheetName) {
      // ExcelJSを使用してシート削除
      const ExcelJS = require('exceljs');
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(excelBuffer);

      // 指定されたシート以外を削除
      const targetSheet = workbook.getWorksheet(options.sheetName);
      if (!targetSheet) {
        throw new Error(`Worksheet "${options.sheetName}" not found`);
      }

      // 全シートを取得して、指定されたシート以外を削除
      const sheetsToRemove = workbook.worksheets.filter(
        (sheet: any) => sheet.name !== options.sheetName
      );

      sheetsToRemove.forEach((sheet: any) => {
        workbook.removeWorksheet(sheet.id);
      });

      // 再度Bufferに変換
      const filteredBuffer = await workbook.xlsx.writeBuffer();

      // sofficeでPDFに変換
      const pdfBuffer = await this.sofficeConverter.convertExcelToPDF(
        Buffer.from(filteredBuffer)
      );

      return pdfBuffer;
    }

    // sofficeでPDFに変換
    const pdfBuffer = await this.sofficeConverter.convertExcelToPDF(excelBuffer);

    return pdfBuffer;
  }

  /**
   * PDFをBase64形式で生成
   */
  async generatePDFAsBase64(
    templateBase64: string,
    data: PlaceholderData,
    options: PDFGenerationOptions = {}
  ): Promise<string> {
    const buffer = await this.generatePDF(templateBase64, data, options);
    return buffer.toString('base64');
  }

  /**
   * LibreOfficeがインストールされているか確認
   */
  async checkLibreOfficeInstalled(): Promise<boolean> {
    return await this.sofficeConverter.checkSofficeInstalled();
  }

  /**
   * LibreOfficeのバージョンを取得
   */
  async getLibreOfficeVersion(): Promise<string> {
    return await this.sofficeConverter.getLibreOfficeVersion();
  }

  /**
   * sofficeコマンドのパスを設定
   */
  setSofficeCommand(command: string): void {
    this.sofficeConverter.setSofficeCommand(command);
  }

  /**
   * 現在のsofficeコマンドのパスを取得
   */
  getSofficeCommand(): string {
    return this.sofficeConverter.getSofficeCommand();
  }
}
