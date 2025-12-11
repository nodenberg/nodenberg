import ExcelJS from 'exceljs';
import { ExcelTemplateManager } from './excelTemplateManager';
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
  private templateManager: ExcelTemplateManager;
  private placeholderReplacer: PlaceholderReplacer;
  private sofficeConverter: SofficeConverter;

  constructor(sofficeOptions?: SofficeConversionOptions) {
    this.templateManager = new ExcelTemplateManager();
    this.placeholderReplacer = new PlaceholderReplacer();
    this.sofficeConverter = new SofficeConverter(sofficeOptions);
  }

  /**
   * エクセル帳票からPDFを生成（出力パターン2）
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

    // テンプレートを読み込む
    const workbook = await this.templateManager.loadWorkbookFromBase64(templateBase64);

    // プレースホルダーを置換
    await this.placeholderReplacer.replacePlaceholders(workbook, data);

    // 特定のシートのみを処理する場合
    if (options.sheetName) {
      // 指定されたシート以外を削除
      const targetSheet = workbook.getWorksheet(options.sheetName);
      if (!targetSheet) {
        throw new Error(`Worksheet "${options.sheetName}" not found`);
      }

      // 全シートを取得して、指定されたシート以外を削除
      const sheetsToRemove = workbook.worksheets.filter(
        (sheet) => sheet.name !== options.sheetName
      );

      sheetsToRemove.forEach((sheet) => {
        workbook.removeWorksheet(sheet.id);
      });
    }

    // ワークブックをBufferに変換
    const excelBuffer = await workbook.xlsx.writeBuffer();

    // sofficeでPDFに変換
    const pdfBuffer = await this.sofficeConverter.convertExcelToPDF(
      Buffer.from(excelBuffer)
    );

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
