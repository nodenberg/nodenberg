import ExcelJS from 'exceljs';
import { ExcelTemplateManager } from './excelTemplateManager';
import { PlaceholderReplacer, PlaceholderData } from './placeholderReplacer';

export class ExcelGenerator {
  private templateManager: ExcelTemplateManager;
  private placeholderReplacer: PlaceholderReplacer;

  constructor() {
    this.templateManager = new ExcelTemplateManager();
    this.placeholderReplacer = new PlaceholderReplacer();
  }

  /**
   * エクセル帳票を生成（出力パターン1）
   * プレースホルダーを置換したエクセルファイルを生成
   */
  async generateExcel(
    templateBase64: string,
    data: PlaceholderData
  ): Promise<Buffer> {
    // テンプレートを読み込む
    const workbook = await this.templateManager.loadWorkbookFromBase64(templateBase64);

    // シート情報と印刷設定を読み取る（ログ出力用）
    const workbookInfo = this.templateManager.getWorkbookInfo(workbook);
    console.log('Workbook Info:', workbookInfo);

    workbook.eachSheet((worksheet) => {
      const printSettings = this.templateManager.getPrintSettings(worksheet);
      console.log(`Print Settings for ${worksheet.name}:`, printSettings);
    });

    // 印刷設定を保持
    const printSettingsMap = this.capturePrintSettings(workbook);

    // プレースホルダーを置換
    const processedWorkbook = await this.placeholderReplacer.replacePlaceholders(
      workbook,
      data
    );

    // 印刷設定を再適用
    this.restorePrintSettings(processedWorkbook, printSettingsMap);

    // Bufferとして出力
    const buffer = await processedWorkbook.xlsx.writeBuffer();
    return Buffer.from(buffer);
  }

  /**
   * エクセル帳票をBase64形式で生成
   */
  async generateExcelAsBase64(
    templateBase64: string,
    data: PlaceholderData
  ): Promise<string> {
    const buffer = await this.generateExcel(templateBase64, data);
    return buffer.toString('base64');
  }

  /**
   * テンプレート内のプレースホルダーを検出
   */
  async findPlaceholders(templateBase64: string): Promise<string[]> {
    const workbook = await this.templateManager.loadWorkbookFromBase64(templateBase64);
    return this.placeholderReplacer.findPlaceholders(workbook);
  }

  /**
   * テンプレート内のプレースホルダー詳細情報を取得
   */
  async getPlaceholderInfo(templateBase64: string) {
    const workbook = await this.templateManager.loadWorkbookFromBase64(templateBase64);
    return this.placeholderReplacer.getPlaceholderInfo(workbook);
  }

  /**
   * テンプレートの基本情報を取得
   */
  async getTemplateInfo(templateBase64: string) {
    const workbook = await this.templateManager.loadWorkbookFromBase64(templateBase64);
    return this.templateManager.getWorkbookInfo(workbook);
  }

  /**
   * 特定のシートの印刷設定を取得
   */
  async getPrintSettings(templateBase64: string, sheetName?: string) {
    const workbook = await this.templateManager.loadWorkbookFromBase64(templateBase64);

    if (sheetName) {
      const worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        throw new Error(`Worksheet "${sheetName}" not found`);
      }
      return this.templateManager.getPrintSettings(worksheet);
    }

    // 全シートの印刷設定を取得
    const printSettings: { [key: string]: any } = {};
    workbook.eachSheet((worksheet) => {
      printSettings[worksheet.name] = this.templateManager.getPrintSettings(worksheet);
    });

    return printSettings;
  }

  /**
   * エクセル帳票を生成し、ストリームで返す
   */
  async generateExcelStream(
    templateBase64: string,
    data: PlaceholderData
  ): Promise<NodeJS.ReadableStream> {
    const workbook = await this.templateManager.loadWorkbookFromBase64(templateBase64);

    // 印刷設定を保持
    const printSettingsMap = this.capturePrintSettings(workbook);

    // プレースホルダーを置換
    await this.placeholderReplacer.replacePlaceholders(workbook, data);

    // 印刷設定を再適用
    this.restorePrintSettings(workbook, printSettingsMap);

    // ストリームとして出力
    const stream = new (require('stream').PassThrough)();
    await workbook.xlsx.write(stream);
    return stream;
  }

  /**
   * ワークブックの印刷設定を保持
   */
  private capturePrintSettings(workbook: ExcelJS.Workbook): Map<number, any> {
    const settings = new Map<number, any>();
    workbook.eachSheet((worksheet) => {
      settings.set(worksheet.id, {
        pageSetup: JSON.parse(JSON.stringify(worksheet.pageSetup)),
        headerFooter: JSON.parse(JSON.stringify(worksheet.headerFooter)),
      });
    });
    return settings;
  }

  /**
   * ワークブックの印刷設定を再適用
   */
  private restorePrintSettings(workbook: ExcelJS.Workbook, settings: Map<number, any>): void {
    workbook.eachSheet((worksheet) => {
      const saved = settings.get(worksheet.id);
      if (saved) {
        worksheet.pageSetup = saved.pageSetup;
        worksheet.headerFooter = saved.headerFooter;
      }
    });
  }
}
