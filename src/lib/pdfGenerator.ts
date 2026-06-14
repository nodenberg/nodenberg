import { PlaceholderData, PrintLayoutOptions } from './placeholderReplacer';
import { SofficeConverter, SofficeConversionOptions } from './sofficeConverter';
import { ExcelGenerator } from './excelGenerator';

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

  /**
   * 特定のシートのみをPDFに変換（sheetIdで指定）
   * sheetName と同時に指定された場合は sheetId を優先します。
   */
  sheetId?: number;

  /**
   * Excel/PDF共通の印刷レイアウト設定
   */
  printLayout?: PrintLayoutOptions;
}

export class PDFGenerator {
  private excelGenerator: ExcelGenerator;
  private sofficeConverter: SofficeConverter;

  constructor(sofficeOptions?: SofficeConversionOptions) {
    this.excelGenerator = new ExcelGenerator();
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

    const excelBuffer = await this.excelGenerator.generateExcel(templateBase64, data, {
      sheetName: options.sheetName,
      sheetId: options.sheetId,
      printLayout: options.printLayout,
    });

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
