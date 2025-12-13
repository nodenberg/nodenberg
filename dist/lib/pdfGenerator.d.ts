import { PlaceholderData } from './placeholderReplacer';
import { SofficeConversionOptions } from './sofficeConverter';
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
export declare class PDFGenerator {
    private placeholderReplacer;
    private sofficeConverter;
    constructor(sofficeOptions?: SofficeConversionOptions);
    /**
     * エクセル帳票からPDFを生成（test9方式）
     * LibreOfficeのsoffice headlessを使用してExcelの見た目そのままPDF化
     */
    generatePDF(templateBase64: string, data: PlaceholderData, options?: PDFGenerationOptions): Promise<Buffer>;
    /**
     * PDFをBase64形式で生成
     */
    generatePDFAsBase64(templateBase64: string, data: PlaceholderData, options?: PDFGenerationOptions): Promise<string>;
    /**
     * LibreOfficeがインストールされているか確認
     */
    checkLibreOfficeInstalled(): Promise<boolean>;
    /**
     * LibreOfficeのバージョンを取得
     */
    getLibreOfficeVersion(): Promise<string>;
    /**
     * sofficeコマンドのパスを設定
     */
    setSofficeCommand(command: string): void;
    /**
     * 現在のsofficeコマンドのパスを取得
     */
    getSofficeCommand(): string;
}
//# sourceMappingURL=pdfGenerator.d.ts.map