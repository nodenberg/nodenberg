export interface SofficeConversionOptions {
    /**
     * LibreOfficeのsofficeコマンドのパス
     * 指定しない場合は環境変数PATHから検索
     */
    sofficeCommand?: string;
    /**
     * 一時ファイルを保存するディレクトリ
     * 指定しない場合はシステムのテンポラリディレクトリ
     */
    tempDir?: string;
    /**
     * タイムアウト時間（ミリ秒）
     * デフォルト: 30000 (30秒)
     */
    timeout?: number;
    /**
     * コマンド実行モード
     * 'simple': sofficeコマンドを直接実行（デフォルト）
     * 'cmd': Windowsの場合cmd.exe経由で実行
     */
    executionMode?: 'simple' | 'cmd';
}
export declare class SofficeConverter {
    private sofficeCommand;
    private tempDir;
    private timeout;
    private executionMode;
    constructor(options?: SofficeConversionOptions);
    /**
     * コマンドラインを構築
     */
    private buildCommand;
    /**
     * sofficeコマンドのパスを検出
     */
    private detectSofficeCommand;
    /**
     * LibreOfficeがインストールされているか確認
     */
    checkSofficeInstalled(): Promise<boolean>;
    /**
     * ExcelファイルをPDFに変換
     * @param excelBuffer - Excelファイルのバッファ
     * @returns PDFファイルのバッファ
     */
    convertExcelToPDF(excelBuffer: Buffer): Promise<Buffer>;
    /**
     * Base64エンコードされたExcelファイルをPDFに変換
     * @param base64Excel - Base64エンコードされたExcelデータ
     * @returns Base64エンコードされたPDFデータ
     */
    convertBase64ExcelToPDF(base64Excel: string): Promise<string>;
    /**
     * 複数のExcelファイルをPDFに一括変換
     * @param excelBuffers - Excelファイルのバッファの配列
     * @returns PDFファイルのバッファの配列
     */
    convertMultipleExcelsToPDF(excelBuffers: Buffer[]): Promise<Buffer[]>;
    /**
     * sofficeコマンドのパスを設定
     */
    setSofficeCommand(command: string): void;
    /**
     * 現在のsofficeコマンドのパスを取得
     */
    getSofficeCommand(): string;
    /**
     * LibreOfficeのバージョンを取得
     */
    getLibreOfficeVersion(): Promise<string>;
    /**
     * 実行モードを設定
     */
    setExecutionMode(mode: 'simple' | 'cmd'): void;
    /**
     * 現在の実行モードを取得
     */
    getExecutionMode(): 'simple' | 'cmd';
}
//# sourceMappingURL=sofficeConverter.d.ts.map