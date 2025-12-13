import { PlaceholderData } from './placeholderReplacer';
export declare class ExcelGenerator {
    private placeholderReplacer;
    constructor();
    /**
     * エクセル帳票を生成（test9方式）
     * sharedStrings.xmlを直接編集することで印刷設定を完全保持
     */
    generateExcel(templateBase64: string, data: PlaceholderData): Promise<Buffer>;
    /**
     * エクセル帳票をBase64形式で生成
     */
    generateExcelAsBase64(templateBase64: string, data: PlaceholderData): Promise<string>;
    /**
     * テンプレート内のプレースホルダーを検出
     */
    findPlaceholders(templateBase64: string): Promise<string[]>;
    /**
     * テンプレート内のプレースホルダー詳細情報を取得
     */
    getPlaceholderInfo(templateBase64: string): Promise<{
        placeholder: string;
        key: string;
        count: number;
    }[]>;
    /**
     * テンプレートの基本情報を取得
     * （簡易版 - ZIP構造から取得）
     */
    getTemplateInfo(templateBase64: string): Promise<{
        sheetCount: number;
        sheets: {
            id: number;
            name: any;
            rowCount: number;
            columnCount: number;
        }[];
    }>;
    /**
     * エクセル帳票を生成し、ストリームで返す
     */
    generateExcelStream(templateBase64: string, data: PlaceholderData): Promise<NodeJS.ReadableStream>;
}
//# sourceMappingURL=excelGenerator.d.ts.map