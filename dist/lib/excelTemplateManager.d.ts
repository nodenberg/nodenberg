import ExcelJS from 'exceljs';
export interface ExcelTemplate {
    id: string;
    name: string;
    base64Data: string;
    jsonTemplate?: any;
    uploadedAt: Date;
}
export declare class ExcelTemplateManager {
    /**
     * Base64文字列からExcelJSワークブックを読み込む
     */
    loadWorkbookFromBase64(base64Data: string): Promise<ExcelJS.Workbook>;
    /**
     * エクセルテンプレートをアップロードし、必要に応じてJSONテンプレートを生成
     */
    uploadTemplate(templateId: string, templateName: string, base64Data: string, generateJsonTemplate?: boolean): Promise<ExcelTemplate>;
    /**
     * ワークブックからJSONテンプレートを生成
     * ライブラリによってはJSON形式のテンプレートが必要な場合に使用
     */
    private generateJsonTemplate;
    /**
     * ワークブックのシート情報を取得
     */
    getWorkbookInfo(workbook: ExcelJS.Workbook): {
        sheetCount: number;
        sheets: {
            id: number;
            name: string;
            state: ExcelJS.WorksheetState;
            rowCount: number;
            columnCount: number;
            actualRowCount: number;
            actualColumnCount: number;
        }[];
        creator: string;
        created: Date;
        modified: Date;
    };
    /**
     * ワークブックの印刷設定を取得
     */
    getPrintSettings(worksheet: ExcelJS.Worksheet): {
        paperSize: ExcelJS.PaperSize | undefined;
        orientation: "portrait" | "landscape" | undefined;
        horizontalCentered: boolean | undefined;
        verticalCentered: boolean | undefined;
        margins: ExcelJS.Margins | undefined;
        printArea: string | undefined;
        printTitlesRow: string | undefined;
        printTitlesColumn: string | undefined;
        fitToPage: boolean | undefined;
        fitToHeight: number | undefined;
        fitToWidth: number | undefined;
        scale: number | undefined;
    };
}
//# sourceMappingURL=excelTemplateManager.d.ts.map