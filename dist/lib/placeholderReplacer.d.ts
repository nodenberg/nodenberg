export interface PlaceholderData {
    [key: string]: string | number | Date | null;
}
export declare class PlaceholderReplacer {
    private placeholderPattern;
    /**
     * Excelファイル（Buffer）内のプレースホルダーを置換
     * sharedStrings.xmlを直接編集することで印刷設定を完全保持
     */
    replacePlaceholders(excelBuffer: Buffer, data: PlaceholderData): Promise<Buffer>;
    /**
     * Excelファイル（Buffer）内のプレースホルダーを検出
     */
    findPlaceholders(excelBuffer: Buffer): Promise<string[]>;
    /**
     * プレースホルダーの詳細情報を取得
     */
    getPlaceholderInfo(excelBuffer: Buffer): Promise<Array<{
        placeholder: string;
        key: string;
        count: number;
    }>>;
    /**
     * 日付をフォーマット（yyyy/MM/dd形式）
     */
    private formatDate;
}
//# sourceMappingURL=placeholderReplacer.d.ts.map