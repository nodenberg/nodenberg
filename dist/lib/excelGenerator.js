"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelGenerator = void 0;
const placeholderReplacer_1 = require("./placeholderReplacer");
class ExcelGenerator {
    constructor() {
        this.placeholderReplacer = new placeholderReplacer_1.PlaceholderReplacer();
    }
    /**
     * エクセル帳票を生成（test9方式）
     * sharedStrings.xmlを直接編集することで印刷設定を完全保持
     */
    async generateExcel(templateBase64, data) {
        // Base64をBufferに変換
        const templateBuffer = Buffer.from(templateBase64, 'base64');
        // プレースホルダーを置換（印刷設定は完全保持される）
        const resultBuffer = await this.placeholderReplacer.replacePlaceholders(templateBuffer, data);
        return resultBuffer;
    }
    /**
     * エクセル帳票をBase64形式で生成
     */
    async generateExcelAsBase64(templateBase64, data) {
        const buffer = await this.generateExcel(templateBase64, data);
        return buffer.toString('base64');
    }
    /**
     * テンプレート内のプレースホルダーを検出
     */
    async findPlaceholders(templateBase64) {
        const templateBuffer = Buffer.from(templateBase64, 'base64');
        return this.placeholderReplacer.findPlaceholders(templateBuffer);
    }
    /**
     * テンプレート内のプレースホルダー詳細情報を取得
     */
    async getPlaceholderInfo(templateBase64) {
        const templateBuffer = Buffer.from(templateBase64, 'base64');
        return this.placeholderReplacer.getPlaceholderInfo(templateBuffer);
    }
    /**
     * テンプレートの基本情報を取得
     * （簡易版 - ZIP構造から取得）
     */
    async getTemplateInfo(templateBase64) {
        const JSZip = require('jszip');
        const templateBuffer = Buffer.from(templateBase64, 'base64');
        const zip = await JSZip.loadAsync(templateBuffer);
        // ワークシート数を取得
        const worksheetFiles = Object.keys(zip.files).filter((filename) => filename.startsWith('xl/worksheets/sheet') && filename.endsWith('.xml'));
        const sheets = await Promise.all(worksheetFiles.map(async (filename, index) => {
            const content = await zip.file(filename)?.async('string');
            // シート名を取得（workbook.xmlから）
            const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
            const sheetNameMatch = workbookXml?.match(new RegExp(`<sheet[^>]*sheetId="${index + 1}"[^>]*name="([^"]*)"`));
            const sheetName = sheetNameMatch?.[1] || `Sheet${index + 1}`;
            // 行数・列数の概算
            const dimensionMatch = content?.match(/<dimension[^>]*ref="([^"]*)"/);
            const dimension = dimensionMatch?.[1] || 'A1';
            return {
                id: index + 1,
                name: sheetName,
                rowCount: 0, // 簡易版では省略
                columnCount: 0, // 簡易版では省略
            };
        }));
        return {
            sheetCount: worksheetFiles.length,
            sheets,
        };
    }
    /**
     * エクセル帳票を生成し、ストリームで返す
     */
    async generateExcelStream(templateBase64, data) {
        const buffer = await this.generateExcel(templateBase64, data);
        const { PassThrough } = require('stream');
        const stream = new PassThrough();
        stream.end(buffer);
        return stream;
    }
}
exports.ExcelGenerator = ExcelGenerator;
//# sourceMappingURL=excelGenerator.js.map