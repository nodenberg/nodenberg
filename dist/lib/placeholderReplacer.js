"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.PlaceholderReplacer = void 0;
const jszip_1 = __importDefault(require("jszip"));
/**
 * XML特殊文字をエスケープ（W3C準拠）
 */
function escapeXml(text) {
    return text
        .replace(/&/g, '&amp;') // & を最初に処理（重要）
        .replace(/</g, '&lt;') // < を変換
        .replace(/>/g, '&gt;') // > を変換
        .replace(/"/g, '&quot;') // " を変換
        .replace(/'/g, '&apos;'); // ' を変換
}
/**
 * 正規表現用に文字列をエスケープ
 */
function escapeRegExp(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
/**
 * XML内のプレースホルダーを検出
 */
function detectPlaceholdersInXml(xml) {
    const placeholderMap = new Map();
    const regex = /\{\{([^}]+)\}\}/g;
    let match;
    while ((match = regex.exec(xml)) !== null) {
        const placeholder = match[0];
        const key = match[1];
        if (placeholderMap.has(placeholder)) {
            placeholderMap.get(placeholder).count++;
        }
        else {
            placeholderMap.set(placeholder, {
                placeholder: placeholder,
                key: key,
                count: 1
            });
        }
    }
    return Array.from(placeholderMap.values());
}
class PlaceholderReplacer {
    constructor() {
        this.placeholderPattern = /\{\{([^}]+)\}\}/g;
    }
    /**
     * Excelファイル（Buffer）内のプレースホルダーを置換
     * sharedStrings.xmlを直接編集することで印刷設定を完全保持
     */
    async replacePlaceholders(excelBuffer, data) {
        // 1. ZIPとして読み込み
        const zip = await jszip_1.default.loadAsync(excelBuffer);
        // 2. sharedStrings.xmlを取得（プレースホルダーはここにある）
        const sharedStringsFile = zip.file('xl/sharedStrings.xml');
        if (!sharedStringsFile) {
            throw new Error('sharedStrings.xmlが見つかりません');
        }
        let sharedStringsXml = await sharedStringsFile.async('string');
        // 3. プレースホルダー置換（自動エスケープ）
        let replacedSharedStrings = sharedStringsXml;
        Object.entries(data).forEach(([key, value]) => {
            const placeholder = `{{${key}}}`;
            // 値を文字列に変換
            let stringValue;
            if (value === null || value === undefined) {
                stringValue = '';
            }
            else if (value instanceof Date) {
                stringValue = this.formatDate(value);
            }
            else {
                stringValue = String(value);
            }
            // XMLエスケープ
            const escapedValue = escapeXml(stringValue);
            // プレースホルダーを検索して置換
            const regex = new RegExp(escapeRegExp(placeholder), 'g');
            replacedSharedStrings = replacedSharedStrings.replace(regex, escapedValue);
        });
        // 4. ZIPに書き戻し
        zip.file('xl/sharedStrings.xml', replacedSharedStrings);
        // 5. Bufferとして返す
        const outputBuffer = await zip.generateAsync({
            type: 'nodebuffer',
            compression: 'DEFLATE',
            compressionOptions: { level: 9 }
        });
        return outputBuffer;
    }
    /**
     * Excelファイル（Buffer）内のプレースホルダーを検出
     */
    async findPlaceholders(excelBuffer) {
        // 1. ZIPとして読み込み
        const zip = await jszip_1.default.loadAsync(excelBuffer);
        // 2. sharedStrings.xmlを取得
        const sharedStringsFile = zip.file('xl/sharedStrings.xml');
        if (!sharedStringsFile) {
            return [];
        }
        const sharedStringsXml = await sharedStringsFile.async('string');
        // 3. プレースホルダー検出
        const placeholders = detectPlaceholdersInXml(sharedStringsXml);
        // 4. キーのみを返す（重複を除去してソート）
        const keys = placeholders.map(p => p.key);
        return Array.from(new Set(keys)).sort();
    }
    /**
     * プレースホルダーの詳細情報を取得
     */
    async getPlaceholderInfo(excelBuffer) {
        // 1. ZIPとして読み込み
        const zip = await jszip_1.default.loadAsync(excelBuffer);
        // 2. sharedStrings.xmlを取得
        const sharedStringsFile = zip.file('xl/sharedStrings.xml');
        if (!sharedStringsFile) {
            return [];
        }
        const sharedStringsXml = await sharedStringsFile.async('string');
        // 3. プレースホルダー検出
        return detectPlaceholdersInXml(sharedStringsXml);
    }
    /**
     * 日付をフォーマット（yyyy/MM/dd形式）
     */
    formatDate(date) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}/${month}/${day}`;
    }
}
exports.PlaceholderReplacer = PlaceholderReplacer;
//# sourceMappingURL=placeholderReplacer.js.map