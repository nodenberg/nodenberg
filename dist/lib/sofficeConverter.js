"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.SofficeConverter = void 0;
const child_process_1 = require("child_process");
const util_1 = require("util");
const fs = __importStar(require("fs"));
const path = __importStar(require("path"));
const os = __importStar(require("os"));
const execAsync = (0, util_1.promisify)(child_process_1.exec);
class SofficeConverter {
    constructor(options = {}) {
        this.sofficeCommand = options.sofficeCommand || this.detectSofficeCommand();
        this.tempDir = options.tempDir || os.tmpdir();
        this.timeout = options.timeout || 30000;
        this.executionMode = options.executionMode || 'simple';
    }
    /**
     * コマンドラインを構築
     */
    buildCommand(args) {
        const platform = os.platform();
        if (this.executionMode === 'simple') {
            // シンプルモード: sofficeコマンドを直接実行
            if (platform === 'win32') {
                // Windowsの場合は引用符で囲む
                return `"${this.sofficeCommand}" ${args}`;
            }
            else {
                return `${this.sofficeCommand} ${args}`;
            }
        }
        else {
            // cmdモード: Windows環境でcmd.exe経由で実行
            if (platform === 'win32') {
                return `cmd /c ""${this.sofficeCommand}" ${args}"`;
            }
            else {
                return `"${this.sofficeCommand}" ${args}`;
            }
        }
    }
    /**
     * sofficeコマンドのパスを検出
     */
    detectSofficeCommand() {
        const platform = os.platform();
        if (platform === 'win32') {
            // Windowsの場合はsoffice.comを使用（コンソールアプリケーション用）
            const possiblePaths = [
                'C:\\Program Files\\LibreOffice\\program\\soffice.com',
                'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.com',
                'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
                'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
                'soffice',
            ];
            return possiblePaths[0]; // デフォルトパスを返す
        }
        else if (platform === 'darwin') {
            // macOSの場合
            return '/Applications/LibreOffice.app/Contents/MacOS/soffice';
        }
        else {
            // Linuxの場合
            return 'soffice';
        }
    }
    /**
     * LibreOfficeがインストールされているか確認
     */
    async checkSofficeInstalled() {
        try {
            const command = this.buildCommand('--version');
            const { stdout } = await execAsync(command, {
                timeout: 5000,
            });
            return stdout.includes('LibreOffice') || stdout.includes('OpenOffice');
        }
        catch (error) {
            console.error('Error checking soffice installation:', error);
            return false;
        }
    }
    /**
     * ExcelファイルをPDFに変換
     * @param excelBuffer - Excelファイルのバッファ
     * @returns PDFファイルのバッファ
     */
    async convertExcelToPDF(excelBuffer) {
        // LibreOfficeがインストールされているか確認
        const isInstalled = await this.checkSofficeInstalled();
        if (!isInstalled) {
            throw new Error(`LibreOffice is not installed or not found at: ${this.sofficeCommand}\n` +
                'Please install LibreOffice from https://www.libreoffice.org/download/download/');
        }
        // 一時ディレクトリを作成
        const workDir = path.join(this.tempDir, `soffice-${Date.now()}-${Math.random().toString(36).substring(7)}`);
        fs.mkdirSync(workDir, { recursive: true });
        try {
            // 一時的なExcelファイルを作成
            const inputFile = path.join(workDir, 'input.xlsx');
            fs.writeFileSync(inputFile, excelBuffer);
            // sofficeコマンドでPDFに変換
            // --headless: ヘッドレスモード
            // --convert-to pdf: PDF形式に変換
            // --outdir: 出力ディレクトリ
            const command = this.buildCommand(`--headless --convert-to pdf --outdir "${workDir}" "${inputFile}"`);
            console.log('Executing soffice command:', command);
            const { stdout, stderr } = await execAsync(command, {
                timeout: this.timeout,
                maxBuffer: 10 * 1024 * 1024, // 10MB
            });
            if (stderr) {
                console.warn('soffice stderr:', stderr);
            }
            console.log('soffice stdout:', stdout);
            // 生成されたPDFファイルを読み込み
            const outputFile = path.join(workDir, 'input.pdf');
            if (!fs.existsSync(outputFile)) {
                throw new Error('PDF file was not generated by soffice');
            }
            const pdfBuffer = fs.readFileSync(outputFile);
            return pdfBuffer;
        }
        finally {
            // 一時ファイルを削除
            try {
                fs.rmSync(workDir, { recursive: true, force: true });
            }
            catch (error) {
                console.warn('Failed to remove temporary directory:', error);
            }
        }
    }
    /**
     * Base64エンコードされたExcelファイルをPDFに変換
     * @param base64Excel - Base64エンコードされたExcelデータ
     * @returns Base64エンコードされたPDFデータ
     */
    async convertBase64ExcelToPDF(base64Excel) {
        const excelBuffer = Buffer.from(base64Excel, 'base64');
        const pdfBuffer = await this.convertExcelToPDF(excelBuffer);
        return pdfBuffer.toString('base64');
    }
    /**
     * 複数のExcelファイルをPDFに一括変換
     * @param excelBuffers - Excelファイルのバッファの配列
     * @returns PDFファイルのバッファの配列
     */
    async convertMultipleExcelsToPDF(excelBuffers) {
        const results = [];
        for (const buffer of excelBuffers) {
            const pdfBuffer = await this.convertExcelToPDF(buffer);
            results.push(pdfBuffer);
        }
        return results;
    }
    /**
     * sofficeコマンドのパスを設定
     */
    setSofficeCommand(command) {
        this.sofficeCommand = command;
    }
    /**
     * 現在のsofficeコマンドのパスを取得
     */
    getSofficeCommand() {
        return this.sofficeCommand;
    }
    /**
     * LibreOfficeのバージョンを取得
     */
    async getLibreOfficeVersion() {
        try {
            const command = this.buildCommand('--version');
            const { stdout } = await execAsync(command, {
                timeout: 5000,
            });
            return stdout.trim();
        }
        catch (error) {
            throw new Error('Failed to get LibreOffice version');
        }
    }
    /**
     * 実行モードを設定
     */
    setExecutionMode(mode) {
        this.executionMode = mode;
    }
    /**
     * 現在の実行モードを取得
     */
    getExecutionMode() {
        return this.executionMode;
    }
}
exports.SofficeConverter = SofficeConverter;
//# sourceMappingURL=sofficeConverter.js.map