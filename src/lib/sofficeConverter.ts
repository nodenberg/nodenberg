import { execFile } from 'child_process';
import { promisify } from 'util';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';

const execFileAsync = promisify(execFile);

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

export class SofficeConverter {
  private sofficeCommand: string;
  private tempDir: string;
  private timeout: number;
  private executionMode: 'simple' | 'cmd';

  constructor(options: SofficeConversionOptions = {}) {
    this.sofficeCommand = options.sofficeCommand || this.detectSofficeCommand();
    this.tempDir = options.tempDir || os.tmpdir();
    this.timeout = options.timeout || 30000;
    this.executionMode = options.executionMode || 'simple';
  }

  private validateConfiguredPaths(): void {
    if (!this.sofficeCommand || this.sofficeCommand.includes('\0')) {
      throw new Error('Invalid soffice command');
    }
    if (!this.tempDir || this.tempDir.includes('\0')) {
      throw new Error('Invalid temporary directory');
    }
  }

  private isSofficeDebugEnabled(): boolean {
    return process.env.DEBUG_SOFFICE_COMMAND === 'true';
  }

  /**
   * headless実行時に最低限必要な環境変数を補う
   */
  private buildExecEnv(): NodeJS.ProcessEnv {
    return {
      ...process.env,
      HOME: process.env.HOME || this.tempDir,
      XDG_RUNTIME_DIR: process.env.XDG_RUNTIME_DIR || this.tempDir,
    };
  }

  private async runProcess(args: string[], timeout: number): Promise<{ stdout: string; stderr: string }> {
    this.validateConfiguredPaths();

    if (os.platform() === 'win32' && this.executionMode === 'cmd') {
      return execFileAsync('cmd', ['/c', this.sofficeCommand, ...args], {
        timeout,
        maxBuffer: 10 * 1024 * 1024,
        env: this.buildExecEnv(),
      });
    }

    return execFileAsync(this.sofficeCommand, args, {
      timeout,
      maxBuffer: 10 * 1024 * 1024,
      env: this.buildExecEnv(),
    });
  }

  private async resolveSofficeCommandPath(): Promise<string> {
    if (path.isAbsolute(this.sofficeCommand)) {
      return this.sofficeCommand;
    }

    try {
      if (os.platform() === 'win32') {
        const { stdout } = await execFileAsync('where', [this.sofficeCommand], {
          timeout: 5000,
          maxBuffer: 1024 * 1024,
          env: this.buildExecEnv(),
        });
        return stdout.split(/\r?\n/).map((line) => line.trim()).find(Boolean) || this.sofficeCommand;
      }

      const { stdout } = await execFileAsync('which', [this.sofficeCommand], {
        timeout: 5000,
        maxBuffer: 1024 * 1024,
        env: this.buildExecEnv(),
      });
      return stdout.trim() || this.sofficeCommand;
    } catch {
      return this.sofficeCommand;
    }
  }

  private async logSofficeDebugContext(args: string[]): Promise<void> {
    if (!this.isSofficeDebugEnabled()) {
      return;
    }

    const resolvedCommand = await this.resolveSofficeCommandPath();
    let version = '';
    let versionError = '';

    try {
      const result = await this.runProcess(['--version'], 5000);
      version = result.stdout.trim();
    } catch (error) {
      versionError = error instanceof Error ? error.message : String(error);
    }

    console.log('[soffice-debug]', JSON.stringify({
      configuredCommand: this.sofficeCommand,
      resolvedCommand,
      executionMode: this.executionMode,
      version,
      versionError,
      args,
      env: {
        HOME: this.buildExecEnv().HOME,
        XDG_RUNTIME_DIR: this.buildExecEnv().XDG_RUNTIME_DIR,
      },
    }));
  }

  /**
   * sofficeコマンドのパスを検出
   */
  private detectSofficeCommand(): string {
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
    } else if (platform === 'darwin') {
      // macOSの場合
      return '/Applications/LibreOffice.app/Contents/MacOS/soffice';
    } else {
      // Linuxの場合
      return 'soffice';
    }
  }

  /**
   * LibreOfficeがインストールされているか確認
   */
  async checkSofficeInstalled(): Promise<boolean> {
    try {
      const { stdout } = await this.runProcess(['--version'], 5000);
      return stdout.includes('LibreOffice') || stdout.includes('OpenOffice');
    } catch (error) {
      console.error('Error checking soffice installation:', error);
      return false;
    }
  }

  /**
   * ExcelファイルをPDFに変換
   * @param excelBuffer - Excelファイルのバッファ
   * @returns PDFファイルのバッファ
   */
  async convertExcelToPDF(excelBuffer: Buffer): Promise<Buffer> {
    // LibreOfficeがインストールされているか確認
    const isInstalled = await this.checkSofficeInstalled();
    if (!isInstalled) {
      throw new Error(
        `LibreOffice is not installed or not found at: ${this.sofficeCommand}\n` +
        'Please install LibreOffice from https://www.libreoffice.org/download/download/'
      );
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
      const args = ['--headless', '--convert-to', 'pdf', '--outdir', workDir, inputFile];
      const outputFile = path.join(workDir, 'input.pdf');

      await this.logSofficeDebugContext(args);
      console.log('Executing soffice command:', this.sofficeCommand, args);

      let stdout = '';
      let stderr = '';

      try {
        const result = await this.runProcess(args, this.timeout);
        stdout = result.stdout;
        stderr = result.stderr;
      } catch (error) {
        const execError = error as Error & { stdout?: string; stderr?: string };
        stdout = execError.stdout || '';
        stderr = execError.stderr || '';

        // soffice は warning を出しつつ終了コード 1 を返すことがある。
        // PDF が生成できていれば成功扱いにする。
        if (!fs.existsSync(outputFile)) {
          throw error;
        }

        console.warn('soffice exited with a non-zero status, but PDF was generated successfully');
      }

      if (stderr) {
        console.warn('soffice stderr:', stderr);
      }

      if (stdout) {
        console.log('soffice stdout:', stdout);
      }

      if (!fs.existsSync(outputFile)) {
        throw new Error('PDF file was not generated by soffice');
      }

      const pdfBuffer = fs.readFileSync(outputFile);

      return pdfBuffer;
    } finally {
      // 一時ファイルを削除
      try {
        fs.rmSync(workDir, { recursive: true, force: true });
      } catch (error) {
        console.warn('Failed to remove temporary directory:', error);
      }
    }
  }

  /**
   * Base64エンコードされたExcelファイルをPDFに変換
   * @param base64Excel - Base64エンコードされたExcelデータ
   * @returns Base64エンコードされたPDFデータ
   */
  async convertBase64ExcelToPDF(base64Excel: string): Promise<string> {
    const excelBuffer = Buffer.from(base64Excel, 'base64');
    const pdfBuffer = await this.convertExcelToPDF(excelBuffer);
    return pdfBuffer.toString('base64');
  }

  /**
   * 複数のExcelファイルをPDFに一括変換
   * @param excelBuffers - Excelファイルのバッファの配列
   * @returns PDFファイルのバッファの配列
   */
  async convertMultipleExcelsToPDF(excelBuffers: Buffer[]): Promise<Buffer[]> {
    const results: Buffer[] = [];

    for (const buffer of excelBuffers) {
      const pdfBuffer = await this.convertExcelToPDF(buffer);
      results.push(pdfBuffer);
    }

    return results;
  }

  /**
   * sofficeコマンドのパスを設定
   */
  setSofficeCommand(command: string): void {
    this.sofficeCommand = command;
  }

  /**
   * 現在のsofficeコマンドのパスを取得
   */
  getSofficeCommand(): string {
    return this.sofficeCommand;
  }

  /**
   * LibreOfficeのバージョンを取得
   */
  async getLibreOfficeVersion(): Promise<string> {
    try {
      const { stdout } = await this.runProcess(['--version'], 5000);
      return stdout.trim();
    } catch (error) {
      throw new Error('Failed to get LibreOffice version');
    }
  }

  /**
   * 実行モードを設定
   */
  setExecutionMode(mode: 'simple' | 'cmd'): void {
    this.executionMode = mode;
  }

  /**
   * 現在の実行モードを取得
   */
  getExecutionMode(): 'simple' | 'cmd' {
    return this.executionMode;
  }
}
