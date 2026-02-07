#!/usr/bin/env node

/**
 * LibreOffice初期化スクリプト
 *
 * アプリケーション起動前にLibreOfficeを初期化して、
 * 初回PDF生成時の印刷設定の問題を回避します。
 * Dockerfileで呼び出して実行する。
 */

const { exec } = require('child_process');
const { promisify } = require('util');
const fs = require('fs');
const path = require('path');
const os = require('os');

const execAsync = promisify(exec);

async function initializeLibreOffice() {
  console.log('[LibreOffice Init] Starting LibreOffice initialization...');

  try {
    // LibreOfficeがインストールされているか確認
    const { stdout } = await execAsync('soffice --version', { timeout: 5000 });

    if (!stdout.includes('LibreOffice')) {
      console.warn('[LibreOffice Init] LibreOffice not found, skipping initialization');
      return;
    }

    console.log('[LibreOffice Init] LibreOffice detected:', stdout.trim());

    // ダミーのExcelファイルを作成（最小限のxlsxファイル）
    const workDir = path.join(os.tmpdir(), `libreoffice-init-${Date.now()}`);
    fs.mkdirSync(workDir, { recursive: true });

    try {
      const dummyXlsx = path.join(workDir, 'dummy.xlsx');

      // 最小限の有効なxlsxファイルを作成
      // （実際のxlsxファイルの代わりに、シンプルなExcelJSで生成）
      const ExcelJS = require('exceljs');
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');
      worksheet.getCell('A1').value = 'Init';
      await workbook.xlsx.writeFile(dummyXlsx);

      // LibreOfficeでPDFに変換（これで初期化される）
      const command = `soffice --headless --convert-to pdf --outdir "${workDir}" "${dummyXlsx}"`;

      console.log('[LibreOffice Init] Running initialization command...');

      const { stdout: convertOutput, stderr } = await execAsync(command, {
        timeout: 30000,
        maxBuffer: 10 * 1024 * 1024,
      });

      if (stderr && !stderr.includes('Warning: failed to launch javaldx')) {
        console.warn('[LibreOffice Init] stderr:', stderr);
      }

      console.log('[LibreOffice Init] Initialization complete');
      console.log('[LibreOffice Init] User profile created at ~/.config/libreoffice/');

    } finally {
      // 一時ファイルを削除
      try {
        fs.rmSync(workDir, { recursive: true, force: true });
      } catch (error) {
        console.warn('[LibreOffice Init] Failed to remove temp directory:', error.message);
      }
    }

  } catch (error) {
    console.error('[LibreOffice Init] Initialization failed:', error.message);
    console.warn('[LibreOffice Init] PDF generation may be slower on first request');
  }
}

// スクリプトとして直接実行された場合のみ実行
if (require.main === module) {
  initializeLibreOffice()
    .then(() => {
      console.log('[LibreOffice Init] Done');
      process.exit(0);
    })
    .catch((error) => {
      console.error('[LibreOffice Init] Fatal error:', error);
      process.exit(1);
    });
}

module.exports = { initializeLibreOffice };
