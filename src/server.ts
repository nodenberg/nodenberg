import 'dotenv/config';
import express, { Request, Response, NextFunction } from 'express';
import cors from 'cors';
import { ExcelGenerator } from './lib/excelGenerator';
import { PDFGenerator } from './lib/pdfGenerator';
import { ExcelTemplateManager } from './lib/excelTemplateManager';

const app = express();
const port = process.env.PORT || 3000;
const API_KEY = process.env.API_KEY;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
// Preflight (OPTIONS) は API_KEY 認証の前に通す（ブラウザからの POST が CORS preflight になるため）
app.options(/.*/, cors());

// API Key Authentication Middleware (required)
const apiKeyAuth = (req: Request, res: Response, next: NextFunction) => {
  // Skip authentication for health check endpoint
  if (req.path === '/health') {
    return next();
  }
  // CORS preflight は API_KEY を付けないため、ここで止めるとブラウザからのリクエストが全滅する
  if (req.method === 'OPTIONS') {
    return next();
  }

  // Check if API_KEY is configured
  if (!API_KEY) {
    return res.status(500).json({
      error: 'Server Configuration Error',
      message: 'API_KEY is not configured on the server. Please set API_KEY in .env file.',
    });
  }

  // Check for API key in header
  const providedKey = req.headers['x-api-key'] as string;

  if (!providedKey || providedKey !== API_KEY) {
    return res.status(401).json({
      error: 'Unauthorized',
      message: 'Valid API key required. Provide it in the X-API-Key header.',
    });
  }

  next();
};

// Apply API key authentication to all routes
app.use(apiKeyAuth);

// Error handling helper
const handleError = (res: Response, error: unknown, message: string) => {
  console.error(message, error);
  res.status(500).json({
    error: message,
    details: error instanceof Error ? error.message : String(error),
  });
};

type SheetSelectBy = 'id' | 'name';

function parseDisplayOrder(body: any): number {
  const value = body?.displayOrder;
  const asNumber = typeof value === 'number' ? value : Number(value);
  if (!Number.isInteger(asNumber) || asNumber <= 0) {
    throw new Error('Invalid displayOrder (must be integer >= 1)');
  }
  return asNumber;
}

function parseSheetSelector(body: any):
  | { sheetId: number; sheetName?: undefined }
  | { sheetName: string; sheetId?: undefined }
  | null {
  const sheetSelectBy = body?.sheetSelectBy as SheetSelectBy | undefined;
  const sheetSelectValue = body?.sheetSelectValue as unknown;

  // デフォルト: どちらも揃っていない場合は未指定扱い
  if (!sheetSelectBy || sheetSelectValue === undefined || sheetSelectValue === null) return null;

  if (sheetSelectBy === 'id') {
    const asNumber = typeof sheetSelectValue === 'number' ? sheetSelectValue : Number(sheetSelectValue);
    if (!Number.isInteger(asNumber) || asNumber <= 0) {
      throw new Error('Invalid sheetSelectValue for sheetSelectBy="id" (must be integer >= 1)');
    }
    return { sheetId: asNumber };
  }

  if (sheetSelectBy === 'name') {
    if (typeof sheetSelectValue !== 'string' || sheetSelectValue.trim() === '') {
      throw new Error('Invalid sheetSelectValue for sheetSelectBy="name" (must be non-empty string)');
    }
    return { sheetName: sheetSelectValue };
  }

  throw new Error('Invalid sheetSelectBy (must be "id" or "name")');
}

// ===== Health Check Endpoint =====
app.get('/health', (req: Request, res: Response) => {
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    service: 'Nodenberg API Server',
    version: '0.0.1',
  });
});

// ===== Template Management Endpoints =====

/**
 * POST /template/placeholders
 * Detect placeholders in Excel template
 */
app.post('/template/placeholders', async (req: Request, res: Response) => {
  try {
    const { templateBase64, detailed } = req.body;

    if (!templateBase64) {
      return res.status(400).json({ error: 'Template base64 data is required' });
    }

    const generator = new ExcelGenerator();

    if (detailed) {
      const placeholderInfo = await generator.getPlaceholderInfo(templateBase64);
      return res.json({
        success: true,
        placeholders: placeholderInfo,
      });
    } else {
      const placeholders = await generator.findPlaceholders(templateBase64);
      return res.json({
        success: true,
        placeholders,
      });
    }
  } catch (error) {
    handleError(res, error, 'Failed to find placeholders');
  }
});

/**
 * POST /template/info
 * Get template information (sheet count, dimensions, etc.)
 */
app.post('/template/info', async (req: Request, res: Response) => {
  try {
    const { templateBase64 } = req.body;

    if (!templateBase64) {
      return res.status(400).json({ error: 'Template base64 data is required' });
    }

    const generator = new ExcelGenerator();
    const templateInfo = await generator.getTemplateInfo(templateBase64);

    return res.json({
      success: true,
      templateInfo,
      note: 'Print settings are automatically preserved (test9 method)',
    });
  } catch (error) {
    handleError(res, error, 'Failed to get template info');
  }
});

/**
 * POST /template/sheets
 * Get sheet list with display order and sheetId
 */
app.post('/template/sheets', async (req: Request, res: Response) => {
  try {
    const { templateBase64 } = req.body;

    if (!templateBase64) {
      return res.status(400).json({ error: 'Template base64 data is required' });
    }

    const generator = new ExcelGenerator();
    const sheets = await generator.getTemplateSheets(templateBase64);

    return res.json({
      success: true,
      sheetCount: sheets.length,
      sheets,
    });
  } catch (error) {
    handleError(res, error, 'Failed to get sheet list');
  }
});

async function handleTemplateValidate(req: Request, res: Response) {
  try {
    const { templateId, templateName, base64Data, generateJsonTemplate } = req.body;

    if (!templateId || !templateName || !base64Data) {
      return res.status(400).json({
        error: 'Template ID, name, and base64 data are required',
      });
    }

    const manager = new ExcelTemplateManager();
    const template = await manager.uploadTemplate(
      templateId,
      templateName,
      base64Data,
      generateJsonTemplate || false
    );

    return res.json({
      success: true,
      template: {
        id: template.id,
        name: template.name,
        uploadedAt: template.uploadedAt,
        hasJsonTemplate: !!template.jsonTemplate,
      },
    });
  } catch (error) {
    handleError(res, error, 'Failed to validate template');
  }
}

/**
 * POST /template/validate
 * Validate template payload and optionally generate JSON template
 */
app.post('/template/validate', handleTemplateValidate);

/**
 * POST /template/upload (deprecated alias)
 * Backward-compatible alias of /template/validate
 */
app.post('/template/upload', (req: Request, res: Response) => {
  res.setHeader('Deprecation', 'true');
  res.setHeader('Sunset', '2026-12-31');
  return handleTemplateValidate(req, res);
});

// ===== Document Generation Endpoints =====

/**
 * POST /generate/excel
 * Generate Excel file from template with placeholder replacement
 */
app.post('/generate/excel', async (req: Request, res: Response) => {
  try {
    const { templateBase64, data } = req.body;

    if (!templateBase64) {
      return res.status(400).json({ error: 'Template base64 data is required' });
    }

    if (!data) {
      return res.status(400).json({ error: 'Placeholder data is required' });
    }

    const generator = new ExcelGenerator();
    let excelBuffer: Buffer;
    try {
      const selector = parseSheetSelector(req.body);
      excelBuffer = await generator.generateExcel(templateBase64, data, selector || {});
    } catch (e) {
      return res.status(400).json({
        error: 'Invalid sheet selector',
        details: e instanceof Error ? e.message : String(e),
      });
    }

    // Convert Buffer to base64
    const base64Result = excelBuffer.toString('base64');

    return res.json({
      success: true,
      data: base64Result,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  } catch (error) {
    handleError(res, error, 'Failed to generate Excel');
  }
});

/**
 * POST /generate/excel/by-display-order
 * Generate Excel file by selecting a sheet via display order (1-based)
 */
app.post('/generate/excel/by-display-order', async (req: Request, res: Response) => {
  try {
    const { templateBase64, data } = req.body;

    if (!templateBase64) {
      return res.status(400).json({ error: 'Template base64 data is required' });
    }

    if (!data) {
      return res.status(400).json({ error: 'Placeholder data is required' });
    }

    const generator = new ExcelGenerator();
    let excelBuffer: Buffer;
    try {
      const displayOrder = parseDisplayOrder(req.body);
      const sheets = await generator.getTemplateSheets(templateBase64);
      const selectedSheet = sheets.find((s) => s.displayOrder === displayOrder);
      if (!selectedSheet) {
        return res.status(400).json({
          error: 'Invalid displayOrder',
          details: `displayOrder=${displayOrder} is out of range (1..${sheets.length})`,
        });
      }

      excelBuffer = await generator.generateExcel(templateBase64, data, { sheetName: selectedSheet.name });
    } catch (e) {
      return res.status(400).json({
        error: 'Invalid displayOrder',
        details: e instanceof Error ? e.message : String(e),
      });
    }

    const base64Result = excelBuffer.toString('base64');

    return res.json({
      success: true,
      data: base64Result,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  } catch (error) {
    handleError(res, error, 'Failed to generate Excel');
  }
});

/**
 * POST /generate/pdf
 * Generate PDF file from Excel template (requires LibreOffice)
 */
app.post('/generate/pdf', async (req: Request, res: Response) => {
  try {
    const { templateBase64, data, options } = req.body;

    if (!templateBase64) {
      return res.status(400).json({ error: 'Template base64 data is required' });
    }

    if (!data) {
      return res.status(400).json({ error: 'Placeholder data is required' });
    }

    const generator = new PDFGenerator(options);

    // Check if LibreOffice is installed
    const isInstalled = await generator.checkLibreOfficeInstalled();
    if (!isInstalled) {
      return res.status(503).json({
        error: 'LibreOffice is not installed',
        details: 'PDF generation requires LibreOffice. Please install it from https://www.libreoffice.org/download/download/',
        sofficeCommand: generator.getSofficeCommand(),
        installInstructions: {
          windows: 'Download and install from https://www.libreoffice.org/download/download/',
          macOS: 'Run: brew install --cask libreoffice',
          linux: 'Run: sudo apt-get install libreoffice',
        },
      });
    }

    let pdfBuffer: Buffer;
    try {
      const selector = parseSheetSelector(req.body);
      const pdfOptions = { ...(options || {}), ...(selector || {}) };
      pdfBuffer = await generator.generatePDF(templateBase64, data, pdfOptions);
    } catch (e) {
      return res.status(400).json({
        error: 'Invalid sheet selector',
        details: e instanceof Error ? e.message : String(e),
      });
    }

    // Convert Buffer to base64
    const base64Result = pdfBuffer.toString('base64');

    return res.json({
      success: true,
      data: base64Result,
      mimeType: 'application/pdf',
    });
  } catch (error) {
    handleError(res, error, 'Failed to generate PDF');
  }
});

/**
 * POST /generate/pdf/by-display-order
 * Generate PDF by selecting a sheet via display order (1-based)
 */
app.post('/generate/pdf/by-display-order', async (req: Request, res: Response) => {
  try {
    const { templateBase64, data, options } = req.body;

    if (!templateBase64) {
      return res.status(400).json({ error: 'Template base64 data is required' });
    }

    if (!data) {
      return res.status(400).json({ error: 'Placeholder data is required' });
    }

    const generator = new PDFGenerator(options);

    const isInstalled = await generator.checkLibreOfficeInstalled();
    if (!isInstalled) {
      return res.status(503).json({
        error: 'LibreOffice is not installed',
        details: 'PDF generation requires LibreOffice. Please install it from https://www.libreoffice.org/download/download/',
        sofficeCommand: generator.getSofficeCommand(),
        installInstructions: {
          windows: 'Download and install from https://www.libreoffice.org/download/download/',
          macOS: 'Run: brew install --cask libreoffice',
          linux: 'Run: sudo apt-get install libreoffice',
        },
      });
    }

    let pdfBuffer: Buffer;
    try {
      const displayOrder = parseDisplayOrder(req.body);
      const excelGenerator = new ExcelGenerator();
      const sheets = await excelGenerator.getTemplateSheets(templateBase64);
      const selectedSheet = sheets.find((s) => s.displayOrder === displayOrder);
      if (!selectedSheet) {
        return res.status(400).json({
          error: 'Invalid displayOrder',
          details: `displayOrder=${displayOrder} is out of range (1..${sheets.length})`,
        });
      }

      const pdfOptions = { ...(options || {}), sheetName: selectedSheet.name };
      pdfBuffer = await generator.generatePDF(templateBase64, data, pdfOptions);
    } catch (e) {
      return res.status(400).json({
        error: 'Invalid displayOrder',
        details: e instanceof Error ? e.message : String(e),
      });
    }

    const base64Result = pdfBuffer.toString('base64');

    return res.json({
      success: true,
      data: base64Result,
      mimeType: 'application/pdf',
    });
  } catch (error) {
    handleError(res, error, 'Failed to generate PDF');
  }
});

// ===== 404 Handler =====
app.use((req: Request, res: Response) => {
  res.status(404).json({
    error: 'Not Found',
    message: `Endpoint ${req.method} ${req.path} does not exist`,
    availableEndpoints: [
      'GET /health',
      'POST /template/placeholders',
      'POST /template/info',
      'POST /template/sheets',
      'POST /template/validate',
      'POST /template/upload',
      'POST /generate/excel',
      'POST /generate/excel/by-display-order',
      'POST /generate/pdf',
      'POST /generate/pdf/by-display-order',
    ],
  });
});

// ===== Start Server =====
app.listen(port, () => {
  console.log('╔══════════════════════════════════════════════════════════════╗');
  console.log('║                                                              ║');
  console.log('║   Nodenberg API Server                                       ║');
  console.log('║   Excel & PDF Generation with test9 Method                   ║');
  console.log('║                                                              ║');
  console.log('╚══════════════════════════════════════════════════════════════╝');
  console.log('');
  console.log(`✓ Server running on: http://localhost:${port}`);
  console.log(`✓ Health check:      http://localhost:${port}/health`);
  console.log(`✓ API Authentication: ${API_KEY ? '🔒 Enabled' : '🔓 Disabled'}`);
  console.log('');
  console.log('Available Endpoints:');
  console.log('  • GET  /health');
  console.log('  • POST /template/placeholders');
  console.log('  • POST /template/info');
  console.log('  • POST /template/sheets');
  console.log('  • POST /template/validate');
  console.log('  • POST /template/upload');
  console.log('  • POST /generate/excel');
  console.log('  • POST /generate/excel/by-display-order');
  console.log('  • POST /generate/pdf');
  console.log('  • POST /generate/pdf/by-display-order');
  console.log('');
  console.log('Press Ctrl+C to stop');
  console.log('');
});

// ===== Graceful Shutdown =====
process.on('SIGTERM', () => {
  console.log('SIGTERM received, shutting down gracefully...');
  process.exit(0);
});

process.on('SIGINT', () => {
  console.log('\nSIGINT received, shutting down gracefully...');
  process.exit(0);
});
