import express from 'express';
import cors from 'cors';
import { ExcelGenerator } from './lib/excelGenerator';
import { PDFGenerator } from './lib/pdfGenerator';

const app = express();
const port = process.env.PORT || 4000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// Helper to handle errors
const handleError = (res: express.Response, error: unknown, message: string) => {
  console.error(message, error);
  res.status(500).json({
    error: message,
    details: error instanceof Error ? error.message : String(error),
  });
};

/**
 * Endpoint: POST /api/generate/excel
 * Description: Generates an Excel file from a template and data.
 */
app.post('/api/generate/excel', async (req, res) => {
  try {
    const { templateBase64, data } = req.body;

    if (!templateBase64) {
      return res.status(400).json({ error: 'Template base64 data is required' });
    }
    if (!data) {
      return res.status(400).json({ error: 'Placeholder data is required' });
    }

    const generator = new ExcelGenerator();
    const excelBuffer = await generator.generateExcel(templateBase64, data);
    const base64Result = excelBuffer.toString('base64');

    res.json({
      success: true,
      data: base64Result,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  } catch (error) {
    handleError(res, error, 'Failed to generate Excel');
  }
});

/**
 * Endpoint: POST /api/generate/pdf
 * Description: Generates a PDF file from an Excel template and data.
 */
app.post('/api/generate/pdf', async (req, res) => {
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
        details: 'PDF generation requires LibreOffice.',
        sofficeCommand: generator.getSofficeCommand(),
      });
    }

    const pdfBuffer = await generator.generatePDF(templateBase64, data, options || {});
    const base64Result = pdfBuffer.toString('base64');

    res.json({
      success: true,
      data: base64Result,
      mimeType: 'application/pdf',
    });
  } catch (error) {
    handleError(res, error, 'Failed to generate PDF');
  }
});

/**
 * Endpoint: POST /api/template/info
 * Description: Analyzes a template and returns info (sheets, placeholders).
 */
app.post('/api/template/info', async (req, res) => {
  try {
    const { templateBase64 } = req.body;
    if (!templateBase64) {
      return res.status(400).json({ error: 'Template base64 data is required' });
    }

    const generator = new ExcelGenerator();
    const templateInfo = await generator.getTemplateInfo(templateBase64);
    const placeholders = await generator.findPlaceholders(templateBase64);

    res.json({
      success: true,
      info: {
        ...templateInfo,
        placeholders
      },
      note: 'Print settings are automatically preserved (test9 method)'
    });

  } catch (error) {
    handleError(res, error, 'Failed to analyze template');
  }
});

// Start Server
app.listen(port, () => {
  console.log(`Express API Server listening at http://localhost:${port}`);
});
