"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
require("dotenv/config");
const express_1 = __importDefault(require("express"));
const cors_1 = __importDefault(require("cors"));
const excelGenerator_1 = require("./lib/excelGenerator");
const pdfGenerator_1 = require("./lib/pdfGenerator");
const excelTemplateManager_1 = require("./lib/excelTemplateManager");
const app = (0, express_1.default)();
const port = process.env.PORT || 3000;
const API_KEY = process.env.API_KEY;
// Middleware
app.use((0, cors_1.default)());
app.use(express_1.default.json({ limit: '50mb' }));
// API Key Authentication Middleware (optional)
const apiKeyAuth = (req, res, next) => {
    // Skip authentication if API_KEY is not set
    if (!API_KEY) {
        return next();
    }
    // Skip authentication for health check endpoint
    if (req.path === '/health') {
        return next();
    }
    // Check for API key in header
    const providedKey = req.headers['x-api-key'];
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
const handleError = (res, error, message) => {
    console.error(message, error);
    res.status(500).json({
        error: message,
        details: error instanceof Error ? error.message : String(error),
    });
};
// ===== Health Check Endpoint =====
app.get('/health', (req, res) => {
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
app.post('/template/placeholders', async (req, res) => {
    try {
        const { templateBase64, detailed } = req.body;
        if (!templateBase64) {
            return res.status(400).json({ error: 'Template base64 data is required' });
        }
        const generator = new excelGenerator_1.ExcelGenerator();
        if (detailed) {
            const placeholderInfo = await generator.getPlaceholderInfo(templateBase64);
            return res.json({
                success: true,
                placeholders: placeholderInfo,
            });
        }
        else {
            const placeholders = await generator.findPlaceholders(templateBase64);
            return res.json({
                success: true,
                placeholders,
            });
        }
    }
    catch (error) {
        handleError(res, error, 'Failed to find placeholders');
    }
});
/**
 * POST /template/info
 * Get template information (sheet count, dimensions, etc.)
 */
app.post('/template/info', async (req, res) => {
    try {
        const { templateBase64 } = req.body;
        if (!templateBase64) {
            return res.status(400).json({ error: 'Template base64 data is required' });
        }
        const generator = new excelGenerator_1.ExcelGenerator();
        const templateInfo = await generator.getTemplateInfo(templateBase64);
        return res.json({
            success: true,
            templateInfo,
            note: 'Print settings are automatically preserved (test9 method)',
        });
    }
    catch (error) {
        handleError(res, error, 'Failed to get template info');
    }
});
/**
 * POST /template/upload
 * Upload template (store with optional JSON template generation)
 */
app.post('/template/upload', async (req, res) => {
    try {
        const { templateId, templateName, base64Data, generateJsonTemplate } = req.body;
        if (!templateId || !templateName || !base64Data) {
            return res.status(400).json({
                error: 'Template ID, name, and base64 data are required',
            });
        }
        const manager = new excelTemplateManager_1.ExcelTemplateManager();
        const template = await manager.uploadTemplate(templateId, templateName, base64Data, generateJsonTemplate || false);
        return res.json({
            success: true,
            template: {
                id: template.id,
                name: template.name,
                uploadedAt: template.uploadedAt,
                hasJsonTemplate: !!template.jsonTemplate,
            },
        });
    }
    catch (error) {
        handleError(res, error, 'Failed to upload template');
    }
});
// ===== Document Generation Endpoints =====
/**
 * POST /generate/excel
 * Generate Excel file from template with placeholder replacement
 */
app.post('/generate/excel', async (req, res) => {
    try {
        const { templateBase64, data } = req.body;
        if (!templateBase64) {
            return res.status(400).json({ error: 'Template base64 data is required' });
        }
        if (!data) {
            return res.status(400).json({ error: 'Placeholder data is required' });
        }
        const generator = new excelGenerator_1.ExcelGenerator();
        const excelBuffer = await generator.generateExcel(templateBase64, data);
        // Convert Buffer to base64
        const base64Result = excelBuffer.toString('base64');
        return res.json({
            success: true,
            data: base64Result,
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
    }
    catch (error) {
        handleError(res, error, 'Failed to generate Excel');
    }
});
/**
 * POST /generate/pdf
 * Generate PDF file from Excel template (requires LibreOffice)
 */
app.post('/generate/pdf', async (req, res) => {
    try {
        const { templateBase64, data, options } = req.body;
        if (!templateBase64) {
            return res.status(400).json({ error: 'Template base64 data is required' });
        }
        if (!data) {
            return res.status(400).json({ error: 'Placeholder data is required' });
        }
        const generator = new pdfGenerator_1.PDFGenerator(options);
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
        const pdfBuffer = await generator.generatePDF(templateBase64, data, options || {});
        // Convert Buffer to base64
        const base64Result = pdfBuffer.toString('base64');
        return res.json({
            success: true,
            data: base64Result,
            mimeType: 'application/pdf',
        });
    }
    catch (error) {
        handleError(res, error, 'Failed to generate PDF');
    }
});
// ===== 404 Handler =====
app.use((req, res) => {
    res.status(404).json({
        error: 'Not Found',
        message: `Endpoint ${req.method} ${req.path} does not exist`,
        availableEndpoints: [
            'GET /health',
            'POST /template/placeholders',
            'POST /template/info',
            'POST /template/upload',
            'POST /generate/excel',
            'POST /generate/pdf',
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
    console.log(`✓ GUI Test Client:   http://localhost:${port}/`);
    console.log('');
    console.log('Available Endpoints:');
    console.log('  • GET  /health');
    console.log('  • POST /template/placeholders');
    console.log('  • POST /template/info');
    console.log('  • POST /template/upload');
    console.log('  • POST /generate/excel');
    console.log('  • POST /generate/pdf');
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
//# sourceMappingURL=server.js.map