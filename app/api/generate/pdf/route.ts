import { NextRequest, NextResponse } from 'next/server';
import { PDFGenerator } from '@/lib/pdfGenerator';

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { templateBase64, data, options } = body;

    if (!templateBase64) {
      return NextResponse.json(
        { error: 'Template base64 data is required' },
        { status: 400 }
      );
    }

    if (!data) {
      return NextResponse.json(
        { error: 'Placeholder data is required' },
        { status: 400 }
      );
    }

    const generator = new PDFGenerator(options);

    // LibreOfficeがインストールされているか確認
    const isInstalled = await generator.checkLibreOfficeInstalled();
    if (!isInstalled) {
      return NextResponse.json(
        {
          error: 'LibreOffice is not installed',
          details: 'PDF generation requires LibreOffice. Please install it from https://www.libreoffice.org/download/download/',
          sofficeCommand: generator.getSofficeCommand(),
          installInstructions: {
            windows: 'Download and install from https://www.libreoffice.org/download/download/',
            macOS: 'Run: brew install --cask libreoffice',
            linux: 'Run: sudo apt-get install libreoffice',
          },
        },
        { status: 503 }
      );
    }

    const pdfBuffer = await generator.generatePDF(templateBase64, data, options || {});

    // Bufferをbase64に変換して返す
    const base64Result = pdfBuffer.toString('base64');

    return NextResponse.json({
      success: true,
      data: base64Result,
      mimeType: 'application/pdf',
    });
  } catch (error) {
    console.error('Error generating PDF:', error);
    return NextResponse.json(
      {
        error: 'Failed to generate PDF',
        details: error instanceof Error ? error.message : String(error),
      },
      { status: 500 }
    );
  }
}
