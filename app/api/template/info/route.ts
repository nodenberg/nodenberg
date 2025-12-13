import { NextRequest, NextResponse } from 'next/server';
import { ExcelGenerator } from '@/lib/excelGenerator';

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { templateBase64, sheetName } = body;

    if (!templateBase64) {
      return NextResponse.json(
        { error: 'Template base64 data is required' },
        { status: 400 }
      );
    }

    const generator = new ExcelGenerator();

    const templateInfo = await generator.getTemplateInfo(templateBase64);

    return NextResponse.json({
      success: true,
      templateInfo,
      note: 'Print settings are automatically preserved (test9 method)',
    });
  } catch (error) {
    console.error('Error getting template info:', error);
    return NextResponse.json(
      {
        error: 'Failed to get template info',
        details: error instanceof Error ? error.message : String(error),
      },
      { status: 500 }
    );
  }
}
