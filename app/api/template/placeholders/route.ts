import { NextRequest, NextResponse } from 'next/server';
import { ExcelGenerator } from '@/lib/excelGenerator';

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { templateBase64, detailed } = body;

    if (!templateBase64) {
      return NextResponse.json(
        { error: 'Template base64 data is required' },
        { status: 400 }
      );
    }

    const generator = new ExcelGenerator();

    if (detailed) {
      const placeholderInfo = await generator.getPlaceholderInfo(templateBase64);
      return NextResponse.json({
        success: true,
        placeholders: placeholderInfo,
      });
    } else {
      const placeholders = await generator.findPlaceholders(templateBase64);
      return NextResponse.json({
        success: true,
        placeholders,
      });
    }
  } catch (error) {
    console.error('Error finding placeholders:', error);
    return NextResponse.json(
      {
        error: 'Failed to find placeholders',
        details: error instanceof Error ? error.message : String(error),
      },
      { status: 500 }
    );
  }
}
