import { NextRequest, NextResponse } from 'next/server';
import { ExcelGenerator } from '@/lib/excelGenerator';

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { templateBase64, data } = body;

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

    const generator = new ExcelGenerator();
    const excelBuffer = await generator.generateExcel(templateBase64, data);

    // Bufferをbase64に変換して返す
    const base64Result = excelBuffer.toString('base64');

    return NextResponse.json({
      success: true,
      data: base64Result,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  } catch (error) {
    console.error('Error generating Excel:', error);
    return NextResponse.json(
      {
        error: 'Failed to generate Excel',
        details: error instanceof Error ? error.message : String(error),
      },
      { status: 500 }
    );
  }
}
