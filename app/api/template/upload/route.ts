import { NextRequest, NextResponse } from 'next/server';
import { ExcelTemplateManager } from '@/lib/excelTemplateManager';

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { templateId, templateName, base64Data, generateJsonTemplate } = body;

    if (!templateId || !templateName || !base64Data) {
      return NextResponse.json(
        { error: 'Template ID, name, and base64 data are required' },
        { status: 400 }
      );
    }

    const manager = new ExcelTemplateManager();
    const template = await manager.uploadTemplate(
      templateId,
      templateName,
      base64Data,
      generateJsonTemplate || false
    );

    return NextResponse.json({
      success: true,
      template: {
        id: template.id,
        name: template.name,
        uploadedAt: template.uploadedAt,
        hasJsonTemplate: !!template.jsonTemplate,
      },
    });
  } catch (error) {
    console.error('Error uploading template:', error);
    return NextResponse.json(
      {
        error: 'Failed to upload template',
        details: error instanceof Error ? error.message : String(error),
      },
      { status: 500 }
    );
  }
}
