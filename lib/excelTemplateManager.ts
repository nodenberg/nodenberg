import ExcelJS from 'exceljs';

export interface ExcelTemplate {
  id: string;
  name: string;
  base64Data: string;
  jsonTemplate?: any;
  uploadedAt: Date;
}

export class ExcelTemplateManager {
  /**
   * Base64文字列からExcelJSワークブックを読み込む
   */
  async loadWorkbookFromBase64(base64Data: string): Promise<ExcelJS.Workbook> {
    const workbook = new ExcelJS.Workbook();

    // Base64をBufferに変換
    const buffer = Buffer.from(base64Data, 'base64');

    // ワークブックを読み込む
    await workbook.xlsx.load(buffer as any);

    return workbook;
  }

  /**
   * エクセルテンプレートをアップロードし、必要に応じてJSONテンプレートを生成
   */
  async uploadTemplate(
    templateId: string,
    templateName: string,
    base64Data: string,
    generateJsonTemplate: boolean = false
  ): Promise<ExcelTemplate> {
    const template: ExcelTemplate = {
      id: templateId,
      name: templateName,
      base64Data,
      uploadedAt: new Date(),
    };

    if (generateJsonTemplate) {
      const workbook = await this.loadWorkbookFromBase64(base64Data);
      template.jsonTemplate = await this.generateJsonTemplate(workbook);
    }

    return template;
  }

  /**
   * ワークブックからJSONテンプレートを生成
   * ライブラリによってはJSON形式のテンプレートが必要な場合に使用
   */
  private async generateJsonTemplate(workbook: ExcelJS.Workbook): Promise<any> {
    const jsonTemplate: any = {
      sheets: [],
      metadata: {
        creator: workbook.creator || '',
        lastModifiedBy: workbook.lastModifiedBy || '',
        created: workbook.created,
        modified: workbook.modified,
      },
    };

    workbook.eachSheet((worksheet, sheetId) => {
      const sheetData: any = {
        id: sheetId,
        name: worksheet.name,
        state: worksheet.state,
        properties: {
          tabColor: worksheet.properties.tabColor,
          outlineLevelCol: worksheet.properties.outlineLevelCol,
          outlineLevelRow: worksheet.properties.outlineLevelRow,
          defaultRowHeight: worksheet.properties.defaultRowHeight,
          defaultColWidth: worksheet.properties.defaultColWidth,
        },
        pageSetup: {
          paperSize: worksheet.pageSetup.paperSize,
          orientation: worksheet.pageSetup.orientation,
          horizontalCentered: worksheet.pageSetup.horizontalCentered,
          verticalCentered: worksheet.pageSetup.verticalCentered,
          margins: worksheet.pageSetup.margins,
          printArea: worksheet.pageSetup.printArea,
          printTitlesRow: worksheet.pageSetup.printTitlesRow,
          printTitlesColumn: worksheet.pageSetup.printTitlesColumn,
        },
        views: worksheet.views,
        columns: [] as any[],
        rows: [] as any[],
      };

      // 列情報を保存
      worksheet.columns.forEach((column, index) => {
        if (column) {
          sheetData.columns.push({
            key: column.key,
            width: column.width,
            hidden: column.hidden,
            outlineLevel: column.outlineLevel,
          });
        }
      });

      // 行とセル情報を保存
      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const rowData: any = {
          number: rowNumber,
          height: row.height,
          hidden: row.hidden,
          outlineLevel: row.outlineLevel,
          cells: [] as any[],
        };

        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const cellData: any = {
            address: cell.address,
            value: cell.value,
            type: cell.type,
            style: {
              font: cell.font,
              alignment: cell.alignment,
              border: cell.border,
              fill: cell.fill,
              numFmt: cell.numFmt,
            },
            merge: cell.isMerged ? cell.master?.address : undefined,
          };

          rowData.cells.push(cellData);
        });

        sheetData.rows.push(rowData);
      });

      jsonTemplate.sheets.push(sheetData);
    });

    return jsonTemplate;
  }

  /**
   * ワークブックのシート情報を取得
   */
  getWorkbookInfo(workbook: ExcelJS.Workbook) {
    const sheets = workbook.worksheets.map((worksheet) => ({
      id: worksheet.id,
      name: worksheet.name,
      state: worksheet.state,
      rowCount: worksheet.rowCount,
      columnCount: worksheet.columnCount,
      actualRowCount: worksheet.actualRowCount,
      actualColumnCount: worksheet.actualColumnCount,
    }));

    return {
      sheetCount: workbook.worksheets.length,
      sheets,
      creator: workbook.creator,
      created: workbook.created,
      modified: workbook.modified,
    };
  }

  /**
   * ワークブックの印刷設定を取得
   */
  getPrintSettings(worksheet: ExcelJS.Worksheet) {
    return {
      paperSize: worksheet.pageSetup.paperSize,
      orientation: worksheet.pageSetup.orientation,
      horizontalCentered: worksheet.pageSetup.horizontalCentered,
      verticalCentered: worksheet.pageSetup.verticalCentered,
      margins: worksheet.pageSetup.margins,
      printArea: worksheet.pageSetup.printArea,
      printTitlesRow: worksheet.pageSetup.printTitlesRow,
      printTitlesColumn: worksheet.pageSetup.printTitlesColumn,
      fitToPage: worksheet.pageSetup.fitToPage,
      fitToHeight: worksheet.pageSetup.fitToHeight,
      fitToWidth: worksheet.pageSetup.fitToWidth,
      scale: worksheet.pageSetup.scale,
    };
  }
}
