import ExcelJS from 'exceljs';

export interface PlaceholderData {
  [key: string]: string | number | Date | null;
}

export class PlaceholderReplacer {
  private placeholderPattern = /\{\{([^}]+)\}\}/g;

  /**
   * ワークブック内の全てのプレースホルダーを置換
   */
  async replacePlaceholders(
    workbook: ExcelJS.Workbook,
    data: PlaceholderData
  ): Promise<ExcelJS.Workbook> {
    workbook.eachSheet((worksheet) => {
      this.replaceInWorksheet(worksheet, data);
    });

    return workbook;
  }

  /**
   * ワークシート内のプレースホルダーを置換
   */
  private replaceInWorksheet(
    worksheet: ExcelJS.Worksheet,
    data: PlaceholderData
  ): void {
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        this.replaceInCell(cell, data);
      });
    });
  }

  /**
   * セル内のプレースホルダーを置換
   */
  private replaceInCell(cell: ExcelJS.Cell, data: PlaceholderData): void {
    const cellValue = cell.value;

    // セルの値が文字列の場合
    if (typeof cellValue === 'string') {
      const newValue = this.replacePlaceholderString(cellValue, data);
      if (newValue !== cellValue) {
        // スタイルを保持したまま値を更新
        const originalStyle = {
          font: cell.font,
          alignment: cell.alignment,
          border: cell.border,
          fill: cell.fill,
          numFmt: cell.numFmt,
        };

        cell.value = newValue;

        // スタイルを再適用
        cell.font = originalStyle.font;
        cell.alignment = originalStyle.alignment;
        cell.border = originalStyle.border;
        cell.fill = originalStyle.fill;
        cell.numFmt = originalStyle.numFmt;
      }
    }
    // リッチテキストの場合
    else if (cellValue && typeof cellValue === 'object' && 'richText' in cellValue) {
      const richText = cellValue as ExcelJS.CellRichTextValue;
      if (richText.richText) {
        richText.richText.forEach((segment) => {
          if (segment.text) {
            segment.text = this.replacePlaceholderString(segment.text, data);
          }
        });
        cell.value = richText;
      }
    }
    // 数式の場合
    else if (cellValue && typeof cellValue === 'object' && 'formula' in cellValue) {
      const formula = cellValue as ExcelJS.CellFormulaValue;
      if (formula.formula) {
        formula.formula = this.replacePlaceholderString(formula.formula, data);
        cell.value = formula;
      }
    }
  }

  /**
   * 文字列内のプレースホルダーを置換
   */
  private replacePlaceholderString(text: string, data: PlaceholderData): string {
    return text.replace(this.placeholderPattern, (match, fieldName) => {
      const trimmedFieldName = fieldName.trim();

      // データに該当するキーがあれば置換
      if (trimmedFieldName in data) {
        const value = data[trimmedFieldName];

        // null または undefined の場合は空文字列
        if (value === null || value === undefined) {
          return '';
        }

        // Date型の場合はフォーマット
        if (value instanceof Date) {
          return this.formatDate(value);
        }

        return String(value);
      }

      // データにキーが存在しない場合はそのまま返す
      return match;
    });
  }

  /**
   * 日付をフォーマット（yyyy/MM/dd形式）
   */
  private formatDate(date: Date): string {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}/${month}/${day}`;
  }

  /**
   * ワークブック内のプレースホルダーを検出
   */
  findPlaceholders(workbook: ExcelJS.Workbook): string[] {
    const placeholders = new Set<string>();

    workbook.eachSheet((worksheet) => {
      worksheet.eachRow({ includeEmpty: false }, (row) => {
        row.eachCell({ includeEmpty: false }, (cell) => {
          const cellValue = cell.value;

          if (typeof cellValue === 'string') {
            const matches = cellValue.matchAll(this.placeholderPattern);
            for (const match of matches) {
              placeholders.add(match[1].trim());
            }
          }
        });
      });
    });

    return Array.from(placeholders).sort();
  }

  /**
   * プレースホルダーの情報を取得（セルの位置も含む）
   */
  getPlaceholderInfo(workbook: ExcelJS.Workbook): Array<{
    placeholder: string;
    sheetName: string;
    cellAddress: string;
    cellValue: string;
  }> {
    const placeholderInfo: Array<{
      placeholder: string;
      sheetName: string;
      cellAddress: string;
      cellValue: string;
    }> = [];

    workbook.eachSheet((worksheet) => {
      worksheet.eachRow({ includeEmpty: false }, (row) => {
        row.eachCell({ includeEmpty: false }, (cell) => {
          const cellValue = cell.value;

          if (typeof cellValue === 'string') {
            const matches = cellValue.matchAll(this.placeholderPattern);
            for (const match of matches) {
              placeholderInfo.push({
                placeholder: match[1].trim(),
                sheetName: worksheet.name,
                cellAddress: cell.address,
                cellValue: cellValue,
              });
            }
          }
        });
      });
    });

    return placeholderInfo;
  }
}
