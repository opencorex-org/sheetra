import { Workbook } from '../core/workbook';
import { Worksheet } from '../core/worksheet';
// import { Row } from '../core/row';
// import { Cell } from '../core/cell';
import { StyleBuilder } from '../core/styles';
import { SectionConfig, ExportOptions } from '../types';
import { ExcelWriter, CSVWriter, JSONWriter } from '../writers';

export class ExportBuilder {
  private workbook: Workbook;
  private currentSheet: Worksheet;

  constructor(sheetName: string = 'Sheet1') {
    this.workbook = new Workbook();
    this.currentSheet = new Worksheet(sheetName);
    this.workbook.addSheet(this.currentSheet);
  }

  addHeaderRow(headers: string[], style?: any): this {
    const row = this.currentSheet.createRow();
    headers.forEach(header => {
      row.createCell(header, style);
    });
    return this;
  }

  /**
   * Add multiple data rows to the sheet
   * @param data Array of row data
   * @param fields Optional array of field names (for object data)
   * @param styles Optional array of styles per row or per cell
   */
  addDataRows(data: any[], fields?: string[], styles?: (import('../types').CellStyle | import('../types').CellStyle[] | undefined)[]): this {
    data.forEach((item, rowIdx) => {
      const row = this.currentSheet.createRow();
      let rowStyle: import('../types').CellStyle | undefined = undefined;
      let cellStyles: (import('../types').CellStyle | undefined)[] | undefined = undefined;
      if (styles && styles[rowIdx]) {
        if (Array.isArray(styles[rowIdx])) {
          cellStyles = styles[rowIdx] as import('../types').CellStyle[];
        } else {
          rowStyle = styles[rowIdx] as import('../types').CellStyle;
        }
      }

      if (fields && fields.length > 0) {
        fields.forEach((field, colIdx) => {
          const value = this.getNestedValue(item, field);
          const style = cellStyles ? cellStyles[colIdx] : rowStyle;
          row.createCell(value, style);
        });
      } else if (Array.isArray(item)) {
        item.forEach((value, colIdx) => {
          const style = cellStyles ? cellStyles[colIdx] : rowStyle;
          row.createCell(value, style);
        });
      } else if (typeof item === 'object') {
        Object.values(item).forEach((value, colIdx) => {
          const style = cellStyles ? cellStyles[colIdx] : rowStyle;
          row.createCell(value, style);
        });
      } else {
        row.createCell(item, rowStyle);
      }
    });
    return this;
  }

  addSection(config: SectionConfig): this {
    // Add section header
    const headerRow = this.currentSheet.createRow();
    headerRow.setOutlineLevel(config.level);
    
    const headerCell = headerRow.createCell(config.title);
    headerCell.setStyle(
      StyleBuilder.create()
        .bold(true)
        .backgroundColor('#E0E0E0')
        .build()
    );

    // Add section data
    if (config.data && config.data.length > 0) {
      if (config.groupBy && config.subSections) {
        // Handle grouped data
        const groupedData = this.groupData(config.data, config.groupBy);
        
        Object.entries(groupedData).forEach(([key, items]) => {
          const subHeaderRow = this.currentSheet.createRow();
          subHeaderRow.setOutlineLevel(config.level + 1);
          subHeaderRow.createCell(`${config.groupBy}: ${key}`);
          
          items.forEach((item: any) => {
            const dataRow = this.currentSheet.createRow();
            dataRow.setOutlineLevel(config.level + 2);
            
            if (config.fields) {
              config.fields.forEach((field: string) => {
                dataRow.createCell(this.getNestedValue(item, field));
              });
            }
          });
        });
      } else {
        // Add data rows with outline level
        config.data.forEach((item: any) => {
          const row = this.currentSheet.createRow();
          row.setOutlineLevel(config.level + 1);
          
          if (config.fields) {
            config.fields.forEach((field: string) => {
              row.createCell(this.getNestedValue(item, field));
            });
          }
        });
      }
    }

    return this;
  }

  addSections(sections: SectionConfig[]): this {
    sections.forEach(section => this.addSection(section));
    return this;
  }

  setColumnWidths(widths: number[]): this {
    widths.forEach((width) => {
      this.currentSheet.createColumn(width);
    });
    return this;
  }

  /**
   * Merge cells in the current worksheet
   * @param startRow Start row index (0-based)
   * @param startCol Start column index (0-based)
   * @param endRow End row index (0-based)
   * @param endCol End column index (0-based)
   */
  mergeCells(startRow: number, startCol: number, endRow: number, endCol: number): this {
    this.currentSheet.mergeCells(startRow, startCol, endRow, endCol);
    return this;
  }

  /**
   * Set alignment for a specific cell
   * @param row Row index (0-based)
   * @param col Column index (0-based)
   * @param horizontal Horizontal alignment
   * @param vertical Vertical alignment (optional)
   */
  setAlignment(
    row: number,
    col: number,
    horizontal: 'left' | 'center' | 'right',
    vertical?: 'top' | 'middle' | 'bottom'
  ): this {
    const rowObj = this.currentSheet.getRow(row);
    if (rowObj) {
      const cells = rowObj.getCells();
      if (cells[col]) {
        const style = StyleBuilder.create().align(horizontal);
        if (vertical) {
          style.verticalAlign(vertical);
        }
        cells[col].setStyle(style.build());
      }
    }
    return this;
  }

  /**
   * Set alignment for a range of cells
   * @param startRow Start row index (0-based)
   * @param startCol Start column index (0-based)
   * @param endRow End row index (0-based)
   * @param endCol End column index (0-based)
   * @param horizontal Horizontal alignment
   * @param vertical Vertical alignment (optional)
   */
  setRangeAlignment(
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number,
    horizontal: 'left' | 'center' | 'right',
    vertical?: 'top' | 'middle' | 'bottom'
  ): this {
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        this.setAlignment(r, c, horizontal, vertical);
      }
    }
    return this;
  }

  autoSizeColumns(): this {
    const rows = this.currentSheet.getRows();
    const maxLengths: number[] = [];

    rows.forEach(row => {
      row.getCells().forEach((cell, index) => {
        const value = cell.toData().value;
        const length = value ? String(value).length : 0;
        
        if (!maxLengths[index] || length > maxLengths[index]) {
          maxLengths[index] = Math.min(length, 50); // Cap at 50 characters
        }
      });
    });

    maxLengths.forEach((length) => {
      this.currentSheet.createColumn(Math.max(length, 10)); // Minimum width 10
    });

    return this;
  }

  addStyle(style: any): this {
    // Apply style to last row
    const rows = this.currentSheet.getRows();
    if (rows.length > 0) {
      const lastRow = rows[rows.length - 1];
      lastRow.getCells().forEach(cell => {
        cell.setStyle(style);
      });
    }
    return this;
  }

  private groupData(data: any[], field: string): Record<string, any[]> {
    return data.reduce((groups: Record<string, any[]>, item) => {
      const key = this.getNestedValue(item, field);
      if (!groups[key]) {
        groups[key] = [];
      }
      groups[key].push(item);
      return groups;
    }, {});
  }

  private getNestedValue(obj: any, path: string): any {
    return path.split('.').reduce((current, key) => {
      return current && current[key] !== undefined ? current[key] : undefined;
    }, obj);
  }

  build(): Workbook {
    return this.workbook;
  }

  async export(options: ExportOptions): Promise<Blob> {
    return this.workbook.export(options);
  }

  download(options: ExportOptions = {}) {
    const filename = options.filename || 'export.xlsx';
    const format = options.format || (
      filename.endsWith('.csv') ? 'csv' :
      filename.endsWith('.json') ? 'json' :
      'xlsx'
    );
    let writer;
    if (format === 'xlsx') writer = ExcelWriter;
    else if (format === 'csv') writer = CSVWriter;
    else if (format === 'json') writer = JSONWriter;
    else throw new Error('Unsupported format');

    writer.write(this.workbook, options).then((blob: Blob | MediaSource) => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 100);
    });
  }

  static create(sheetName?: string): ExportBuilder {
    return new ExportBuilder(sheetName);
  }
}