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

  addDataRows(data: any[], fields?: string[]): this {
    data.forEach(item => {
      const row = this.currentSheet.createRow();
      
      if (fields && fields.length > 0) {
        fields.forEach(field => {
          const value = this.getNestedValue(item, field);
          row.createCell(value);
        });
      } else if (Array.isArray(item)) {
        item.forEach(value => row.createCell(value));
      } else if (typeof item === 'object') {
        Object.values(item).forEach(value => row.createCell(value));
      } else {
        row.createCell(item);
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