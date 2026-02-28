import { WorkbookData, ExportOptions } from '../types';
import { Worksheet } from './worksheet';
import { ExcelWriter } from '../writers/excel-writer';
import { CSVWriter } from '../writers/csv-writer';
import { JSONWriter } from '../writers/json-writer';

export class Workbook {
  private sheets: Worksheet[] = [];
  private properties: Record<string, any> = {};

  constructor(data?: WorkbookData) {
    if (data) {
      this.fromData(data);
    }
  }

  addSheet(sheet: Worksheet): this {
    this.sheets.push(sheet);
    return this;
  }

  createSheet(name: string): Worksheet {
    const sheet = new Worksheet(name);
    this.sheets.push(sheet);
    return sheet;
  }

  getSheet(index: number): Worksheet | undefined {
    return this.sheets[index];
  }

  getSheetByName(name: string): Worksheet | undefined {
    return this.sheets.find(s => s.getName() === name);
  }

  removeSheet(index: number): boolean {
    if (index >= 0 && index < this.sheets.length) {
      this.sheets.splice(index, 1);
      return true;
    }
    return false;
  }

  setProperty(key: string, value: any): this {
    this.properties[key] = value;
    return this;
  }

  toData(): WorkbookData {
    return {
      sheets: this.sheets.map(sheet => sheet.toData()),
      properties: this.properties
    };
  }

  fromData(data: WorkbookData): void {
    this.sheets = data.sheets.map(sheetData => Worksheet.fromData(sheetData));
    this.properties = data.properties || {};
  }

  async export(options: ExportOptions = {}): Promise<Blob> {
    const { format = 'xlsx' } = options;

    switch (format) {
      case 'csv':
        return CSVWriter.write(this, options);
      case 'json':
        return JSONWriter.write(this, options);
      case 'xlsx':
      default:
        return ExcelWriter.write(this, options);
    }
  }

  download(options: ExportOptions = {}): void {
    const { filename = 'workbook.xlsx' } = options;
    
    this.export(options).then(blob => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.click();
      window.URL.revokeObjectURL(url);
    });
  }

  static create(): Workbook {
    return new Workbook();
  }

  static fromData(data: WorkbookData): Workbook {
    return new Workbook(data);
  }
}