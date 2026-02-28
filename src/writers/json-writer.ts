import { Workbook } from '../core/workbook';
import { Worksheet } from '../core/worksheet';
// import { Row } from '../core/row';
import { ExportOptions } from '../types';

export class JSONWriter {
  static async write(workbook: Workbook, _options: ExportOptions): Promise<Blob> {
    const sheets = workbook['sheets'];
    const jsonData = this.generateJSONData(sheets[0]); // Use first sheet for JSON
    
    return new Blob([JSON.stringify(jsonData, null, 2)], { 
      type: 'application/json;charset=utf-8;' 
    });
  }

  private static generateJSONData(sheet: Worksheet): any[] {
    const rows = sheet.getRows();
    if (rows.length === 0) return [];

    // Assume first row contains headers
    const headers = rows[0].getCells().map((cell): string => {
      const value = cell.toData().value;
      return value ? String(value) : `Column${headers.length + 1}`;
    });

    const data: any[] = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const rowData: any = {};
      
      row.getCells().forEach((cell, index) => {
        if (index < headers.length) {
          const cellData = cell.toData();
          let value = cellData.value;
          
          if (cellData.type === 'date' && value instanceof Date) {
            value = value.toISOString();
          }
          
          rowData[headers[index]] = value;
        }
      });
      
      data.push(rowData);
    }

    return data;
  }
}