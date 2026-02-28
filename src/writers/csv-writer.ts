import { Workbook } from '../core/workbook';
import { Worksheet } from '../core/worksheet';
import { Row } from '../core/row';
import { ExportOptions } from '../types';

export class CSVWriter {
  static async write(workbook: Workbook, _options: ExportOptions): Promise<Blob> {
    const sheets = workbook['sheets'];
    const csvData = this.generateCSVData(sheets[0]); // Use first sheet for CSV
    
    return new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
  }

  private static generateCSVData(sheet: Worksheet): string {
    const rows = sheet.getRows();
    const csvRows: string[] = [];

    rows.forEach((row: Row) => {
      const rowData: string[] = [];
      row.getCells().forEach(cell => {
        const cellData = cell.toData();
        let value = '';
        
        if (cellData.value !== null && cellData.value !== undefined) {
          if (cellData.type === 'date' && cellData.value instanceof Date) {
            value = cellData.value.toISOString().split('T')[0];
          } else {
            value = String(cellData.value);
          }
        }
        
        // Escape quotes and wrap in quotes if contains comma or quote
        if (value.includes(',') || value.includes('"') || value.includes('\n')) {
          value = '"' + value.replace(/"/g, '""') + '"';
        }
        
        rowData.push(value);
      });
      
      csvRows.push(rowData.join(','));
    });

    return csvRows.join('\n');
  }
}