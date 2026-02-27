// Core exports
export { Workbook } from './core/workbook';
export { Worksheet } from './core/worksheet';
export { Row } from './core/row';
export { Cell } from './core/cell';
export { Column } from './core/column';
export { StyleBuilder } from './core/styles';

// Builder exports
export { ExportBuilder } from './builders/export-builder';
export { SheetBuilder } from './builders/sheet-builder';
export { SectionBuilder } from './builders/section-builder';

// Writer exports
export { ExcelWriter } from './writers/excel-writer';
export { CSVWriter } from './writers/csv-writer';
export { JSONWriter } from './writers/json-writer';

// Formatter exports
export { DateFormatter } from './formatters/date-formatter';
export { NumberFormatter } from './formatters/number-formatter';

// Type exports
export * from './types';

// Utility exports
export * from './utils/helpers';

// Main export function for backward compatibility
import { ExportBuilder } from './builders/export-builder';
import { Workbook } from './core/workbook';

export function exportToExcel(data: any[], options?: any): Workbook {
  const builder = ExportBuilder.create(options?.sheetName || 'Sheet1');
  
  if (options?.headers) {
    builder.addHeaderRow(options.headers);
  }
  
  builder.addDataRows(data, options?.fields);
  
  if (options?.columnWidths) {
    builder.setColumnWidths(options.columnWidths);
  }
  
  return builder.build();
}

export function exportToCSV(data: any[], options?: any): string {
  // Simplified CSV export
  const headers = options?.headers || Object.keys(data[0] || {});
  const rows = [headers];
  
  data.forEach(item => {
    const row = headers.map((header: string | number) => item[header] || '');
    rows.push(row);
  });
  
  return rows.map(row => 
    row.map((cell: any) => 
      String(cell).includes(',') ? `"${cell}"` : cell
    ).join(',')
  ).join('\n');
}

// Version
export const VERSION = '1.0.0';