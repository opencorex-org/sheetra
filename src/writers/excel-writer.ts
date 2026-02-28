import { Workbook } from '../core/workbook';
import { Worksheet } from '../core/worksheet';
import { Row } from '../core/row';
import { Cell } from '../core/cell';
import { ExportOptions } from '../types';
import { DateFormatter } from '../formatters/date-formatter';

export class ExcelWriter {
    static async write(workbook: Workbook, _options: ExportOptions): Promise<Blob> {
        const data = this.generateExcelData(workbook);
        const buffer = this.createExcelFile(data);
        const uint8Buffer = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
        const arrayBuffer = uint8Buffer.slice().buffer;
        return new Blob([arrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    }

    private static generateExcelData(workbook: Workbook): any {
        const sheets = workbook['sheets'];

        return {
            Sheets: sheets.reduce((acc: any, sheet: Worksheet, index: number) => {
                const sheetName = sheet.getName() || `Sheet${index + 1}`;
                acc[sheetName] = this.generateSheetData(sheet);
                return acc;
            }, {}),
            SheetNames: sheets.map((sheet: Worksheet, index: number) => sheet.getName() || `Sheet${index + 1}`)
        };
    }

    private static generateSheetData(sheet: Worksheet): any {
        const rows = sheet.getRows();
        const data: any[][] = [];

        rows.forEach((row: Row) => {
            const rowData: any[] = [];
            row.getCells().forEach((cell: Cell) => {
                const cellData = cell.toData();

                if (cellData.formula) {
                    rowData.push({ f: cellData.formula });
                } else if (
                    cellData.type === 'date' &&
                    (typeof cellData.value === 'string' ||
                     typeof cellData.value === 'number' ||
                     cellData.value instanceof Date)
                ) {
                    rowData.push(DateFormatter.toExcelDate(cellData.value));
                } else {
                    rowData.push(cellData.value);
                }
            });
            data.push(rowData);
        });

        return {
            '!ref': this.getSheetRange(data),
            '!rows': rows.map(row => ({
                hpt: row['height'],
                hidden: row['hidden'],
                outlineLevel: row['outlineLevel'],
                collapsed: row['collapsed']
            })),
            '!cols': sheet.getColumns().map(col => ({
                wch: col['width'],
                hidden: col['hidden'],
                outlineLevel: col['outlineLevel'],
                collapsed: col['collapsed']
            })),
            '!merges': sheet['mergeCells'],
            '!freeze': sheet['freezePane']
        };
    }

    private static getSheetRange(data: any[][]): string {
        if (data.length === 0) return 'A1:A1';

        const lastRow = data.length;
        const lastCol = Math.max(...data.map(row => row.length));

        return `A1:${this.columnToLetter(lastCol)}${lastRow}`;
    }

    private static columnToLetter(column: number): string {
        let letters = '';
        while (column > 0) {
            const temp = (column - 1) % 26;
            letters = String.fromCharCode(temp + 65) + letters;
            column = (column - temp - 1) / 26;
        }
        return letters || 'A';
    }

    private static createExcelFile(data: any): Uint8Array {
        // This is a simplified Excel file generator
        // In a real implementation, you'd generate proper XLSX binary format
        // For now, we'll return a simple XML-based structure

        const xmlHeader = '<?xml version="1.0" encoding="UTF-8"?>';
        const workbook = this.generateWorkbookXML(data);

        const encoder = new TextEncoder();
        return encoder.encode(xmlHeader + workbook);
    }

    private static generateWorkbookXML(data: any): string {
        let xml = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
        xml += '<sheets>';

        data.SheetNames.forEach((name: string, index: number) => {
            xml += `<sheet name="${name}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`;
        });

        xml += '</sheets>';
        xml += '</workbook>';

        return xml;
    }
}