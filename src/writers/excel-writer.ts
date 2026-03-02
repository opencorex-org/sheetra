import { Workbook } from '../core/workbook';
import { Worksheet } from '../core/worksheet';
import { Row } from '../core/row';
import { Cell } from '../core/cell';
import { ExportOptions } from '../types';
import { DateFormatter } from '../formatters/date-formatter';

/**
 * Minimal ZIP file creator for XLSX generation (no external dependencies)
 * Uses STORE method (no compression) for simplicity and compatibility
 */
class ZipWriter {
    private files: { name: string; data: Uint8Array }[] = [];
    private encoder = new TextEncoder();

    addFile(name: string, content: string): void {
        this.files.push({
            name,
            data: this.encoder.encode(content)
        });
    }

    generate(): Uint8Array {
        const localFiles: Uint8Array[] = [];
        const centralDirectory: Uint8Array[] = [];
        let offset = 0;

        for (const file of this.files) {
            const localHeader = this.createLocalFileHeader(file.name, file.data);
            localFiles.push(localHeader);
            localFiles.push(file.data);

            const centralHeader = this.createCentralDirectoryHeader(file.name, file.data, offset);
            centralDirectory.push(centralHeader);

            offset += localHeader.length + file.data.length;
        }

        const centralDirStart = offset;
        let centralDirSize = 0;
        for (const header of centralDirectory) {
            centralDirSize += header.length;
        }

        const endOfCentralDir = this.createEndOfCentralDirectory(
            this.files.length,
            centralDirSize,
            centralDirStart
        );

        // Combine all parts
        const totalSize = offset + centralDirSize + endOfCentralDir.length;
        const result = new Uint8Array(totalSize);
        let pos = 0;

        for (const local of localFiles) {
            result.set(local, pos);
            pos += local.length;
        }
        for (const central of centralDirectory) {
            result.set(central, pos);
            pos += central.length;
        }
        result.set(endOfCentralDir, pos);

        return result;
    }

    private createLocalFileHeader(name: string, data: Uint8Array): Uint8Array {
        const nameBytes = this.encoder.encode(name);
        const crc = this.crc32(data);
        const header = new Uint8Array(30 + nameBytes.length);
        const view = new DataView(header.buffer);

        view.setUint32(0, 0x04034b50, true);  // Local file header signature
        view.setUint16(4, 20, true);           // Version needed to extract
        view.setUint16(6, 0, true);             // General purpose bit flag
        view.setUint16(8, 0, true);             // Compression method (STORE)
        view.setUint16(10, 0, true);            // File last modification time
        view.setUint16(12, 0, true);            // File last modification date
        view.setUint32(14, crc, true);          // CRC-32
        view.setUint32(18, data.length, true);  // Compressed size
        view.setUint32(22, data.length, true);  // Uncompressed size
        view.setUint16(26, nameBytes.length, true); // File name length
        view.setUint16(28, 0, true);            // Extra field length

        header.set(nameBytes, 30);
        return header;
    }

    private createCentralDirectoryHeader(name: string, data: Uint8Array, localHeaderOffset: number): Uint8Array {
        const nameBytes = this.encoder.encode(name);
        const crc = this.crc32(data);
        const header = new Uint8Array(46 + nameBytes.length);
        const view = new DataView(header.buffer);

        view.setUint32(0, 0x02014b50, true);   // Central directory file header signature
        view.setUint16(4, 20, true);            // Version made by
        view.setUint16(6, 20, true);            // Version needed to extract
        view.setUint16(8, 0, true);             // General purpose bit flag
        view.setUint16(10, 0, true);            // Compression method (STORE)
        view.setUint16(12, 0, true);            // File last modification time
        view.setUint16(14, 0, true);            // File last modification date
        view.setUint32(16, crc, true);          // CRC-32
        view.setUint32(20, data.length, true);  // Compressed size
        view.setUint32(24, data.length, true);  // Uncompressed size
        view.setUint16(28, nameBytes.length, true); // File name length
        view.setUint16(30, 0, true);            // Extra field length
        view.setUint16(32, 0, true);            // File comment length
        view.setUint16(34, 0, true);            // Disk number start
        view.setUint16(36, 0, true);            // Internal file attributes
        view.setUint32(38, 0, true);            // External file attributes
        view.setUint32(42, localHeaderOffset, true); // Relative offset of local header

        header.set(nameBytes, 46);
        return header;
    }

    private createEndOfCentralDirectory(numFiles: number, centralDirSize: number, centralDirOffset: number): Uint8Array {
        const header = new Uint8Array(22);
        const view = new DataView(header.buffer);

        view.setUint32(0, 0x06054b50, true);    // End of central directory signature
        view.setUint16(4, 0, true);              // Number of this disk
        view.setUint16(6, 0, true);              // Disk where central directory starts
        view.setUint16(8, numFiles, true);       // Number of central directory records on this disk
        view.setUint16(10, numFiles, true);      // Total number of central directory records
        view.setUint32(12, centralDirSize, true); // Size of central directory
        view.setUint32(16, centralDirOffset, true); // Offset of start of central directory
        view.setUint16(20, 0, true);             // Comment length

        return header;
    }

    private crc32(data: Uint8Array): number {
        let crc = 0xFFFFFFFF;
        for (let i = 0; i < data.length; i++) {
            crc ^= data[i];
            for (let j = 0; j < 8; j++) {
                crc = (crc >>> 1) ^ (crc & 1 ? 0xEDB88320 : 0);
            }
        }
        return (crc ^ 0xFFFFFFFF) >>> 0;
    }
}

export class ExcelWriter {
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    static async write(workbook: Workbook, _options: ExportOptions): Promise<Blob> {
        const zip = new ZipWriter();
        const sheets = workbook['sheets'] as Worksheet[];
        const sheetNames = sheets.map((sheet: Worksheet, index: number) => 
            sheet.getName() || `Sheet${index + 1}`
        );

        // Add required XLSX files
        zip.addFile('[Content_Types].xml', this.generateContentTypes(sheetNames));
        zip.addFile('_rels/.rels', this.generateRels());
        zip.addFile('xl/workbook.xml', this.generateWorkbook(sheetNames));
        zip.addFile('xl/_rels/workbook.xml.rels', this.generateWorkbookRels(sheetNames));
        zip.addFile('xl/styles.xml', this.generateStyles());
        zip.addFile('xl/sharedStrings.xml', this.generateSharedStrings(sheets));

        // Add each worksheet
        sheets.forEach((sheet, index) => {
            zip.addFile(`xl/worksheets/sheet${index + 1}.xml`, this.generateWorksheet(sheet));
        });

        const buffer = zip.generate();
        return new Blob([buffer.buffer as ArrayBuffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
    }

    private static escapeXml(str: string): string {
        if (typeof str !== 'string') return String(str ?? '');
        return str
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    private static generateContentTypes(sheetNames: string[]): string {
        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
        xml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        xml += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        xml += '<Default Extension="xml" ContentType="application/xml"/>';
        xml += '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
        xml += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
        xml += '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
        
        sheetNames.forEach((_, index) => {
            xml += `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
        });
        
        xml += '</Types>';
        return xml;
    }

    private static generateRels(): string {
        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
        xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        xml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>';
        xml += '</Relationships>';
        return xml;
    }

    private static generateWorkbook(sheetNames: string[]): string {
        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
        xml += '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
        xml += '<sheets>';
        
        sheetNames.forEach((name, index) => {
            xml += `<sheet name="${this.escapeXml(name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`;
        });
        
        xml += '</sheets>';
        xml += '</workbook>';
        return xml;
    }

    private static generateWorkbookRels(sheetNames: string[]): string {
        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
        xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        
        sheetNames.forEach((_, index) => {
            xml += `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`;
        });
        
        xml += `<Relationship Id="rId${sheetNames.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
        xml += `<Relationship Id="rId${sheetNames.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`;
        xml += '</Relationships>';
        return xml;
    }

    private static generateStyles(): string {
        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
        xml += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
        xml += '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>';
        xml += '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>';
        xml += '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>';
        xml += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';
        xml += '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>';
        xml += '</styleSheet>';
        return xml;
    }

    private static generateSharedStrings(sheets: Worksheet[]): string {
        const strings: string[] = [];
        
        sheets.forEach((sheet) => {
            const rows = sheet.getRows();
            rows.forEach((row: Row) => {
                row.getCells().forEach((cell: Cell) => {
                    const cellData = cell.toData();
                    if (typeof cellData.value === 'string' && !cellData.formula) {
                        if (!strings.includes(cellData.value)) {
                            strings.push(cellData.value);
                        }
                    }
                });
            });
        });

        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
        xml += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">`;
        
        strings.forEach(str => {
            xml += `<si><t>${this.escapeXml(str)}</t></si>`;
        });
        
        xml += '</sst>';
        return xml;
    }

    private static generateWorksheet(sheet: Worksheet): string {
        const rows = sheet.getRows();
        const columns = sheet.getColumns();
        const allStrings = this.collectAllStrings(sheet);
        
        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
        xml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
        
        // Add sheet dimension
        const dimension = this.getSheetDimension(rows);
        xml += `<dimension ref="${dimension}"/>`;
        
        // Add column widths if defined
        if (columns.length > 0) {
            xml += '<cols>';
            columns.forEach((col, index) => {
                const colData = col.toData();
                if (colData.width !== undefined && colData.width > 0) {
                    // Excel width is in character units, convert from pixels: width / 7
                    const excelWidth = colData.width / 7;
                    xml += `<col min="${index + 1}" max="${index + 1}" width="${excelWidth}" customWidth="1"`;
                    if (colData.hidden) {
                        xml += ' hidden="1"';
                    }
                    xml += '/>'; 
                }
            });
            xml += '</cols>';
        }
        
        xml += '<sheetData>';
        
        rows.forEach((row: Row, rowIndex: number) => {
            const cells = row.getCells();
            if (cells.length > 0) {
                xml += `<row r="${rowIndex + 1}">`;
                
                cells.forEach((cell: Cell, colIndex: number) => {
                    const cellData = cell.toData();
                    const cellRef = this.columnToLetter(colIndex + 1) + (rowIndex + 1);
                    
                    if (cellData.formula) {
                        xml += `<c r="${cellRef}"><f>${this.escapeXml(cellData.formula)}</f></c>`;
                    } else if (cellData.value !== null && cellData.value !== undefined && cellData.value !== '') {
                        if (typeof cellData.value === 'number') {
                            xml += `<c r="${cellRef}"><v>${cellData.value}</v></c>`;
                        } else if (typeof cellData.value === 'boolean') {
                            xml += `<c r="${cellRef}" t="b"><v>${cellData.value ? 1 : 0}</v></c>`;
                        } else if (cellData.type === 'date') {
                            const excelDate = DateFormatter.toExcelDate(cellData.value);
                            xml += `<c r="${cellRef}"><v>${excelDate}</v></c>`;
                        } else {
                            // String value - use shared string index
                            const stringIndex = allStrings.indexOf(String(cellData.value));
                            if (stringIndex >= 0) {
                                xml += `<c r="${cellRef}" t="s"><v>${stringIndex}</v></c>`;
                            } else {
                                // Inline string as fallback
                                xml += `<c r="${cellRef}" t="inlineStr"><is><t>${this.escapeXml(String(cellData.value))}</t></is></c>`;
                            }
                        }
                    }
                });
                
                xml += '</row>';
            }
        });
        
        xml += '</sheetData>';
        xml += '</worksheet>';
        return xml;
    }

    private static collectAllStrings(sheet: Worksheet): string[] {
        const strings: string[] = [];
        const rows = sheet.getRows();
        
        rows.forEach((row: Row) => {
            row.getCells().forEach((cell: Cell) => {
                const cellData = cell.toData();
                if (typeof cellData.value === 'string' && !cellData.formula) {
                    if (!strings.includes(cellData.value)) {
                        strings.push(cellData.value);
                    }
                }
            });
        });
        
        return strings;
    }

    private static getSheetDimension(rows: Row[]): string {
        if (rows.length === 0) return 'A1';
        
        let maxCol = 1;
        rows.forEach(row => {
            const cellCount = row.getCells().length;
            if (cellCount > maxCol) maxCol = cellCount;
        });
        
        return `A1:${this.columnToLetter(maxCol)}${rows.length}`;
    }

    private static columnToLetter(column: number): string {
        let letters = '';
        while (column > 0) {
            const temp = (column - 1) % 26;
            letters = String.fromCharCode(temp + 65) + letters;
            column = Math.floor((column - temp - 1) / 26);
        }
        return letters || 'A';
    }
}