import { Workbook } from '../core/workbook';
import { Worksheet } from '../core/worksheet';
import { Row } from '../core/row';
import { Cell } from '../core/cell';
import { ExportOptions, CellStyle } from '../types';
import { DateFormatter } from '../formatters/date-formatter';

/**
 * Style registry to track unique styles and their indices
 */
interface StyleRegistry {
    fonts: Map<string, number>;
    fills: Map<string, number>;
    borders: Map<string, number>;
    cellXfs: Map<string, number>;
}

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

        // Collect all styles from sheets
        const styleRegistry = this.collectStyles(sheets);

        // Add required XLSX files
        zip.addFile('[Content_Types].xml', this.generateContentTypes(sheetNames));
        zip.addFile('_rels/.rels', this.generateRels());
        zip.addFile('xl/workbook.xml', this.generateWorkbook(sheetNames));
        zip.addFile('xl/_rels/workbook.xml.rels', this.generateWorkbookRels(sheetNames));
        zip.addFile('xl/styles.xml', this.generateStyles(styleRegistry));
        zip.addFile('xl/sharedStrings.xml', this.generateSharedStrings(sheets));

        // Add each worksheet
        sheets.forEach((sheet, index) => {
            zip.addFile(`xl/worksheets/sheet${index + 1}.xml`, this.generateWorksheet(sheet, styleRegistry));
        });

        const buffer = zip.generate();
        return new Blob([buffer.buffer as ArrayBuffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
    }

    /**
     * Collect all unique styles from sheets
     */
    private static collectStyles(sheets: Worksheet[]): StyleRegistry {
        const registry: StyleRegistry = {
            fonts: new Map(),
            fills: new Map(),
            borders: new Map(),
            cellXfs: new Map()
        };

        // Add default font (Calibri 11)
        const defaultFont = this.serializeFont({});
        registry.fonts.set(defaultFont, 0);

        // Add required fills (none and gray125)
        registry.fills.set('none', 0);
        registry.fills.set('gray125', 1);

        // Add default border
        registry.borders.set('none', 0);

        // Add default cellXf (no style)
        registry.cellXfs.set('default', 0);

        // Collect styles from all cells
        sheets.forEach((sheet) => {
            const rows = sheet.getRows();
            rows.forEach((row: Row) => {
                row.getCells().forEach((cell: Cell) => {
                    const cellData = cell.toData();
                    if (cellData.style) {
                        this.registerStyle(cellData.style, registry);
                    }
                });
            });
        });

        return registry;
    }

    /**
     * Register a style and get its cellXf index
     */
    private static registerStyle(style: CellStyle, registry: StyleRegistry): number {
        // Serialize and register font
        const fontKey = this.serializeFont(style);
        if (!registry.fonts.has(fontKey)) {
            registry.fonts.set(fontKey, registry.fonts.size);
        }

        // Serialize and register fill
        const fillKey = this.serializeFill(style);
        if (fillKey !== 'none' && !registry.fills.has(fillKey)) {
            registry.fills.set(fillKey, registry.fills.size);
        }

        // Serialize and register border
        const borderKey = this.serializeBorder(style);
        if (borderKey !== 'none' && !registry.borders.has(borderKey)) {
            registry.borders.set(borderKey, registry.borders.size);
        }

        // Create cellXf key
        const fontIndex = registry.fonts.get(fontKey) ?? 0;
        const fillIndex = registry.fills.get(fillKey) ?? 0;
        const borderIndex = registry.borders.get(borderKey) ?? 0;
        let alignment = this.serializeAlignment(style);
        if (!alignment) alignment = '{}';
        const xfKey = `f${fontIndex}:l${fillIndex}:b${borderIndex}:a${alignment}`;
        
        if (!registry.cellXfs.has(xfKey)) {
            registry.cellXfs.set(xfKey, registry.cellXfs.size);
        }

        return registry.cellXfs.get(xfKey) ?? 0;
    }

    /**
     * Get style index for a cell style
     */
    private static getStyleIndex(style: CellStyle | undefined, registry: StyleRegistry): number {
        if (!style) return 0;
        return this.registerStyle(style, registry);
    }

    private static serializeFont(style: CellStyle): string {
        const bold = style.bold || style.font?.bold;
        const italic = style.italic || style.font?.italic;
        const underline = style.underline || style.font?.underline;
        const color = style.color || style.font?.color;
        const size = style.fontSize || style.font?.size || 11;
        const name = style.fontFamily || style.font?.name || 'Calibri';
        
        return JSON.stringify({ bold, italic, underline, color, size, name });
    }

    private static serializeFill(style: CellStyle): string {
        const bgColor = style.backgroundColor || style.fill?.fgColor;
        if (!bgColor) return 'none';
        return JSON.stringify({ bgColor });
    }

    private static serializeBorder(style: CellStyle): string {
        if (!style.border && !style.borderAll) return 'none';
        return JSON.stringify({ border: style.border, borderAll: style.borderAll });
    }

    private static serializeAlignment(style: CellStyle): string {
        const h = style.alignment;
        const v = style.verticalAlignment;
        const wrap = style.wrapText;
        if (!h && !v && !wrap) return '';
        return JSON.stringify({ h, v, wrap });
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

    private static generateStyles(registry: StyleRegistry): string {
        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
        xml += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
        
        // Generate fonts
        xml += `<fonts count="${registry.fonts.size}">`;
        const fontEntries = Array.from(registry.fonts.entries()).sort((a, b) => a[1] - b[1]);
        for (const [fontKey] of fontEntries) {
            const font = JSON.parse(fontKey);
            xml += '<font>';
            if (font.bold) xml += '<b/>';
            if (font.italic) xml += '<i/>';
            if (font.underline) xml += '<u/>';
            xml += `<sz val="${font.size || 11}"/>`;
            if (font.color) {
                const colorHex = this.colorToARGB(font.color);
                xml += `<color rgb="${colorHex}"/>`;
            }
            xml += `<name val="${font.name || 'Calibri'}"/>`;
            xml += '</font>';
        }
        xml += '</fonts>';

        // Generate fills
        xml += `<fills count="${registry.fills.size}">`;
        const fillEntries = Array.from(registry.fills.entries()).sort((a, b) => a[1] - b[1]);
        for (const [fillKey] of fillEntries) {
            if (fillKey === 'none') {
                xml += '<fill><patternFill patternType="none"/></fill>';
            } else if (fillKey === 'gray125') {
                xml += '<fill><patternFill patternType="gray125"/></fill>';
            } else {
                const fill = JSON.parse(fillKey);
                const colorHex = this.colorToARGB(fill.bgColor);
                xml += `<fill><patternFill patternType="solid"><fgColor rgb="${colorHex}"/><bgColor indexed="64"/></patternFill></fill>`;
            }
        }
        xml += '</fills>';

        // Generate borders
        xml += `<borders count="${registry.borders.size}">`;
        const borderEntries = Array.from(registry.borders.entries()).sort((a, b) => a[1] - b[1]);
        for (const [borderKey] of borderEntries) {
            if (borderKey === 'none') {
                xml += '<border><left/><right/><top/><bottom/><diagonal/></border>';
            } else {
                const borderData = JSON.parse(borderKey);
                xml += '<border>';
                
                if (borderData.borderAll) {
                    const style = borderData.borderAll.style || 'thin';
                    const color = borderData.borderAll.color ? this.colorToARGB(borderData.borderAll.color) : 'FF000000';
                    xml += `<left style="${style}"><color rgb="${color}"/></left>`;
                    xml += `<right style="${style}"><color rgb="${color}"/></right>`;
                    xml += `<top style="${style}"><color rgb="${color}"/></top>`;
                    xml += `<bottom style="${style}"><color rgb="${color}"/></bottom>`;
                } else if (borderData.border) {
                    const b = borderData.border;
                    xml += this.generateBorderEdge('left', b.left);
                    xml += this.generateBorderEdge('right', b.right);
                    xml += this.generateBorderEdge('top', b.top);
                    xml += this.generateBorderEdge('bottom', b.bottom);
                }
                
                xml += '<diagonal/></border>';
            }
        }
        xml += '</borders>';

        // Generate cellStyleXfs (base styles)
        xml += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';

        // Generate cellXfs (cell formats)
        xml += `<cellXfs count="${registry.cellXfs.size}">`;
        const xfEntries = Array.from(registry.cellXfs.entries()).sort((a, b) => a[1] - b[1]);
        for (const [xfKey] of xfEntries) {
            if (xfKey === 'default') {
                xml += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
            } else {
                // Parse xf key: f{fontIndex}:l{fillIndex}:b{borderIndex}:a{alignment}
                const parts = xfKey.split(':');
                const fontId = parseInt(parts[0].substring(1)) || 0;
                const fillId = parseInt(parts[1].substring(1)) || 0;
                const borderId = parseInt(parts[2].substring(1)) || 0;
                const alignmentJson = parts[3].substring(1);

                let xf = `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"`;
                if (fontId > 0) xf += ' applyFont="1"';
                if (fillId > 0) xf += ' applyFill="1"';
                if (borderId > 0) xf += ' applyBorder="1"';

                let align: any = {};
                if (alignmentJson && alignmentJson.trim().startsWith('{')) {
                    try {
                        align = JSON.parse(alignmentJson);
                    } catch (e) {
                        align = {};
                    }
                }
                if (Object.keys(align).length > 0) {
                    xf += ' applyAlignment="1">';
                    xf += '<alignment';
                    if (align.h) xf += ` horizontal="${align.h}"`;
                    if (align.v) xf += ` vertical="${align.v === 'middle' ? 'center' : align.v}"`;
                    if (align.wrap) xf += ' wrapText="1"';
                    xf += '/></xf>';
                } else {
                    xf += '/>';
                }
                xml += xf;
            }
        }
        xml += '</cellXfs>';
        
        xml += '</styleSheet>';
        return xml;
    }

    private static generateBorderEdge(side: string, edge?: { style?: string; color?: string }): string {
        if (!edge || !edge.style) return `<${side}/>`;
        const color = edge.color ? this.colorToARGB(edge.color) : 'FF000000';
        return `<${side} style="${edge.style}"><color rgb="${color}"/></${side}>`;
    }

    private static colorToARGB(color: string): string {
        if (!color) return 'FF000000';
        // Remove # if present
        let hex = color.replace('#', '').toUpperCase();
        // Add alpha if not present (6 chars -> 8 chars)
        if (hex.length === 6) {
            hex = 'FF' + hex;
        }
        // Handle 3 char hex
        if (hex.length === 3) {
            hex = 'FF' + hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
        }
        return hex;
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

    private static generateWorksheet(sheet: Worksheet, styleRegistry: StyleRegistry): string {
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
                    const styleIndex = this.getStyleIndex(cellData.style, styleRegistry);
                    const styleAttr = styleIndex > 0 ? ` s="${styleIndex}"` : '';
                    
                    if (cellData.formula) {
                        xml += `<c r="${cellRef}"${styleAttr}><f>${this.escapeXml(cellData.formula)}</f></c>`;
                    } else if (cellData.value !== null && cellData.value !== undefined && cellData.value !== '') {
                        if (typeof cellData.value === 'number') {
                            xml += `<c r="${cellRef}"${styleAttr}><v>${cellData.value}</v></c>`;
                        } else if (typeof cellData.value === 'boolean') {
                            xml += `<c r="${cellRef}"${styleAttr} t="b"><v>${cellData.value ? 1 : 0}</v></c>`;
                        } else if (cellData.type === 'date') {
                            const excelDate = DateFormatter.toExcelDate(cellData.value);
                            xml += `<c r="${cellRef}"${styleAttr}><v>${excelDate}</v></c>`;
                        } else {
                            // String value - use shared string index
                            const stringIndex = allStrings.indexOf(String(cellData.value));
                            if (stringIndex >= 0) {
                                xml += `<c r="${cellRef}"${styleAttr} t="s"><v>${stringIndex}</v></c>`;
                            } else {
                                // Inline string as fallback
                                xml += `<c r="${cellRef}"${styleAttr} t="inlineStr"><is><t>${this.escapeXml(String(cellData.value))}</t></is></c>`;
                            }
                        }
                    } else if (styleIndex > 0) {
                        // Empty cell with style
                        xml += `<c r="${cellRef}"${styleAttr}/>`;
                    }
                });
                
                xml += '</row>';
            }
        });
        
        xml += '</sheetData>';
        
        // Add merged cells if any
        const sheetData = sheet.toData();
        if (sheetData.mergeCells && sheetData.mergeCells.length > 0) {
            xml += '<mergeCells>';
            sheetData.mergeCells.forEach(merge => {
                const startRef = this.columnToLetter(merge.startCol + 1) + (merge.startRow + 1);
                const endRef = this.columnToLetter(merge.endCol + 1) + (merge.endRow + 1);
                xml += `<mergeCell ref="${startRef}:${endRef}"/>`;
            });
            xml += '</mergeCells>';
        }
        
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