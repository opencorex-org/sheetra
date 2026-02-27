import { Worksheet } from '../core/worksheet';
import { Row } from '../core/row';
import { Cell } from '../core/cell';
import { Column } from '../core/column';
import { StyleBuilder } from '../core/styles';
import { 
  CellStyle, 
  CellData, 
  RowData, 
  ColumnData, 
  MergeCell,
  FreezePane,
  PrintOptions,
  HeaderFooter,
  AutoFilter,
  TableConfig,
  DataValidation,
  ConditionalFormatRule,
  ChartConfig,
  Drawing
} from '../types';

/**
 * SheetBuilder provides a fluent API for building Excel worksheets
 */
export class SheetBuilder {
  private worksheet: Worksheet;
  private currentRow: number = 0;
  private currentCol: number = 0;

  /**
   * Create a new SheetBuilder instance
   * @param name Worksheet name
   */
  constructor(name: string = 'Sheet1') {
    this.worksheet = new Worksheet(name);
  }

  /**
   * Get the underlying worksheet
   */
  getWorksheet(): Worksheet {
    return this.worksheet;
  }

  /**
   * Set the worksheet name
   */
  setName(name: string): this {
    this.worksheet.setName(name);
    return this;
  }

  /**
   * Add a header row with styling
   * @param headers Array of header text
   * @param style Optional style for all headers
   */
  addHeaderRow(headers: string[], style?: CellStyle): this {
    const row = this.createRow();
    
    headers.forEach((header, index) => {
      const cellStyle = style || StyleBuilder.create()
        .bold(true)
        .backgroundColor('#F0F0F0')
        .borderAll('thin')
        .align('center')
        .build();
      
      row.createCell(header, cellStyle);
    });

    return this;
  }

  /**
   * Add a title row
   * @param title Title text
   * @param colSpan Number of columns to span
   */
  addTitle(title: string, colSpan: number = 1): this {
    const row = this.createRow();
    const cell = row.createCell(title);
    
    cell.setStyle(
      StyleBuilder.create()
        .bold(true)
        .fontSize(14)
        .align('center')
        .build()
    );

    if (colSpan > 1) {
      this.mergeCells(this.currentRow - 1, 0, this.currentRow - 1, colSpan - 1);
    }

    return this;
  }

  /**
   * Add a subtitle row
   * @param subtitle Subtitle text
   * @param colSpan Number of columns to span
   */
  addSubtitle(subtitle: string, colSpan: number = 1): this {
    const row = this.createRow();
    const cell = row.createCell(subtitle);
    
    cell.setStyle(
      StyleBuilder.create()
        .italic(true)
        .color('#666666')
        .align('center')
        .build()
    );

    if (colSpan > 1) {
      this.mergeCells(this.currentRow - 1, 0, this.currentRow - 1, colSpan - 1);
    }

    return this;
  }

  /**
   * Add a row of data
   * @param data Array of cell values
   * @param styles Optional array of styles per cell
   */
  addRow(data: any[], styles?: (CellStyle | undefined)[]): this {
    const row = this.createRow();
    
    data.forEach((value, index) => {
      const style = styles && styles[index] ? styles[index] : undefined;
      row.createCell(value, style);
    });

    return this;
  }

  /**
   * Add multiple rows of data
   * @param rows Array of row data
   * @param styles Optional array of styles per row or per cell
   */
  addRows(rows: any[][], styles?: (CellStyle | CellStyle[] | undefined)[]): this {
    rows.forEach((rowData, rowIndex) => {
      const rowStyles = styles && styles[rowIndex];
      
      if (Array.isArray(rowStyles)) {
        this.addRow(rowData, rowStyles);
      } else {
        this.addRow(rowData, rowStyles ? [rowStyles] : undefined);
      }
    });

    return this;
  }

  /**
   * Add data from objects
   * @param data Array of objects
   * @param fields Fields to extract (keys or dot notation paths)
   * @param headers Optional header labels
   */
  addObjects<T extends Record<string, any>>(
    data: T[],
    fields: (keyof T | string)[],
    headers?: string[]
  ): this {
    // Add headers if provided
    if (headers) {
      this.addHeaderRow(headers);
    }

    // Add data rows
    data.forEach(item => {
      const rowData = fields.map(field => this.getNestedValue(item, field as string));
      this.addRow(rowData);
    });

    return this;
  }

  /**
   * Add a section with header and data
   * @param title Section title
   * @param data Section data
   * @param fields Fields to display
   * @param level Outline level
   */
  addSection(
    title: string,
    data: any[],
    fields: string[],
    level: number = 0
  ): this {
    // Add section header
    const headerRow = this.createRow();
    headerRow.setOutlineLevel(level);
    
    const headerCell = headerRow.createCell(title);
    headerCell.setStyle(
      StyleBuilder.create()
        .bold(true)
        .fontSize(12)
        .backgroundColor('#E0E0E0')
        .borderAll('thin')
        .build()
    );

    // Merge header across all columns
    if (fields.length > 1) {
      this.mergeCells(this.currentRow - 1, 0, this.currentRow - 1, fields.length - 1);
    }

    // Add field headers
    const fieldRow = this.createRow();
    fieldRow.setOutlineLevel(level + 1);
    
    fields.forEach(field => {
      fieldRow.createCell(
        this.formatFieldName(field),
        StyleBuilder.create()
          .bold(true)
          .backgroundColor('#F5F5F5')
          .borderAll('thin')
          .build()
      );
    });

    // Add data rows
    data.forEach(item => {
      const row = this.createRow();
      row.setOutlineLevel(level + 2);
      
      fields.forEach(field => {
        const value = this.getNestedValue(item, field);
        row.createCell(value);
      });
    });

    return this;
  }

  /**
   * Add grouped data with sub-sections
   * @param data Data to group
   * @param groupBy Field to group by
   * @param fields Fields to display
   * @param level Outline level
   */
  addGroupedData(
    data: any[],
    groupBy: string,
    fields: string[],
    level: number = 0
  ): this {
    const groups = this.groupData(data, groupBy);
    
    Object.entries(groups).forEach(([key, items]) => {
      // Add group header
      const groupRow = this.createRow();
      groupRow.setOutlineLevel(level);
      
      const groupCell = groupRow.createCell(`${groupBy}: ${key}`);
      groupCell.setStyle(
        StyleBuilder.create()
          .bold(true)
          .backgroundColor('#E8E8E8')
          .build()
      );

      if (fields.length > 1) {
        this.mergeCells(this.currentRow - 1, 0, this.currentRow - 1, fields.length - 1);
      }

      // Add field headers
      const fieldRow = this.createRow();
      fieldRow.setOutlineLevel(level + 1);
      
      fields.forEach(field => {
        fieldRow.createCell(
          this.formatFieldName(field),
          StyleBuilder.create()
            .bold(true)
            .backgroundColor('#F5F5F5')
            .build()
        );
      });

      // Add items
      items.forEach((item: any) => {
        const row = this.createRow();
        row.setOutlineLevel(level + 2);
        
        fields.forEach(field => {
          const value = this.getNestedValue(item, field);
          row.createCell(value);
        });
      });

      // Add group summary
      this.addGroupSummary(items, fields, level + 1);
    });

    return this;
  }

  /**
   * Add summary row for a group
   */
  private addGroupSummary(items: any[], fields: string[], level: number): this {
    const summaryRow = this.createRow();
    summaryRow.setOutlineLevel(level + 1);
    
    // Create summary cells
    fields.forEach((field, index) => {
      if (index === 0) {
        summaryRow.createCell(
          'Group Summary',
          StyleBuilder.create()
            .italic(true)
            .bold(true)
            .build()
        );
      } else {
        const numericValues = items
          .map(item => this.getNestedValue(item, field))
          .filter(val => typeof val === 'number');
        
        if (numericValues.length > 0) {
          const sum = numericValues.reduce((a, b) => a + b, 0);
          summaryRow.createCell(
            sum,
            StyleBuilder.create()
              .bold(true)
              .numberFormat('#,##0.00')
              .build()
          );
        } else {
          summaryRow.createCell('');
        }
      }
    });

    return this;
  }

  /**
   * Add a summary row with totals
   * @param fields Fields to summarize
   * @param functions Summary functions (sum, average, count, min, max)
   * @param label Summary label
   */
  addSummaryRow(
    fields: string[],
    functions: Array<'sum' | 'average' | 'count' | 'min' | 'max'>,
    label: string = 'Total'
  ): this {
    const row = this.createRow();
    
    fields.forEach((field, index) => {
      if (index === 0) {
        row.createCell(
          label,
          StyleBuilder.create()
            .bold(true)
            .borderTop('double')
            .build()
        );
      } else {
        const function_ = functions[index - 1] || 'sum';
        row.createCell(
          `=${function_.toUpperCase()}(${this.getColumnRange(index)})`,
          StyleBuilder.create()
            .bold(true)
            .borderTop('double')
            .numberFormat('#,##0.00')
            .build()
        );
      }
    });

    return this;
  }

  /**
   * Set column widths
   * @param widths Array of column widths
   */
  setColumnWidths(widths: number[]): this {
    widths.forEach((width, index) => {
      const column = this.getOrCreateColumn(index);
      column.setWidth(width);
    });

    return this;
  }

  /**
   * Auto-size columns based on content
   * @param maxWidth Maximum width in characters
   */
  autoSizeColumns(maxWidth: number = 50): this {
    const rows = this.worksheet.getRows();
    const columnWidths: Map<number, number> = new Map();

    rows.forEach(row => {
      row.getCells().forEach((cell, index) => {
        const cellData = cell.toData();
        const value = cellData.value;
        const length = value ? String(value).length : 0;
        
        const currentMax = columnWidths.get(index) || 0;
        columnWidths.set(index, Math.min(Math.max(length, currentMax), maxWidth));
      });
    });

    columnWidths.forEach((width, index) => {
      const column = this.getOrCreateColumn(index);
      column.setWidth(Math.max(width + 2, 8)); // Add padding, minimum 8
    });

    return this;
  }

  /**
   * Set column to auto-fit
   * @param colIndex Column index
   */
  setColumnAutoFit(colIndex: number): this {
    const column = this.getOrCreateColumn(colIndex);
    // Auto-fit will be applied during export
    column.setWidth(0); // 0 indicates auto-fit
    return this;
  }

  /**
   * Hide a column
   * @param colIndex Column index
   */
  hideColumn(colIndex: number): this {
    const column = this.getOrCreateColumn(colIndex);
    column.setHidden(true);
    return this;
  }

  /**
   * Hide a row
   * @param rowIndex Row index
   */
  hideRow(rowIndex: number): this {
    const row = this.worksheet.getRow(rowIndex);
    if (row) {
      row.setHidden(true);
    }
    return this;
  }

  /**
   * Set outline level for a row
   * @param rowIndex Row index
   * @param level Outline level (0-7)
   * @param collapsed Whether the outline is collapsed
   */
  setRowOutlineLevel(rowIndex: number, level: number, collapsed: boolean = false): this {
    const row = this.worksheet.getRow(rowIndex);
    if (row) {
      row.setOutlineLevel(level, collapsed);
    }
    return this;
  }

  /**
   * Set outline level for a column
   * @param colIndex Column index
   * @param level Outline level (0-7)
   * @param collapsed Whether the outline is collapsed
   */
  setColumnOutlineLevel(colIndex: number, level: number, collapsed: boolean = false): this {
    const column = this.getOrCreateColumn(colIndex);
    column.setOutlineLevel(level, collapsed);
    return this;
  }

  /**
   * Create an outline group for rows
   * @param startRow Start row index
   * @param endRow End row index
   * @param level Outline level
   * @param collapsed Whether the group is collapsed
   */
  groupRows(startRow: number, endRow: number, level: number = 1, collapsed: boolean = false): this {
    for (let i = startRow; i <= endRow; i++) {
      const row = this.worksheet.getRow(i);
      if (row) {
        row.setOutlineLevel(level, collapsed && i === startRow);
      }
    }
    return this;
  }

  /**
   * Create an outline group for columns
   * @param startCol Start column index
   * @param endCol End column index
   * @param level Outline level
   * @param collapsed Whether the group is collapsed
   */
  groupColumns(startCol: number, endCol: number, level: number = 1, collapsed: boolean = false): this {
    for (let i = startCol; i <= endCol; i++) {
      const column = this.getOrCreateColumn(i);
      column.setOutlineLevel(level, collapsed && i === startCol);
    }
    return this;
  }

  /**
   * Merge cells
   * @param startRow Start row
   * @param startCol Start column
   * @param endRow End row
   * @param endCol End column
   */
  mergeCells(startRow: number, startCol: number, endRow: number, endCol: number): this {
    this.worksheet.mergeCells(startRow, startCol, endRow, endCol);
    return this;
  }

  /**
   * Freeze panes
   * @param rows Number of rows to freeze
   * @param columns Number of columns to freeze
   */
  freezePanes(rows: number = 0, columns: number = 0): this {
    this.worksheet.setFreezePane(rows, columns);
    return this;
  }

  /**
   * Set print options
   * @param options Print options
   */
  setPrintOptions(options: PrintOptions): this {
    this.worksheet.setPrintOptions(options);
    return this;
  }

  /**
   * Set header and footer
   * @param headerFooter Header/footer configuration
   */
  setHeaderFooter(headerFooter: HeaderFooter): this {
    // Store in worksheet (would be implemented in worksheet class)
    (this.worksheet as any).headerFooter = headerFooter;
    return this;
  }

  /**
   * Add auto-filter
   * @param startRow Start row
   * @param startCol Start column
   * @param endRow End row
   * @param endCol End column
   */
  addAutoFilter(startRow: number, startCol: number, endRow: number, endCol: number): this {
    const range = `${this.columnToLetter(startCol)}${startRow + 1}:${this.columnToLetter(endCol)}${endRow + 1}`;
    (this.worksheet as any).autoFilter = { range };
    return this;
  }

  /**
   * Add a table
   * @param config Table configuration
   */
  addTable(config: TableConfig): this {
    if (!(this.worksheet as any).tables) {
      (this.worksheet as any).tables = [];
    }
    (this.worksheet as any).tables.push(config);
    return this;
  }

  /**
   * Add data validation to a cell or range
   * @param range Cell range (e.g., 'A1:B10')
   * @param validation Data validation rules
   */
  addDataValidation(range: string, validation: DataValidation): this {
    if (!(this.worksheet as any).dataValidations) {
      (this.worksheet as any).dataValidations = {};
    }
    (this.worksheet as any).dataValidations[range] = validation;
    return this;
  }

  /**
   * Add conditional formatting
   * @param range Cell range
   * @param rules Conditional formatting rules
   */
  addConditionalFormatting(range: string, rules: ConditionalFormatRule[]): this {
    if (!(this.worksheet as any).conditionalFormats) {
      (this.worksheet as any).conditionalFormats = {};
    }
    (this.worksheet as any).conditionalFormats[range] = rules;
    return this;
  }

  /**
   * Add a chart
   * @param config Chart configuration
   */
  addChart(config: ChartConfig): this {
    if (!(this.worksheet as any).charts) {
      (this.worksheet as any).charts = [];
    }
    (this.worksheet as any).charts.push(config);
    return this;
  }

  /**
   * Add an image or drawing
   * @param drawing Drawing configuration
   */
  addDrawing(drawing: Drawing): this {
    if (!(this.worksheet as any).drawings) {
      (this.worksheet as any).drawings = [];
    }
    (this.worksheet as any).drawings.push(drawing);
    return this;
  }

  /**
   * Set cell value with style
   * @param row Row index
   * @param col Column index
   * @param value Cell value
   * @param style Cell style
   */
  setCell(row: number, col: number, value: any, style?: CellStyle): this {
    this.worksheet.setCell(row, col, value, style);
    return this;
  }

  /**
   * Get cell value
   * @param row Row index
   * @param col Column index
   */
  getCell(row: number, col: number): CellData | undefined {
    return this.worksheet.getCell(row, col);
  }

  /**
   * Add a comment to a cell
   * @param row Row index
   * @param col Column index
   * @param comment Comment text
   * @param author Comment author
   */
  addComment(row: number, col: number, comment: string, author?: string): this {
    const cell = this.worksheet.getCell(row, col);
    if (cell) {
      // Would need to extend Cell class to support comments
      (cell as any).comment = { text: comment, author };
    }
    return this;
  }

  /**
   * Add a hyperlink to a cell
   * @param row Row index
   * @param col Column index
   * @param url Hyperlink URL
   * @param displayText Display text (optional)
   */
  addHyperlink(row: number, col: number, url: string, displayText?: string): this {
    this.worksheet.setCell(row, col, displayText || url);
    const cell = this.worksheet.getCell(row, col);
    if (cell) {
      (cell as any).hyperlink = url;
    }
    return this;
  }

  /**
   * Apply a style to a range
   * @param startRow Start row
   * @param startCol Start column
   * @param endRow End row
   * @param endCol End column
   * @param style Style to apply
   */
  applyStyleToRange(
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number,
    style: CellStyle
  ): this {
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const cell = this.worksheet.getCell(r, c);
        if (cell) {
          // Would need to update cell style
          (cell as any).style = { ...(cell as any).style, ...style };
        }
      }
    }
    return this;
  }

  /**
   * Insert a blank row
   * @param count Number of blank rows to insert
   */
  insertBlankRows(count: number = 1): this {
    for (let i = 0; i < count; i++) {
      this.createRow();
    }
    return this;
  }

  /**
   * Create a new row
   */
  private createRow(): Row {
    const row = this.worksheet.createRow();
    this.currentRow++;
    return row;
  }

  /**
   * Get or create a column
   */
  private getOrCreateColumn(index: number): Column {
    let column = this.worksheet.getColumn(index);
    if (!column) {
      column = this.worksheet.createColumn();
    }
    return column;
  }

  /**
   * Get nested value from object using dot notation
   */
  private getNestedValue(obj: any, path: string): any {
    return path.split('.').reduce((current, key) => {
      return current && current[key] !== undefined ? current[key] : undefined;
    }, obj);
  }

  /**
   * Format field name for display
   */
  private formatFieldName(field: string): string {
    return field
      .split('.')
      .pop()!
      .split('_')
      .map(word => word.charAt(0).toUpperCase() + word.slice(1))
      .join(' ');
  }

  /**
   * Group data by field
   */
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

  /**
   * Convert column index to letter (A, B, C, ...)
   */
  private columnToLetter(column: number): string {
    let letters = '';
    while (column >= 0) {
      letters = String.fromCharCode((column % 26) + 65) + letters;
      column = Math.floor(column / 26) - 1;
    }
    return letters;
  }

  /**
   * Get column range for formula (e.g., A:A)
   */
  private getColumnRange(colIndex: number): string {
    const colLetter = this.columnToLetter(colIndex);
    return `${colLetter}:${colLetter}`;
  }

  /**
   * Build and return the worksheet
   */
  build(): Worksheet {
    return this.worksheet;
  }

  /**
   * Reset the builder
   */
  reset(): this {
    this.worksheet = new Worksheet(this.worksheet.getName());
    this.currentRow = 0;
    this.currentCol = 0;
    return this;
  }

  /**
   * Create a new SheetBuilder instance
   */
  static create(name?: string): SheetBuilder {
    return new SheetBuilder(name);
  }

  /**
   * Create from existing worksheet
   */
  static fromWorksheet(worksheet: Worksheet): SheetBuilder {
    const builder = new SheetBuilder(worksheet.getName());
    builder.worksheet = worksheet;
    return builder;
  }
}

// Add missing types to CellStyle (to be added in cell.types.ts)
declare module '../types/cell.types' {
  interface CellStyle {
    borderTop?: BorderStyleType | BorderEdge;
    borderBottom?: BorderStyleType | BorderEdge;
    borderLeft?: BorderStyleType | BorderEdge;
    borderRight?: BorderStyleType | BorderEdge;
  }
}

// Helper function to add border methods to StyleBuilder
export function extendStyleBuilder() {
  StyleBuilder.prototype.borderTop = function(style: any, color?: string) {
    if (!this['style'].border) this['style'].border = {};
    this['style'].border.top = typeof style === 'string' 
      ? { style, color } 
      : style;
    return this;
  };

  StyleBuilder.prototype.borderBottom = function(style: any, color?: string) {
    if (!this['style'].border) this['style'].border = {};
    this['style'].border.bottom = typeof style === 'string' 
      ? { style, color } 
      : style;
    return this;
  };

  StyleBuilder.prototype.borderLeft = function(style: any, color?: string) {
    if (!this['style'].border) this['style'].border = {};
    this['style'].border.left = typeof style === 'string' 
      ? { style, color } 
      : style;
    return this;
  };

  StyleBuilder.prototype.borderRight = function(style: any, color?: string) {
    if (!this['style'].border) this['style'].border = {};
    this['style'].border.right = typeof style === 'string' 
      ? { style, color } 
      : style;
    return this;
  };
}