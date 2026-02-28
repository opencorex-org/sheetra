import { 
  WorksheetData, 
//   RowData, 
//   ColumnData, 
  CellData, 
  MergeCell,
  FreezePane,
  PrintOptions 
} from '../types';
import { Row } from './row';
import { Column } from './column';

export class Worksheet {
  private name: string;
  private rows: Row[] = [];
  private columns: Column[] = [];
  private mergedCells: MergeCell[] = [];
  private freezePane?: FreezePane;
  private printOptions?: PrintOptions;

  constructor(name: string) {
    this.name = name;
  }

  getName(): string {
    return this.name;
  }

  setName(name: string): this {
    this.name = name;
    return this;
  }

  addRow(row: Row): this {
    this.rows.push(row);
    return this;
  }

  createRow(index?: number): Row {
    const row = new Row();
    if (index !== undefined) {
      this.rows.splice(index, 0, row);
    } else {
      this.rows.push(row);
    }
    return row;
  }

  getRow(index: number): Row | undefined {
    return this.rows[index];
  }

  getRows(): Row[] {
    return this.rows;
  }

  removeRow(index: number): boolean {
    if (index >= 0 && index < this.rows.length) {
      this.rows.splice(index, 1);
      return true;
    }
    return false;
  }

  addColumn(column: Column): this {
    this.columns.push(column);
    return this;
  }

  createColumn(width?: number): Column {
    const column = new Column(width);
    this.columns.push(column);
    return column;
  }

  getColumn(index: number): Column | undefined {
    return this.columns[index];
  }

  getColumns(): Column[] {
    return this.columns;
  }

  setCell(row: number, col: number, value: any, style?: any): this {
    while (this.rows.length <= row) {
      this.rows.push(new Row());
    }
    this.rows[row].setCell(col, value, style);
    return this;
  }

  getCell(row: number, col: number): CellData | undefined {
    const rowObj = this.rows[row];
    return rowObj?.getCell(col);
  }

  mergeCells(startRow: number, startCol: number, endRow: number, endCol: number): this {
    this.mergedCells.push({ startRow, startCol, endRow, endCol });
    return this;
  }

  setFreezePane(row?: number, col?: number): this {
    this.freezePane = { rows: row, columns: col };
    return this;
  }

  setPrintOptions(options: PrintOptions): this {
    this.printOptions = options;
    return this;
  }

  setOutlineLevel(row: number, level: number, collapsed: boolean = false): this {
    if (this.rows[row]) {
      this.rows[row].setOutlineLevel(level, collapsed);
    }
    return this;
  }

  setColumnOutlineLevel(col: number, level: number, collapsed: boolean = false): this {
    while (this.columns.length <= col) {
      this.columns.push(new Column());
    }
    this.columns[col].setOutlineLevel(level, collapsed);
    return this;
  }

  toData(): WorksheetData {
    return {
      name: this.name,
      rows: this.rows.map(row => row.toData()),
      columns: this.columns.map(col => col.toData()),
      mergeCells: this.mergedCells,
      freezePane: this.freezePane,
      printOptions: this.printOptions
    };
  }

  static fromData(data: WorksheetData): Worksheet {
    const sheet = new Worksheet(data.name);
    sheet.rows = data.rows.map(rowData => Row.fromData(rowData));
    sheet.columns = (data.columns || []).map(colData => Column.fromData(colData));
    sheet.mergedCells = data.mergeCells || [];
    sheet.freezePane = data.freezePane;
    sheet.printOptions = data.printOptions;
    return sheet;
  }
}