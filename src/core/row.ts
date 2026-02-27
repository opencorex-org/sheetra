import { RowData, CellData, CellStyle } from '../types';
import { Cell } from './cell';

export class Row {
  private cells: Cell[] = [];
  private height?: number;
  private hidden: boolean = false;
  private outlineLevel: number = 0;
  private collapsed: boolean = false;

  constructor(cells?: Cell[]) {
    if (cells) {
      this.cells = cells;
    }
  }

  addCell(cell: Cell): this {
    this.cells.push(cell);
    return this;
  }

  createCell(value?: any, style?: CellStyle): Cell {
    const cell = new Cell(value, style);
    this.cells.push(cell);
    return cell;
  }

  setCell(index: number, value: any, style?: CellStyle): this {
    while (this.cells.length <= index) {
      this.cells.push(new Cell());
    }
    this.cells[index].setValue(value).setStyle(style);
    return this;
  }

  getCell(index: number): CellData | undefined {
    return this.cells[index]?.toData();
  }

  getCells(): Cell[] {
    return this.cells;
  }

  setHeight(height: number): this {
    this.height = height;
    return this;
  }

  setHidden(hidden: boolean): this {
    this.hidden = hidden;
    return this;
  }

  setOutlineLevel(level: number, collapsed: boolean = false): this {
    this.outlineLevel = level;
    this.collapsed = collapsed;
    return this;
  }

  toData(): RowData {
    return {
      cells: this.cells.map(cell => cell.toData()),
      height: this.height,
      hidden: this.hidden,
      outlineLevel: this.outlineLevel,
      collapsed: this.collapsed
    };
  }

  static fromData(data: RowData): Row {
    const row = new Row(data.cells.map(cellData => Cell.fromData(cellData)));
    row.height = data.height;
    row.hidden = data.hidden || false;
    row.outlineLevel = data.outlineLevel || 0;
    row.collapsed = data.collapsed || false;
    return row;
  }
}