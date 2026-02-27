import { ColumnData } from '../types';

export class Column {
  private width?: number;
  private hidden: boolean = false;
  private outlineLevel: number = 0;
  private collapsed: boolean = false;

  constructor(width?: number) {
    this.width = width;
  }

  setWidth(width: number): this {
    this.width = width;
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

  toData(): ColumnData {
    return {
      width: this.width,
      hidden: this.hidden,
      outlineLevel: this.outlineLevel,
      collapsed: this.collapsed
    };
  }

  static fromData(data: ColumnData): Column {
    const column = new Column(data.width);
    column.hidden = data.hidden || false;
    column.outlineLevel = data.outlineLevel || 0;
    column.collapsed = data.collapsed || false;
    return column;
  }
}