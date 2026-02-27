import { CellData, CellStyle } from '../types';
import { DateFormatter, NumberFormatter } from '../formatters';

export class Cell {
  private value: any = null;
  private style?: CellStyle;
  private formula?: string;
  private type: 'string' | 'number' | 'date' | 'boolean' | 'formula' = 'string';

  constructor(value?: any, style?: CellStyle) {
    if (value !== undefined) {
      this.setValue(value);
    }
    if (style) {
      this.setStyle(style);
    }
  }

  setValue(value: any): this {
    this.value = value;
    this.inferType();
    return this;
  }

  setStyle(style?: CellStyle): this {
    this.style = style;
    return this;
  }

  setFormula(formula: string): this {
    this.formula = formula;
    this.type = 'formula';
    return this;
  }

  setType(type: 'string' | 'number' | 'date' | 'boolean' | 'formula'): this {
    this.type = type;
    return this;
  }

  private inferType(): void {
    if (this.value instanceof Date) {
      this.type = 'date';
    } else if (typeof this.value === 'number') {
      this.type = 'number';
    } else if (typeof this.value === 'boolean') {
      this.type = 'boolean';
    } else {
      this.type = 'string';
    }
  }

  getFormattedValue(): string {
    if (this.value === null || this.value === undefined) {
      return '';
    }

    switch (this.type) {
      case 'date':
        const dateFormat = typeof this.style?.numberFormat === 'string'
          ? this.style.numberFormat
          : (this.style?.numberFormat as any)?.format;
        return DateFormatter.format(this.value, dateFormat);
      case 'number':
        const numberFormat = typeof this.style?.numberFormat === 'string'
          ? this.style.numberFormat
          : (this.style?.numberFormat as any)?.format;
        return NumberFormatter.format(this.value, numberFormat);
      case 'boolean':
        return this.value ? 'TRUE' : 'FALSE';
      case 'formula':
        return this.formula || '';
      default:
        return String(this.value);
    }
  }

  toData(): CellData {
    return {
      value: this.value,
      style: this.style,
      formula: this.formula,
      type: this.type
    };
  }

  static fromData(data: CellData): Cell {
    const cell = new Cell();
    cell.value = data.value;
    cell.style = data.style;
    cell.formula = data.formula;
    const validTypes = ['string', 'number', 'boolean', 'date', 'formula'] as const;
    cell.type = validTypes.includes(data.type as any) ? data.type as typeof validTypes[number] : 'string';
    
    if (cell.value && !cell.type) {
      cell.inferType();
    }
    
    return cell;
  }
}