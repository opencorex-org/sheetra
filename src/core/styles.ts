import { CellStyle, Border, BorderStyle } from '../types';

export class StyleBuilder {
  private style: CellStyle = {};

  bold(bold: boolean = true): this {
    this.style.bold = bold;
    return this;
  }

  italic(italic: boolean = true): this {
    this.style.italic = italic;
    return this;
  }

  underline(underline: boolean = true): this {
    this.style.underline = underline;
    return this;
  }

  fontSize(size: number): this {
    this.style.fontSize = size;
    return this;
  }

  fontFamily(family: string): this {
    this.style.fontFamily = family;
    return this;
  }

  color(color: string): this {
    this.style.color = color;
    return this;
  }

  backgroundColor(color: string): this {
    this.style.backgroundColor = color;
    return this;
  }

  align(alignment: 'left' | 'center' | 'right'): this {
    this.style.alignment = alignment;
    return this;
  }

  verticalAlign(alignment: 'top' | 'middle' | 'bottom'): this {
    this.style.verticalAlignment = alignment;
    return this;
  }

  wrapText(wrap: boolean = true): this {
    this.style.wrapText = wrap;
    return this;
  }

  border(border: Border): this {
    this.style.border = border;
    return this;
  }

  borderTop(style: BorderStyle['style'] = 'thin', color?: string): this {
    if (!this.style.border) this.style.border = {};
    this.style.border.top = { style, color };
    return this;
  }

  borderRight(style: BorderStyle['style'] = 'thin', color?: string): this {
    if (!this.style.border) this.style.border = {};
    this.style.border.right = { style, color };
    return this;
  }

  borderBottom(style: BorderStyle['style'] = 'thin', color?: string): this {
    if (!this.style.border) this.style.border = {};
    this.style.border.bottom = { style, color };
    return this;
  }

  borderLeft(style: BorderStyle['style'] = 'thin', color?: string): this {
    if (!this.style.border) this.style.border = {};
    this.style.border.left = { style, color };
    return this;
  }

  borderAll(style: BorderStyle['style'] = 'thin', color?: string): this {
    const borderStyle: BorderStyle = { style, color };
    this.style.border = {
      top: borderStyle,
      bottom: borderStyle,
      left: borderStyle,
      right: borderStyle
    };
    return this;
  }

  numberFormat(format: string): this {
    this.style.numberFormat = format;
    return this;
  }

  build(): CellStyle {
    return { ...this.style };
  }

  static create(): StyleBuilder {
    return new StyleBuilder();
  }
}