import { Worksheet } from '../core/worksheet';
// import { Row } from '../core/row';
import { Cell } from '../core/cell';
import { StyleBuilder } from '../core/styles';
import { 
  SectionConfig, 
  CellStyle, 
  ExportFilters,
  ConditionalFormatRule 
} from '../types';

/**
 * Builder class for creating collapsible sections in worksheets
 * Supports nested sections, grouping, summaries, and conditional styling
 */
export class SectionBuilder {
  private worksheet: Worksheet;
  private currentRow: number = 0;
  private sections: Map<string, Section> = new Map();
  private styles: Map<string, CellStyle> = new Map();
  private formatters: Map<string, (value: any) => any> = new Map();
  private conditionalFormats: ConditionalFormatRule[] = [];

  constructor(worksheet: Worksheet) {
    this.worksheet = worksheet;
  }

  /**
   * Add a section to the worksheet
   */
  addSection(config: SectionConfig): this {
    const section = new Section(this.worksheet, config, this.currentRow);
    if (config.title !== undefined) {
      this.sections.set(config.title, section);
    }
    
    // Build the section
    this.currentRow = section.build();
    
    return this;
  }

  /**
   * Add multiple sections at once
   */
  addSections(configs: SectionConfig[]): this {
    configs.forEach(config => this.addSection(config));
    return this;
  }

  /**
   * Add a nested section (child of current section)
   */
  addNestedSection(parentTitle: string, config: SectionConfig): this {
    const parent = this.sections.get(parentTitle);
    if (parent) {
      parent.addSubSection(config);
    }
    return this;
  }

  /**
   * Create a section from data with automatic grouping
   */
  createFromData<T>(
    data: T[],
    options: {
      title: string;
      groupBy?: keyof T | ((item: T) => string);
      fields?: Array<keyof T | string>;
      fieldLabels?: Record<string, string>;
      level?: number;
      collapsed?: boolean;
      summary?: {
        fields: Array<keyof T>;
        functions: Array<'sum' | 'average' | 'count' | 'min' | 'max'>;
      };
    }
  ): this {
    const {
      title,
      groupBy,
      fields = Object.keys(data[0] || {}) as Array<keyof T>,
      fieldLabels = {},
      level = 0,
      collapsed = false,
      summary
    } = options;

    // If no grouping, create a simple section
    if (!groupBy) {
      return this.addSection({
        title,
        level,
        collapsed,
        data,
        fields: fields as string[],
        fieldLabels,
        summary: summary ? {
          fields: summary.fields as string[],
          function: summary.functions[0] || 'count',
          label: 'Total'
        } : undefined
      });
    }

    // Group the data
    const groups = this.groupData(data, groupBy);
    
    // Create main section with groups as subsections
    const sectionConfig: SectionConfig = {
      title,
      level,
      collapsed,
      data: [],
      subSections: Object.entries(groups).map(([key, items]) => ({
        title: `${String(groupBy)}: ${key}`,
        level: level + 1,
        collapsed: true,
        data: items,
        fields: fields as string[],
        fieldLabels,
        summary: summary ? {
          fields: summary.fields as string[],
          function: summary.functions[0] || 'count',
          label: `Total for ${key}`
        } : undefined
      }))
    };

    return this.addSection(sectionConfig);
  }

  /**
   * Add a summary section (totals, averages, etc.)
   */
  addSummarySection(
    data: any[],
    fields: string[],
    functions: Array<'sum' | 'average' | 'count' | 'min' | 'max'>,
    options?: {
      level?: number;
      style?: CellStyle;
      showPercentage?: boolean;
      label?: string;
    }
  ): this {
    const { level = 0, style, showPercentage = false, label = 'Summary' } = options || {};

    // Calculate summary values
    const summary: Record<string, any> = {};
    
    fields.forEach((field, index) => {
      const func = functions[index] || functions[0];
      const values = data.map(item => this.getNestedValue(item, field)).filter(v => v != null);
      
      switch (func) {
        case 'sum':
          summary[field] = values.reduce((a, b) => a + b, 0);
          break;
        case 'average':
          summary[field] = values.reduce((a, b) => a + b, 0) / values.length;
          break;
        case 'count':
          summary[field] = values.length;
          break;
        case 'min':
          summary[field] = Math.min(...values);
          break;
        case 'max':
          summary[field] = Math.max(...values);
          break;
      }

      // Add percentage if requested
      if (showPercentage && func === 'count') {
        const total = data.length;
        summary[`${field}_percentage`] = ((summary[field] / total) * 100).toFixed(1) + '%';
      }
    });

    // Add the summary row
    const summaryRow = this.worksheet.createRow();
    summaryRow.setOutlineLevel(level + 1);

    // Add label cell
    const labelCell = summaryRow.createCell(`${label}:`);
    labelCell.setStyle(
      StyleBuilder.create()
        .bold(true)
        .italic(true)
        .backgroundColor('#f0f0f0')
        .build()
    );

    // Add summary values
    fields.forEach(field => {
      const value = summary[field];
      const cell = summaryRow.createCell(value);
      
      if (style) {
        cell.setStyle(style);
      } else {
        cell.setStyle(
          StyleBuilder.create()
            .bold(true)
            .numberFormat(functions[0] === 'count' ? '#,##0' : '#,##0.00')
            .build()
        );
      }
    });

    return this;
  }

  /**
   * Add a hierarchical section (tree structure)
   */
  addHierarchicalSection<T>(
    items: T[],
    childrenAccessor: (item: T) => T[],
    options?: {
      title?: string;
      level?: number;
      fields?: Array<keyof T>;
      collapsed?: boolean;
      showCount?: boolean;
    }
  ): this {
    const { title = 'Hierarchy', level = 0, fields, collapsed = true, showCount = true } = options || {};

    const buildHierarchy = (items: T[], currentLevel: number): SectionConfig[] => {
      return items.map(item => {
        const children = childrenAccessor(item);
        const hasChildren = children && children.length > 0;
        
        // Get display value for title
        const titleField = fields?.[0] || 'name';
        const titleValue = this.getNestedValue(item, titleField as string);
        const count = hasChildren ? ` (${children.length})` : '';

        return {
          title: `${titleValue}${showCount && hasChildren ? count : ''}`,
          level: currentLevel,
          collapsed,
          data: [item],
          fields: fields as string[],
          subSections: hasChildren ? buildHierarchy(children, currentLevel + 1) : undefined
        };
      });
    };

    const hierarchySections = buildHierarchy(items, level + 1);

    return this.addSection({
      title: title,
      level,
      collapsed,
      data: [],
      subSections: hierarchySections
    });
  }

  /**
   * Add a pivot-like section with multiple dimensions
   */
  addPivotSection(
    data: any[],
    dimensions: {
      rows: string[];
      columns: string[];
      values: Array<{
        field: string;
        aggregate: 'sum' | 'average' | 'count' | 'min' | 'max';
      }>;
    },
    options?: {
      level?: number;
      showGrandTotals?: boolean;
      showSubTotals?: boolean;
    }
  ): this {
    const { level = 0, showGrandTotals = true, showSubTotals = true } = options || {};
    
    // Group by row dimensions
    const rowGroups = this.groupMultiLevel(data, dimensions.rows);
    
    // Create sections for each row group
    Object.entries(rowGroups).forEach(([rowKey, rowItems]) => {
      const rowSection: SectionConfig = {
        title: rowKey,
        level: level + 1,
        collapsed: true,
        data: [],
        subSections: []
      };

      // Group by column dimensions within each row
      if (dimensions.columns.length > 0) {
        const colGroups = this.groupMultiLevel(rowItems, dimensions.columns);
        
        Object.entries(colGroups).forEach(([colKey, colItems]) => {
          // Calculate values for this cell
          const values: Record<string, any> = {};
          dimensions.values.forEach(v => {
            const nums = colItems.map(item => item[v.field]).filter((n: any) => !isNaN(n));
            switch (v.aggregate) {
              case 'sum':
                values[v.field] = nums.reduce((a: number, b: number) => a + b, 0);
                break;
              case 'average':
                values[v.field] = nums.length ? nums.reduce((a: number, b: number) => a + b, 0) / nums.length : 0;
                break;
              case 'count':
                values[v.field] = colItems.length;
                break;
              case 'min':
                values[v.field] = Math.min(...nums);
                break;
              case 'max':
                values[v.field] = Math.max(...nums);
                break;
            }
          });

          rowSection.subSections!.push({
            title: colKey,
            level: level + 2,
            data: [values],
            fields: dimensions.values.map(v => v.field)
          });
        });

        // Add row subtotal
        if (showSubTotals) {
          const subtotalValues: Record<string, any> = {};
          dimensions.values.forEach(v => {
            const nums = rowItems.map(item => item[v.field]).filter((n: any) => !isNaN(n));
            switch (v.aggregate) {
              case 'sum':
                subtotalValues[v.field] = nums.reduce((a: number, b: number) => a + b, 0);
                break;
              case 'average':
                subtotalValues[v.field] = nums.length ? nums.reduce((a: number, b: number) => a + b, 0) / nums.length : 0;
                break;
              case 'count':
                subtotalValues[v.field] = rowItems.length;
                break;
              case 'min':
                subtotalValues[v.field] = Math.min(...nums);
                break;
              case 'max':
                subtotalValues[v.field] = Math.max(...nums);
                break;
            }
          });

          rowSection.subSections!.push({
            title: `${rowKey} Subtotal`,
            level: level + 2,
            data: [subtotalValues],
            fields: dimensions.values.map(v => v.field)
          });
        }
      }

      this.addSection(rowSection);
    });

    // Add grand total
    if (showGrandTotals) {
      const grandTotalValues: Record<string, any> = {};
      dimensions.values.forEach(v => {
        const nums = data.map(item => item[v.field]).filter((n: any) => !isNaN(n));
        switch (v.aggregate) {
          case 'sum':
            grandTotalValues[v.field] = nums.reduce((a: number, b: number) => a + b, 0);
            break;
          case 'average':
            grandTotalValues[v.field] = nums.length ? nums.reduce((a: number, b: number) => a + b, 0) / nums.length : 0;
            break;
          case 'count':
            grandTotalValues[v.field] = data.length;
            break;
          case 'min':
            grandTotalValues[v.field] = Math.min(...nums);
            break;
          case 'max':
            grandTotalValues[v.field] = Math.max(...nums);
            break;
        }
      });

      this.addSection({
        title: 'Grand Total',
        level,
        data: [grandTotalValues],
        fields: dimensions.values.map(v => v.field)
      });
    }

    return this;
  }

  /**
   * Add a timeline section (grouped by date periods)
   */
  addTimelineSection(
    data: any[],
    dateField: string,
    period: 'day' | 'week' | 'month' | 'quarter' | 'year',
    options?: {
      fields?: string[];
      level?: number;
      showTrends?: boolean;
      format?: string;
    }
  ): this {
    const { fields, level = 0, showTrends = true, format } = options || {};

    // Group by date period
    const grouped = this.groupByDate(data, dateField, period, format);

    // Create sections for each period
    Object.entries(grouped).forEach(([periodKey, periodData]) => {
      const section: SectionConfig = {
        title: periodKey,
        level: level + 1,
        collapsed: true,
        data: periodData,
        fields
      };

      // Add trend indicators if requested
      if (showTrends && periodData.length > 0) {
        const prevPeriod = this.getPreviousPeriod(periodKey, grouped);
        if (prevPeriod) {
          const trend = this.calculateTrend(periodData, prevPeriod, fields?.[0] || 'value');
          section.summary = {
            fields: [fields?.[0] || 'value'],
            function: 'sum',
            label: `Trend: ${trend > 0 ? '↑' : '↓'} ${Math.abs(trend).toFixed(1)}%`
          };
        }
      }

      this.addSection(section);
    });

    return this;
  }

  /**
   * Add a filtered section
   */
  addFilteredSection(
    config: SectionConfig,
    filters: ExportFilters
  ): this {
    const filteredData = this.applyFilters(config.data, filters);
    
    return this.addSection({
      ...config,
      data: filteredData,
      title: `${config.title} (Filtered)`
    });
  }

  /**
   * Add conditional formatting to the current section
   */
  addConditionalFormat(rule: ConditionalFormatRule): this {
    this.conditionalFormats.push(rule);
    return this;
  }

  /**
   * Add a custom formatter for a field
   */
  addFormatter(field: string, formatter: (value: any) => any): this {
    this.formatters.set(field, formatter);
    return this;
  }

  /**
   * Register a reusable style
   */
  registerStyle(name: string, style: CellStyle): this {
    this.styles.set(name, style);
    return this;
  }

  /**
   * Apply a registered style
   */
  applyStyle(styleName: string): this {
    const style = this.styles.get(styleName);
    if (style) {
      const lastRow = this.worksheet.getRows().slice(-1)[0];
      if (lastRow) {
        lastRow.getCells().forEach(cell => cell.setStyle(style));
      }
    }
    return this;
  }

  /**
   * Get the current row count
   */
  getCurrentRow(): number {
    return this.currentRow;
  }

  /**
   * Reset the builder
   */
  reset(): this {
    this.currentRow = 0;
    this.sections.clear();
    return this;
  }

  // Private helper methods

  private groupData<T>(data: T[], groupBy: keyof T | ((item: T) => string)): Record<string, T[]> {
    return data.reduce((groups: Record<string, T[]>, item) => {
      const key = typeof groupBy === 'function' 
        ? groupBy(item)
        : String(item[groupBy]);
      
      if (!groups[key]) {
        groups[key] = [];
      }
      groups[key].push(item);
      return groups;
    }, {});
  }

  private groupMultiLevel(data: any[], dimensions: string[]): Record<string, any[]> {
    return data.reduce((groups: Record<string, any[]>, item) => {
      const key = dimensions.map(dim => this.getNestedValue(item, dim)).join(' › ');
      if (!groups[key]) {
        groups[key] = [];
      }
      groups[key].push(item);
      return groups;
    }, {});
  }

  private groupByDate(
    data: any[],
    dateField: string,
    period: 'day' | 'week' | 'month' | 'quarter' | 'year',
    format?: string
  ): Record<string, any[]> {
    return data.reduce((groups: Record<string, any[]>, item) => {
      const date = new Date(this.getNestedValue(item, dateField));
      let key: string;

      switch (period) {
        case 'day':
          key = format || date.toISOString().split('T')[0];
          break;
        case 'week': {
          const week = this.getWeekNumber(date);
          key = format || `${date.getFullYear()}-W${week}`;
          break;
        }
        case 'month':
          key = format || `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
          break;
        case 'quarter': {
          const quarter = Math.floor(date.getMonth() / 3) + 1;
          key = format || `${date.getFullYear()}-Q${quarter}`;
          break;
        }
        case 'year':
          key = format || String(date.getFullYear());
          break;
      }

      if (!groups[key]) {
        groups[key] = [];
      }
      groups[key].push(item);
      return groups;
    }, {});
  }

  private getWeekNumber(date: Date): number {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil(((d.getTime() - yearStart.getTime()) / 86400000 + 1) / 7);
  }

  private getPreviousPeriod(
    currentKey: string,
    groups: Record<string, any[]>
  ): any[] | null {
    const keys = Object.keys(groups).sort();
    const currentIndex = keys.indexOf(currentKey);
    if (currentIndex > 0) {
      return groups[keys[currentIndex - 1]];
    }
    return null;
  }

  private calculateTrend(current: any[], previous: any[], field: string): number {
    const currentSum = current.reduce((sum, item) => sum + (item[field] || 0), 0);
    const previousSum = previous.reduce((sum, item) => sum + (item[field] || 0), 0);
    
    if (previousSum === 0) return 0;
    return ((currentSum - previousSum) / previousSum) * 100;
  }

  private applyFilters(data: any[], filters: ExportFilters): any[] {
    return data.filter(item => {
      // Date range filter
      if (filters.dateRange) {
        const dateValue = new Date(this.getNestedValue(item, filters.dateRange.field || 'date'));
        const start = new Date(filters.dateRange.start);
        const end = new Date(filters.dateRange.end);
        
        if (dateValue < start || dateValue > end) {
          return false;
        }
      }

      // Field filters
      if (filters.filters) {
        for (const [field, values] of Object.entries(filters.filters)) {
          const itemValue = this.getNestedValue(item, field);
          const arrValues = values as any[];
          if (arrValues.length > 0 && !arrValues.includes(itemValue)) {
            return false;
          }
        }
      }

      // Search filter
      if (filters.search) {
        const searchTerm = filters.search.toLowerCase();
        const matches = Object.values(item).some(val => 
          String(val).toLowerCase().includes(searchTerm)
        );
        if (!matches) return false;
      }

      // Custom filter
      if (filters.customFilter && !filters.customFilter(item)) {
        return false;
      }

      return true;
    }).slice(filters.offset || 0, (filters.offset || 0) + (filters.limit || Infinity));
  }

  private getNestedValue(obj: any, path: string): any {
    return path.split('.').reduce((current, key) => {
      return current && current[key] !== undefined ? current[key] : undefined;
    }, obj);
  }
}

/**
 * Internal Section class for building collapsible sections
 */
class Section {
  private worksheet: Worksheet;
  private config: SectionConfig;
  private startRow: number;
  private subSections: Section[] = [];

  constructor(worksheet: Worksheet, config: SectionConfig, startRow: number) {
    this.worksheet = worksheet;
    this.config = config;
    this.startRow = startRow;
  }

  /**
   * Build the section in the worksheet
   */
  build(): number {
    let currentRow = this.startRow;

    // Add section header
    if (this.config.title) {
      const headerRow = this.worksheet.createRow();
      headerRow.setOutlineLevel(this.config.level);
      
      const headerCell = headerRow.createCell(this.config.title);
      
      // Apply header style
      const headerStyle = this.config.headerStyle || 
        StyleBuilder.create()
          .bold(true)
          .fontSize(12)
          .backgroundColor(this.getHeaderColor(this.config.level))
          .borderAll('thin')
          .build();
      
      headerCell.setStyle(headerStyle);
      
      // Add collapse/expand indicator if section has data or subsections
      if (this.hasContent()) {
        const indicatorCell = headerRow.createCell(this.config.collapsed ? '▶' : '▼');
        indicatorCell.setStyle(
          StyleBuilder.create()
            .bold(true)
            .color('#666666')
            .build()
        );
      }
      
      currentRow++;
    }

    // Add section data if not collapsed
    if (!this.config.collapsed && this.config.data.length > 0) {
      // Add headers if fields are specified
      if (this.config.fields && this.config.fields.length > 0) {
        const headerRow = this.worksheet.createRow();
        headerRow.setOutlineLevel(this.config.level + 1);
        
        this.config.fields.forEach((field: string | number) => {
          const label = this.config.fieldLabels?.[field] || field;
          const cell = headerRow.createCell(label);
          cell.setStyle(
            StyleBuilder.create()
              .bold(true)
              .backgroundColor('#e6e6e6')
              .borderAll('thin')
              .build()
          );
        });
        
        currentRow++;
      }

      // Add data rows
      this.config.data.forEach((item: Record<string, unknown>) => {
        const dataRow = this.worksheet.createRow();
        dataRow.setOutlineLevel(this.config.level + 1);
        
        const fields = this.config.fields || Object.keys(item);
        fields.forEach((field: string) => {
          const value = this.getNestedValue(item, field);
          const cell = dataRow.createCell(value);
          
          // Apply conditional formatting if any
          if (this.config.conditionalStyles) {
            this.applyConditionalStyles(cell, item, field);
          }
          
          // Apply field-specific style
          if (this.config.style) {
            cell.setStyle(this.config.style);
          }
        });
        
        currentRow++;
      });

      // Add summary row if configured
      if (this.config.summary) {
        this.addSummaryRow();
        currentRow++;
      }
    }

    // Build sub-sections
    if (this.config.subSections && !this.config.collapsed) {
      this.config.subSections.forEach((subConfig: SectionConfig) => {
        const subSection = new Section(this.worksheet, subConfig, currentRow);
        this.subSections.push(subSection);
        currentRow = subSection.build();
      });
    }

    return currentRow;
  }

  /**
   * Add a sub-section
   */
  addSubSection(config: SectionConfig): void {
    if (!this.config.subSections) {
      this.config.subSections = [];
    }
    this.config.subSections.push(config);
  }

  /**
   * Check if section has any content
   */
  private hasContent(): boolean {
    return this.config.data.length > 0 || 
           (this.config.subSections && this.config.subSections.length > 0);
  }

  /**
   * Get header color based on outline level
   */
  private getHeaderColor(level: number): string {
    const colors = [
      '#e6f2ff', // Level 0 - Light blue
      '#f0f0f0', // Level 1 - Light gray
      '#f9f9f9', // Level 2 - Lighter gray
      '#ffffff', // Level 3+ - White
    ];
    return colors[Math.min(level, colors.length - 1)];
  }

  /**
   * Add summary row (totals, averages, etc.)
   */
  private addSummaryRow(): void {
    if (!this.config.summary) return;

    const summaryRow = this.worksheet.createRow();
    summaryRow.setOutlineLevel(this.config.level + 2);

    // Add label cell
    const labelCell = summaryRow.createCell(this.config.summary.label || 'Total');
    labelCell.setStyle(
      StyleBuilder.create()
        .bold(true)
        .italic(true)
        .backgroundColor('#f5f5f5')
        .build()
    );

    // Calculate and add summary values
    const fields = this.config.fields || [];
    fields.forEach((field: string) => {
      if (this.config.summary!.fields.includes(field)) {
        const values = this.config.data
          .map((item: any) => this.getNestedValue(item, field))
          .filter((v: number | null) => v != null && !isNaN(v));
        
        let result: number;
        switch (this.config.summary!.function) {
          case 'sum':
            result = values.reduce((a: any, b: any) => a + b, 0);
            break;
          case 'average':
            result = values.length ? values.reduce((a: any, b: any) => a + b, 0) / values.length : 0;
            break;
          case 'count':
            result = values.length;
            break;
          case 'min':
            result = Math.min(...values);
            break;
          case 'max':
            result = Math.max(...values);
            break;
          default:
            result = 0;
        }

        const cell = summaryRow.createCell(result);
        cell.setStyle(
          StyleBuilder.create()
            .bold(true)
            .numberFormat(this.config.summary!.function === 'count' ? '#,##0' : '#,##0.00')
            .build()
        );
      } else {
        summaryRow.createCell('');
      }
    });
  }

  /**
   * Apply conditional formatting to a cell
   */
  private applyConditionalStyles(cell: Cell, item: any, field: string): void {
    if (!this.config.conditionalStyles) return;

    for (const rule of this.config.conditionalStyles) {
      if (rule.field === field && rule.condition(item)) {
        cell.setStyle(rule.style);
        break;
      }
    }
  }

  /**
   * Get nested value from object
   */
  private getNestedValue(obj: any, path: string): any {
    return path.split('.').reduce((current, key) => {
      return current && current[key] !== undefined ? current[key] : undefined;
    }, obj);
  }
}

/**
 * Factory function to create a new SectionBuilder
 */
export function createSectionBuilder(worksheet: Worksheet): SectionBuilder {
  return new SectionBuilder(worksheet);
}

/**
 * Pre-built section templates
 */
export const SectionTemplates = {
  /**
   * Create a financial summary section
   */
  financialSummary(data: any[], options?: any): SectionConfig {
    return {
      title: options?.title || 'Financial Summary',
      level: 0,
      data,
      fields: ['date', 'description', 'debit', 'credit', 'balance'],
      fieldLabels: {
        date: 'Date',
        description: 'Description',
        debit: 'Debit ($)',
        credit: 'Credit ($)',
        balance: 'Balance ($)'
      },
      summary: {
        fields: ['debit', 'credit', 'balance'],
        function: 'sum',
        label: 'Totals'
      },
      conditionalStyles: [
        {
          field: 'balance',
          condition: (item: { balance: number; }) => item.balance < 0,
          style: StyleBuilder.create().color('#ff0000').bold(true).build()
        },
        {
          field: 'balance',
          condition: (item: { balance: number; }) => item.balance > 10000,
          style: StyleBuilder.create().color('#008000').bold(true).build()
        }
      ]
    };
  },

  /**
   * Create an inventory section
   */
  inventorySection(parts: any[], instances?: any[]): SectionConfig {
    return {
      title: 'Inventory Status',
      level: 0,
      data: parts,
      fields: ['part_number', 'part_name', 'category', 'current_stock', 'minimum_stock', 'status'],
      fieldLabels: {
        part_number: 'Part #',
        part_name: 'Part Name',
        category: 'Category',
        current_stock: 'Stock',
        minimum_stock: 'Min Stock',
        status: 'Status'
      },
      conditionalStyles: [
        {
          field: 'current_stock',
          condition: (item: { current_stock: number; }) => item.current_stock === 0,
          style: StyleBuilder.create().backgroundColor('#ffebee').color('#c62828').bold(true).build()
        },
        {
          field: 'current_stock',
          condition: (item: { current_stock: number; minimum_stock: number; }) => item.current_stock < item.minimum_stock && item.current_stock > 0,
          style: StyleBuilder.create().backgroundColor('#fff3e0').color('#ef6c00').build()
        }
      ],
      subSections: instances ? [
        {
          title: 'Part Instances',
          level: 1,
          data: instances,
          fields: ['instance_id', 'serial_number', 'status', 'location', 'installed_in_vehicle'],
          groupBy: 'part_number'
        }
      ] : undefined
    };
  },

  /**
   * Create a project timeline section
   */
  projectTimeline(tasks: any[]): SectionConfig {
    const now = new Date();
    
    return {
      title: 'Project Timeline',
      level: 0,
      data: tasks,
      fields: ['task', 'assignee', 'start_date', 'end_date', 'status', 'completion'],
      fieldLabels: {
        task: 'Task',
        assignee: 'Assigned To',
        start_date: 'Start Date',
        end_date: 'End Date',
        status: 'Status',
        completion: 'Completion %'
      },
      conditionalStyles: [
        {
          field: 'end_date',
          condition: (item: { end_date: string | number | Date; status: string; }) => new Date(item.end_date) < now && item.status !== 'Completed',
          style: StyleBuilder.create().backgroundColor('#ffebee').color('#c62828').bold(true).build()
        },
        {
          field: 'completion',
          condition: (item: { completion: number; }) => item.completion === 100,
          style: StyleBuilder.create().backgroundColor('#e8f5e8').color('#2e7d32').build()
        },
        {
          field: 'completion',
          condition: (item: { completion: number; }) => item.completion < 50,
          style: StyleBuilder.create().color('#ff6d00').build()
        }
      ]
    };
  }
};