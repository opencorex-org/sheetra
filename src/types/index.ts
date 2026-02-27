// Main entry point for all types
export * from './workbook.types';
export * from './cell.types';

// Re-export common types for convenience
import { CellStyle, Border, BorderStyle } from './cell.types';
import { WorkbookData, WorksheetData, RowData, ColumnData } from './workbook.types';
import { ExportOptions, SectionConfig, Drawing, ExportFilters } from './export.types';

export type {
    // Cell types
    CellStyle,
    Border,
    BorderStyle,

    // Workbook types
    WorkbookData,
    WorksheetData,
    RowData,
    ColumnData,

    // Export types
    ExportOptions,
    SectionConfig,
    Drawing,
    ExportFilters
};

// Utility types
export type DeepPartial<T> = {
    [P in keyof T]?: T[P] extends object ? DeepPartial<T[P]> : T[P];
};

export type Nullable<T> = T | null | undefined;

export type Record = {
    [key: string]: unknown;
};