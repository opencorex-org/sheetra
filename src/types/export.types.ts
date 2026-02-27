import { Border, Fill } from "./cell.types";

/**
 * Export options for exporting sheets, workbooks, etc.
 */
export interface ExportOptions {
  /** Output file name (with extension) */
  filename?: string;
  /** File format: 'xlsx', 'csv', 'json', etc. */
  format?: 'xlsx' | 'csv' | 'json';
  /** Whether to include styles in export */
  includeStyles?: boolean;
  /** Whether to include hidden rows/columns */
  includeHidden?: boolean;
  /** Password for protected export (if supported) */
  password?: string;
  /** Sheet name (for single-sheet export) */
  sheetName?: string;
  /** Locale for formatting */
  locale?: string;
  /** Custom options for writers */
  [key: string]: any;
}

/**
 * Chart series data
 */
export interface ChartSeries {
  /** Series name */
  name?: string;
  /** Categories (X-axis) */
  categories?: any[];
  /** Values (Y-axis) */
  values: any[];
  /** Series color */
  color?: string;
  /** Series style */
  style?: number;
  /** Marker style (for line charts) */
  marker?: {
    symbol?: 'circle' | 'diamond' | 'square' | 'triangle' | 'none';
    size?: number;
    fill?: string;
    stroke?: string;
  };
  /** Trendline */
  trendline?: {
    type: 'linear' | 'exponential' | 'logarithmic' | 'polynomial' | 'movingAverage';
    order?: number;
    period?: number;
    forward?: number;
    backward?: number;
    intercept?: number;
    displayEquation?: boolean;
    displayRSquared?: boolean;
    name?: string;
  };
  /** Error bars */
  errorBars?: {
    type: 'fixed' | 'percentage' | 'standardDeviation' | 'standardError' | 'custom';
    value?: number;
    plusValues?: any[];
    minusValues?: any[];
    direction?: 'both' | 'plus' | 'minus';
    endStyle?: 'cap' | 'noCap';
  };
}

/**
 * Chart axis configuration
 */
export interface ChartAxis {
  /** Axis title */
  title?: string;
  /** Minimum value */
  min?: number;
  /** Maximum value */
  max?: number;
  /** Major unit */
  majorUnit?: number;
  /** Minor unit */
  minorUnit?: number;
  /** Reverse order */
  reverse?: boolean;
  /** Logarithmic scale */
  logarithmic?: boolean;
  /** Base for logarithmic scale */
  logBase?: number;
  /** Format code */
  format?: string;
  /** Axis position */
  position?: 'bottom' | 'left' | 'right' | 'top';
  /** Crossing point */
  crosses?: 'auto' | 'min' | 'max' | number;
  /** Gridlines */
  gridlines?: {
    major?: boolean;
    minor?: boolean;
    color?: string;
  };
  /** Tick marks */
  tickMarks?: {
    major?: 'none' | 'inside' | 'outside' | 'cross';
    minor?: 'none' | 'inside' | 'outside' | 'cross';
  };
  /** Font */
  font?: {
    name?: string;
    size?: number;
    bold?: boolean;
    color?: string;
  };
}

/**
 * Complete chart configuration
 */
export interface ChartConfig {
  /** Chart title */
  title?: string;
  /** Chart type */
  type: 'area' | 'bar' | 'column' | 'line' | 'pie' | 'doughnut' | 'radar' | 'scatter' | 'stock' | 'surface' | 'bubble' | 'radarArea' | 'radarLine' | 'radarFilled' | 'barStacked' | 'columnStacked' | 'barPercent' | 'columnPercent' | 'areaStacked' | 'areaPercent' | 'lineStacked' | 'linePercent' | 'pie3D' | 'bar3D' | 'column3D' | 'line3D' | 'area3D' | 'surface3D';
  /** Chart sub-type */
  subType?: 'clustered' | 'stacked' | 'percent' | '3D' | '3DClustered' | '3DStacked' | '3DPercent';
  /** Data range */
  dataRange?: string;
  /** Series data */
  series?: ChartSeries[];
  /** X-axis configuration */
  xAxis?: ChartAxis;
  /** Y-axis configuration */
  yAxis?: ChartAxis;
  /** Secondary X-axis */
  xAxis2?: ChartAxis;
  /** Secondary Y-axis */
  yAxis2?: ChartAxis;
  /** Legend */
  legend?: {
    position?: 'top' | 'bottom' | 'left' | 'right' | 'corner' | 'none';
    layout?: 'stack' | 'overlay';
    showSeriesName?: boolean;
    font?: {
      name?: string;
      size?: number;
      bold?: boolean;
      color?: string;
    };
  };
  /** Data labels */
  dataLabels?: {
    show?: boolean;
    position?: 'center' | 'insideEnd' | 'insideBase' | 'outsideEnd' | 'bestFit';
    format?: string;
    font?: {
      name?: string;
      size?: number;
      bold?: boolean;
      color?: string;
    };
    separator?: string;
    showSeriesName?: boolean;
    showCategoryName?: boolean;
    showValue?: boolean;
    showPercentage?: boolean;
    showLeaderLines?: boolean;
  };
  /** Chart size in pixels */
  size?: {
    width: number;
    height: number;
  };
  /** Chart position in pixels (from top-left of sheet) */
  position?: {
    x: number;
    y: number;
  };
  /** Chart style (1-48) */
  style?: number;
  /** Chart template */
  template?: string;
  /** Gap width (for bar/column) */
  gapWidth?: number;
  /** Overlap (for bar/column) */
  overlap?: number;
  /** Vary colors by point */
  varyColors?: boolean;
  /** Smooth lines */
  smooth?: boolean;
  /** Show data table */
  dataTable?: {
    show?: boolean;
    showLegendKeys?: boolean;
    border?: boolean;
    font?: {
      name?: string;
      size?: number;
      bold?: boolean;
      color?: string;
    };
  };
  /** 3D options */
  threeD?: {
    rotation?: number;
    perspective?: number;
    height?: number;
    depth?: number;
    rightAngleAxes?: boolean;
    lighting?: boolean;
  };
  /** Plot area */
  plotArea?: {
    border?: Border;
    fill?: Fill;
    transparency?: number;
  };
  /** Chart area */
  chartArea?: {
    border?: Border;
    fill?: Fill;
    transparency?: number;
  };
}

/**
 * Drawing object (image, shape, etc.)
 */
export interface Drawing {
  /** Drawing type */
  type: 'image' | 'shape' | 'chart' | 'smartArt';
  /** Drawing position */
  from: {
    col: number;
    row: number;
    colOffset?: number; // in pixels
    rowOffset?: number; // in pixels
  };
  /** Drawing size */
  to?: {
    col: number;
    row: number;
    colOffset?: number;
    rowOffset?: number;
  };
  /** Image data (for images) */
  image?: {
    data: string | ArrayBuffer; // base64 or binary
    format: 'png' | 'jpg' | 'gif' | 'bmp' | 'svg';
    name?: string;
    description?: string;
  };
  /** Shape data (for shapes) */
  shape?: {
    type: 'rectangle' | 'ellipse' | 'triangle' | 'line' | 'arrow' | 'callout' | 'flowchart' | 'star' | 'banner';
    text?: string;
    fill?: Fill;
    border?: Border;
    rotation?: number;
    flipHorizontal?: boolean;
    flipVertical?: boolean;
  };
  /** Chart reference */
  chart?: ChartConfig;
  /** Alternative text */
  altText?: string;
  /** Hyperlink */
  hyperlink?: string;
  /** Lock aspect ratio */
  lockAspect?: boolean;
  /** Lock position */
  lockPosition?: boolean;
  /** Print object */
  print?: boolean;
  /** Hidden */
  hidden?: boolean;
}

/**
 * Section configuration for export (e.g., for splitting sheets into sections)
 */
export interface SectionConfig {
  /** Section title */
  title?: string;
  /** Section description */
  description?: string;
  /** Range of rows included in this section (e.g., "A1:D10") */
  range?: string;
  /** Whether the section is collapsible */
  collapsible?: boolean;
  /** Whether the section is initially collapsed */
  collapsed?: boolean;
  /** Custom styles for the section */
  style?: {
    backgroundColor?: string;
    fontColor?: string;
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
    border?: Border;
    fill?: Fill;
  };
  /** Additional custom options */
  [key: string]: any;
}

/**
 * Export filters
 */
export interface ExportFilters {
  [key: string]: any;
}