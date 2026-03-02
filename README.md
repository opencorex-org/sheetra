# Sheetra

**Zero-Dependency Excel/CSV/JSON Export Library**

[![npm version](https://img.shields.io/npm/v/sheetra.svg)](https://www.npmjs.com/package/sheetra)
[![License](https://img.shields.io/github/license/opencorex-org/sheetra)](LICENSE)
[![Build Status](https://github.com/opencorex-org/sheetra/actions/workflows/build.yml/badge.svg)](https://github.com/opencorex-org/sheetra/actions)
[![npm downloads](https://img.shields.io/npm/dm/sheetra.svg)](https://www.npmjs.com/package/sheetra)
[![GitHub stars](https://img.shields.io/github/stars/opencorex-org/sheetra?style=social)](https://github.com/opencorex-org/sheetra)
[![GitHub issues](https://img.shields.io/github/issues/opencorex-org/sheetra)](https://github.com/opencorex-org/sheetra/issues)

Sheetra is a powerful, zero-dependency library for exporting data to Excel (XLSX), CSV, and JSON formats. It provides a clean, fluent API for creating complex spreadsheets with styles, outlines, and sections.

---


## Example Usage

```ts
import { ExportBuilder, StyleBuilder } from 'sheetra';

// --- Comprehensive Example: All Features ---
const users = [
  { name: 'John Doe', age: 30, email: 'john@example.com', department: 'Engineering' },
  { name: 'Jane Smith', age: 25, email: 'jane@example.com', department: 'Marketing' },
];
const parts = [
  { part_number: 'P001', part_name: 'Widget', current_stock: 100, price: 29.99 },
  { part_number: 'P002', part_name: 'Gadget', current_stock: 50, price: 49.99 },
];
const salesData = [
  { product: 'Widget', jan: 1500, feb: 1800, mar: 2100 },
  { product: 'Gadget', jan: 900, feb: 1100, mar: 1300 },
];

// Styled header
const headerStyle = StyleBuilder.create()
  .bold()
  .backgroundColor('#4F81BD')
  .color('#FFFFFF')
  .align('center')
  .build();

ExportBuilder.create('All Features Demo')
  // Set column widths
  .setColumnWidths([150, 100, 200, 120])
  // Add styled header row
  .addHeaderRow(['Name', 'Age', 'Email', 'Department'], headerStyle)
  // Add data rows
  .addDataRows(users.map(u => [u.name, u.age, u.email, u.department]))
  // Add section
  .addSection({ name: 'Parts', title: 'Parts Inventory' })
  .addHeaderRow(['Part Number', 'Part Name', 'Stock', 'Price'])
  .addDataRows(parts.map(p => [p.part_number, p.part_name, p.current_stock, p.price]))
  // Merge cells for a title
  .addDataRows([['Monthly Sales Report', '', '', '']])
  .mergeCells(7, 0, 7, 3)
  .setAlignment(7, 0, 'center')
  // Alignment demo
  .addHeaderRow(['Left', 'Center', 'Right'])
  .setAlignment(8, 0, 'left')
  .setAlignment(8, 1, 'center')
  .setAlignment(8, 2, 'right')
  .addDataRows([
    ['Left Text', 'Center Text', 'Right Text'],
    ['Value 1', 'Value 2', 'Value 3'],
  ])
  .setRangeAlignment(9, 0, 10, 0, 'left')
  .setRangeAlignment(9, 1, 10, 1, 'center')
  .setRangeAlignment(9, 2, 10, 2, 'right')
  // Auto-size columns
  .autoSizeColumns()
  // Download as XLSX
  .download({ filename: 'all-features-demo.xlsx', format: 'xlsx' });
```

This example demonstrates:
- Column widths and auto-sizing
- Merged cells for titles
- Alignment (left, center, right, range)
- Sections and headers
- Styled headers with StyleBuilder
- Data rows and multiple tables
- XLSX export

See the [sheetra-demo](./sheetra-demo/src/App.tsx) for a full-featured React demo with all features in action.
  .addDataRows(users.map(u => [u.name, u.age, u.email, u.department]))
  .download({ filename: 'styled-report.xlsx', format: 'xlsx' });
```

## Troubleshooting

**Upgrading from previous versions?**

If you previously encountered errors with styled exports (e.g., `SyntaxError: Expected ':' after property name in JSON`), upgrade to the latest version. Style and alignment handling is now robust and all edge cases are supported.


---

## Installation

Install Sheetra using your favorite package manager:

```bash
npm install sheetra
# or
yarn add sheetra
# or
pnpm add sheetra
```

---

## Usage

### Basic Example

```typescript
import { ExportBuilder } from 'sheetra';

const builder = ExportBuilder.create('Users')
  .addHeaderRow(['Name', 'Age'])
  .addDataRows([
    ['John', 30],
    ['Jane', 25]
  ]);

builder.download({ filename: 'users.xlsx', format: 'xlsx' });
```

### Advanced Example

```typescript
import { ExportBuilder } from 'sheetra';

const parts = [
  { part_number: 'P001', part_name: 'Widget', current_stock: 100 },
  { part_number: 'P002', part_name: 'Gadget', current_stock: 50 },
];

const instances = [
  { serial_number: 'S001', part_number: 'P001', status: 'Active', location: 'A1' },
  { serial_number: 'S002', part_number: 'P001', status: 'Inactive', location: 'A2' },
  { serial_number: 'S003', part_number: 'P002', status: 'Active', location: 'B1' },
];

ExportBuilder.create('Inventory')
  .addSection({ name: 'Parts' })
  .addHeaderRow(['Part Number', 'Part Name', 'Current Stock'])
  .addDataRows(parts.map(p => [p.part_number, p.part_name, p.current_stock]))
  .addSection({ name: 'Instances' })
  .addHeaderRow(['Serial Number', 'Part Number', 'Status', 'Location'])
  .addDataRows(instances.map(i => [i.serial_number, i.part_number, i.status, i.location]))
  .download({ filename: 'inventory.xlsx', format: 'xlsx' });
```

---

## Documentation

For detailed documentation, visit the [API Reference](docs/API.md) and [Getting Started Guide](docs/GETTING_STARTED.md).

---

## Contributing

We welcome contributions! Please read the [Contributing Guide](docs/CONTRIBUTING.md) to learn how you can help improve Sheetra.

---

## License

This project is licensed under the [Apache 2.0 License](LICENSE).