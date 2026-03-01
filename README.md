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

## Features

- **Zero Dependencies** – Pure TypeScript/JavaScript implementation
- **Multiple Formats** – Export to XLSX, CSV, JSON
- **Styling Support** – Bold, colors, borders, alignment
- **Outlines & Groups** – Create collapsible sections
- **Fluent API** – Chain methods for easy construction
- **TypeScript** – Full type support
- **Browser & Node** – Works in both environments

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