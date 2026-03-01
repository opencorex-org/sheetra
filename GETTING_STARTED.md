# Getting Started with Sheetra

Welcome to Sheetra! This guide will help you get started with using Sheetra to export data to Excel, CSV, and JSON formats.

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

## Basic Usage

### Exporting to Excel

```typescript
import { ExportBuilder } from 'sheetra';

const data = [
    ['Name', 'Age'],
    ['John', 30],
    ['Jane', 25]
];

ExportBuilder.create('Users')
    .addDataRows(data)
    .download({ filename: 'users.xlsx', format: 'xlsx' });
```

---

## Advanced Usage

### Exporting Inventory Data

```typescript
import { ExportBuilder } from 'sheetra';

const parts = [
    { part_number: 'P001', part_name: 'Widget', current_stock: 100 },
    { part_number: 'P002', part_name: 'Gadget', current_stock: 50 },
];

ExportBuilder.create('Inventory')
    .addHeaderRow(['Part Number', 'Part Name', 'Current Stock'])
    .addDataRows(parts.map(p => [p.part_number, p.part_name, p.current_stock]))
    .download({ filename: 'inventory.xlsx', format: 'xlsx' });
```

---

## Additional Resources

- [Contributing Guide](CONTRIBUTING.md)
- [Examples](EXAMPLES.md)
- [CSV Documentation](csv-guide.html)

---

Thank you for using Sheetra! If you have any questions, feel free to open an issue on [GitHub](https://github.com/opencorex-org/sheetra/issues).