# Sheetra Examples

This document provides practical examples of how to use Sheetra to export data to various formats, including Excel, CSV, and JSON. These examples demonstrate the flexibility and power of the Sheetra API.

---

## Basic Example: Exporting to Excel

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

## Advanced Example: Exporting Inventory Data

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

## Exporting to CSV

```typescript
import { ExportBuilder } from 'sheetra';

const data = [
    ['Name', 'Age'],
    ['John', 30],
    ['Jane', 25]
];

ExportBuilder.create('Users')
    .addDataRows(data)
    .download({ filename: 'users.csv', format: 'csv' });
```

---

## Exporting to JSON

```typescript
import { ExportBuilder } from 'sheetra';

const data = [
    { name: 'John', age: 30 },
    { name: 'Jane', age: 25 }
];

ExportBuilder.create('Users')
    .addDataRows(data)
    .download({ filename: 'users.json', format: 'json' });
```

---

For more examples, visit the [Getting Started Guide](GETTING_STARTED.md).