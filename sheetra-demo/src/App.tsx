import React from 'react';
import { ExportBuilder, StyleBuilder } from 'sheetra';
import './App.css';

// Sample data
const users = [
  { name: 'John Doe', age: 30, email: 'john@example.com', department: 'Engineering' },
  { name: 'Jane Smith', age: 25, email: 'jane@example.com', department: 'Marketing' },
  { name: 'Bob Johnson', age: 35, email: 'bob@example.com', department: 'Engineering' },
  { name: 'Alice Brown', age: 28, email: 'alice@example.com', department: 'Sales' },
];

const parts = [
  { part_number: 'P001', part_name: 'Widget', current_stock: 100, price: 29.99 },
  { part_number: 'P002', part_name: 'Gadget', current_stock: 50, price: 49.99 },
  { part_number: 'P003', part_name: 'Component', current_stock: 200, price: 9.99 },
];

const instances = [
  { serial_number: 'S001', part_number: 'P001', status: 'Active', location: 'A1' },
  { serial_number: 'S002', part_number: 'P001', status: 'Inactive', location: 'A2' },
  { serial_number: 'S003', part_number: 'P002', status: 'Active', location: 'B1' },
  { serial_number: 'S004', part_number: 'P003', status: 'Active', location: 'C1' },
];

const salesData = [
  { product: 'Widget', jan: 1500, feb: 1800, mar: 2100 },
  { product: 'Gadget', jan: 900, feb: 1100, mar: 1300 },
  { product: 'Accessory', jan: 500, feb: 600, mar: 700 },
];

function App() {
  // 1. Simple CSV Export
  const handleSimpleExportCsv = () => {
    ExportBuilder.create('Users')
      .addHeaderRow(['Name', 'Age', 'Email', 'Department'])
      .addDataRows(users.map(u => [u.name, u.age, u.email, u.department]))
      .download({ filename: 'users.csv', format: 'csv' });
  };

  // 2. Simple XLSX Export with Column Widths
  const handleSimpleExportXlsx = () => {
    ExportBuilder.create('Users')
      .setColumnWidths([150, 80, 200, 120])
      .addHeaderRow(['Name', 'Age', 'Email', 'Department'])
      .addDataRows(users.map(u => [u.name, u.age, u.email, u.department]))
      .download({ filename: 'users.xlsx', format: 'xlsx' });
  };

  // 3. JSON Export
  const handleJsonExport = () => {
    ExportBuilder.create('Users')
      .addHeaderRow(['Name', 'Age', 'Email', 'Department'])
      .addDataRows(users.map(u => [u.name, u.age, u.email, u.department]))
      .download({ filename: 'users.json', format: 'json' });
  };

  // 4. Export with Sections
  const handleSectionExport = () => {
    ExportBuilder.create('Inventory')
      .setColumnWidths([120, 120, 100, 100])
      .addSection({ name: 'Parts', title: 'Parts Inventory' })
      .addHeaderRow(['Part Number', 'Part Name', 'Stock', 'Price'])
      .addDataRows(parts.map(p => [p.part_number, p.part_name, p.current_stock, p.price]))
      .addDataRows([['', '', '', '']]) // Empty row separator
      .addSection({ name: 'Instances', title: 'Part Instances' })
      .addHeaderRow(['Serial Number', 'Part Number', 'Status', 'Location'])
      .addDataRows(instances.map(i => [i.serial_number, i.part_number, i.status, i.location]))
      .download({ filename: 'inventory-sections.xlsx', format: 'xlsx' });
  };

  // 5. Merge Cells Example
  const handleMergeCellsExport = () => {
    ExportBuilder.create('Report')
      .setColumnWidths([200, 100, 100, 100])
      // Title row (merged across all columns)
      .addDataRows([['Monthly Sales Report', '', '', '']])
      .mergeCells(0, 0, 0, 3)
      .setAlignment(0, 0, 'center')
      // Subtitle row (merged)
      .addDataRows([['Q1 2026 Performance', '', '', '']])
      .mergeCells(1, 0, 1, 3)
      .setAlignment(1, 0, 'center')
      // Empty row
      .addDataRows([['', '', '', '']])
      // Header row
      .addHeaderRow(['Product', 'January', 'February', 'March'])
      .setRangeAlignment(3, 0, 3, 3, 'center')
      // Data rows
      .addDataRows(salesData.map(s => [s.product, s.jan, s.feb, s.mar]))
      .setRangeAlignment(4, 1, 6, 3, 'right')
      // Total row
      .addDataRows([['TOTAL', 2900, 3500, 4100]])
      .setAlignment(7, 0, 'right')
      .setRangeAlignment(7, 1, 7, 3, 'right')
      .download({ filename: 'sales-report-merged.xlsx', format: 'xlsx' });
  };

  // 6. Alignment Examples
  const handleAlignmentExport = () => {
    ExportBuilder.create('Alignment Demo')
      .setColumnWidths([150, 150, 150])
      // Title
      .addDataRows([['Alignment Demonstration', '', '']])
      .mergeCells(0, 0, 0, 2)
      .setAlignment(0, 0, 'center')
      // Headers
      .addHeaderRow(['Left Aligned', 'Center Aligned', 'Right Aligned'])
      .setAlignment(1, 0, 'left')
      .setAlignment(1, 1, 'center')
      .setAlignment(1, 2, 'right')
      // Data rows
      .addDataRows([
        ['Left Text', 'Center Text', 'Right Text'],
        ['Value 1', 'Value 2', 'Value 3'],
        ['123', '456', '789'],
      ])
      .setRangeAlignment(2, 0, 4, 0, 'left')
      .setRangeAlignment(2, 1, 4, 1, 'center')
      .setRangeAlignment(2, 2, 4, 2, 'right')
      .download({ filename: 'alignment-demo.xlsx', format: 'xlsx' });
  };

  // 7. Auto Size Columns
  const handleAutoSizeExport = () => {
    ExportBuilder.create('Auto Sized')
      .addHeaderRow(['Short', 'Medium Length Header', 'This is a very long header text'])
      .addDataRows([
        ['A', 'Medium text here', 'This is some longer content in the cell'],
        ['B', 'Another medium', 'Short'],
        ['C', 'Text', 'Medium length content here'],
      ])
      .autoSizeColumns()
      .download({ filename: 'auto-sized-columns.xlsx', format: 'xlsx' });
  };

  // 8. Complex Report with All Features
  const handleComplexExport = () => {
    ExportBuilder.create('Complete Report')
      .setColumnWidths([180, 120, 100, 100, 120])
      // Main Title
      .addDataRows([['COMPREHENSIVE INVENTORY REPORT', '', '', '', '']])
      .mergeCells(0, 0, 0, 4)
      .setAlignment(0, 0, 'center')
      // Report Info
      .addDataRows([['Generated: March 2, 2026', '', '', '', '']])
      .mergeCells(1, 0, 1, 4)
      .setAlignment(1, 0, 'center')
      // Empty row
      .addDataRows([['', '', '', '', '']])
      // Parts Section Header
      .addDataRows([['PARTS INVENTORY', '', '', '', '']])
      .mergeCells(3, 0, 3, 4)
      .setAlignment(3, 0, 'left')
      // Parts Header
      .addHeaderRow(['Part Number', 'Part Name', 'Stock', 'Price', 'Total Value'])
      .setRangeAlignment(4, 0, 4, 4, 'center')
      // Parts Data
      .addDataRows(parts.map(p => [
        p.part_number,
        p.part_name,
        p.current_stock,
        p.price,
        (p.current_stock * p.price).toFixed(2)
      ]))
      .setRangeAlignment(5, 2, 7, 4, 'right')
      // Parts Total
      .addDataRows([['', '', 'TOTAL:', '', parts.reduce((sum, p) => sum + (p.current_stock * p.price), 0).toFixed(2)]])
      .setAlignment(8, 2, 'right')
      .setAlignment(8, 4, 'right')
      // Empty row
      .addDataRows([['', '', '', '', '']])
      // Instances Section Header
      .addDataRows([['PART INSTANCES', '', '', '', '']])
      .mergeCells(10, 0, 10, 4)
      .setAlignment(10, 0, 'left')
      // Instances Header
      .addHeaderRow(['Serial Number', 'Part Number', 'Status', 'Location', ''])
      .setRangeAlignment(11, 0, 11, 3, 'center')
      // Instances Data
      .addDataRows(instances.map(i => [i.serial_number, i.part_number, i.status, i.location, '']))
      .setRangeAlignment(12, 0, 15, 3, 'left')
      .download({ filename: 'complete-report.xlsx', format: 'xlsx' });
  };

  // 9. Styled Headers Export
  const handleStyledExport = () => {
    const headerStyle = StyleBuilder.create()
      .bold()
      .backgroundColor('#4F81BD')
      .color('#FFFFFF')
      .align('center')
      .build();

    const rowStyle = StyleBuilder.create()
      .color('#333')
      .backgroundColor('#e0e0e0')
      .align('left')
      .build();

    ExportBuilder.create('Styled Report')
      .setColumnWidths([150, 100, 150, 120])
      .addHeaderRow(['Name', 'Age', 'Email', 'Department'], headerStyle)
      .addDataRows(
        users.map(u => [u.name, u.age, u.email, u.department]),
        undefined,
        users.map(() => [rowStyle, rowStyle, rowStyle, rowStyle])
      )
      .download({ filename: 'styled-report.xlsx', format: 'xlsx' });
  };

  return (
    <div
      className="App"
      style={{
        fontFamily: 'Tahoma, Geneva, Verdana, sans-serif',
        padding: 32,
        background: '#f4f4f4',
        minHeight: '100vh',
        boxSizing: 'border-box',
      }}
    >
      <h1 style={{ color: '#222', marginBottom: 8, fontWeight: 'bold', fontSize: 28, borderBottom: '2px solid #bbb', paddingBottom: 8 }}>
        Sheetra Demo - All Features
      </h1>
      <p style={{ color: '#666', marginBottom: 24 }}>
        Each table below has export buttons - click to export with different features.
      </p>

      {/* Data Preview Section */}
      <section style={sectionStyle}>
        <h2 style={sectionTitleStyle}>Data Preview &amp; Export</h2>
        <div style={{ display: 'flex', gap: 32, flexWrap: 'wrap' }}>
          {/* Users Table */}
          <div>
            <h3 style={tableTitleStyle}>Users</h3>
            <table style={tableStyle}>
              <thead>
                <tr>
                  <th style={thStyle}>Name</th>
                  <th style={thStyle}>Age</th>
                  <th style={thStyle}>Email</th>
                  <th style={thStyle}>Department</th>
                </tr>
              </thead>
              <tbody>
                {users.map((u, i) => (
                  <tr key={i}>
                    <td style={tdStyle}>{u.name}</td>
                    <td style={tdStyle}>{u.age}</td>
                    <td style={tdStyle}>{u.email}</td>
                    <td style={tdStyle}>{u.department}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div style={{ marginTop: 12, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              <button onClick={handleSimpleExportCsv} style={smallButtonStyle}>CSV</button>
              <button onClick={handleSimpleExportXlsx} style={smallButtonStyle}>XLSX</button>
              <button onClick={handleJsonExport} style={smallButtonStyle}>JSON</button>
              <button onClick={handleStyledExport} style={{ ...smallButtonStyle, background: '#e8d0ff' }}>Styled</button>
            </div>
          </div>

          {/* Parts Table */}
          <div>
            <h3 style={tableTitleStyle}>Parts</h3>
            <table style={tableStyle}>
              <thead>
                <tr>
                  <th style={thStyle}>Part Number</th>
                  <th style={thStyle}>Part Name</th>
                  <th style={thStyle}>Stock</th>
                  <th style={thStyle}>Price</th>
                </tr>
              </thead>
              <tbody>
                {parts.map((p, i) => (
                  <tr key={i}>
                    <td style={tdStyle}>{p.part_number}</td>
                    <td style={tdStyle}>{p.part_name}</td>
                    <td style={tdStyle}>{p.current_stock}</td>
                    <td style={tdStyle}>${p.price}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div style={{ marginTop: 12, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              <button onClick={handleSectionExport} style={{ ...smallButtonStyle, background: '#e8d0ff' }}>Section Export</button>
              <button onClick={handleComplexExport} style={{ ...smallButtonStyle, background: '#90EE90' }}>Full Report</button>
            </div>
          </div>

          {/* Sales Table */}
          <div>
            <h3 style={tableTitleStyle}>Sales Data</h3>
            <table style={tableStyle}>
              <thead>
                <tr>
                  <th style={thStyle}>Product</th>
                  <th style={thStyle}>Jan</th>
                  <th style={thStyle}>Feb</th>
                  <th style={thStyle}>Mar</th>
                </tr>
              </thead>
              <tbody>
                {salesData.map((s, i) => (
                  <tr key={i}>
                    <td style={tdStyle}>{s.product}</td>
                    <td style={tdStyle}>${s.jan}</td>
                    <td style={tdStyle}>${s.feb}</td>
                    <td style={tdStyle}>${s.mar}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div style={{ marginTop: 12, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              <button onClick={handleMergeCellsExport} style={{ ...smallButtonStyle, background: '#d0e8ff' }}>Merge Cells</button>
              <button onClick={handleAlignmentExport} style={{ ...smallButtonStyle, background: '#d0e8ff' }}>Alignment</button>
            </div>
          </div>

          {/* Auto Size Demo */}
          <div>
            <h3 style={tableTitleStyle}>Auto Size Demo</h3>
            <table style={tableStyle}>
              <thead>
                <tr>
                  <th style={thStyle}>Short</th>
                  <th style={thStyle}>Medium Header</th>
                  <th style={thStyle}>Long Header Text</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td style={tdStyle}>A</td>
                  <td style={tdStyle}>Medium text</td>
                  <td style={tdStyle}>Longer content here</td>
                </tr>
                <tr>
                  <td style={tdStyle}>B</td>
                  <td style={tdStyle}>Another</td>
                  <td style={tdStyle}>Short</td>
                </tr>
              </tbody>
            </table>
            <div style={{ marginTop: 12, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              <button onClick={handleAutoSizeExport} style={smallButtonStyle}>Auto-Size Columns</button>
            </div>
          </div>
        </div>
      </section>

      {/* Feature List */}
      <section style={{ ...sectionStyle, marginTop: 40 }}>
        <h2 style={sectionTitleStyle}>Available Features</h2>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(250px, 1fr))', gap: 16 }}>
          <FeatureCard title="Column Widths" desc="setColumnWidths([100, 200])" />
          <FeatureCard title="Auto Size" desc="autoSizeColumns()" />
          <FeatureCard title="Merge Cells" desc="mergeCells(row1, col1, row2, col2)" />
          <FeatureCard title="Alignment" desc="setAlignment(row, col, 'center')" />
          <FeatureCard title="Range Alignment" desc="setRangeAlignment(...)" />
          <FeatureCard title="Header Rows" desc="addHeaderRow(['A', 'B'])" />
          <FeatureCard title="Data Rows" desc="addDataRows([[1, 2], [3, 4]])" />
          <FeatureCard title="Sections" desc="addSection({ name, title })" />
          <FeatureCard title="Styled Headers" desc="addHeaderRow(headers, style)" />
          <FeatureCard title="CSV Export" desc="format: 'csv'" />
          <FeatureCard title="XLSX Export" desc="format: 'xlsx'" />
          <FeatureCard title="JSON Export" desc="format: 'json'" />
        </div>
      </section>
    </div>
  );
}

// Feature Card Component
const FeatureCard = ({ title, desc }: { title: string; desc: string }) => (
  <div style={{
    background: '#fff',
    border: '1px solid #ddd',
    borderRadius: 8,
    padding: 16,
    boxShadow: '0 2px 4px rgba(0,0,0,0.05)'
  }}>
    <strong style={{ color: '#333' }}>{title}</strong>
    <code style={{
      display: 'block',
      marginTop: 8,
      background: '#f5f5f5',
      padding: 8,
      borderRadius: 4,
      fontSize: 12,
      color: '#666'
    }}>{desc}</code>
  </div>
);

// Styles
const sectionStyle: React.CSSProperties = {
  background: '#fff',
  borderRadius: 8,
  padding: 24,
  marginBottom: 20,
  boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
};

const sectionTitleStyle: React.CSSProperties = {
  margin: 0,
  marginBottom: 8,
  fontSize: 20,
  color: '#333',
};

const smallButtonStyle: React.CSSProperties = {
  background: '#e0e0e0',
  color: '#222',
  border: '1px solid #bbb',
  borderRadius: 4,
  padding: '8px 14px',
  fontSize: 12,
  fontWeight: 500,
  cursor: 'pointer',
};

const tableTitleStyle: React.CSSProperties = {
  fontSize: 16,
  fontWeight: 600,
  color: '#444',
  marginBottom: 8,
  borderLeft: '3px solid #4472C4',
  paddingLeft: 8,
};

const tableStyle: React.CSSProperties = {
  borderCollapse: 'collapse',
  background: '#fff',
  border: '1px solid #ddd',
  fontSize: 13,
};

const thStyle: React.CSSProperties = {
  background: '#4472C4',
  color: '#fff',
  padding: '10px 14px',
  textAlign: 'left',
  fontWeight: 600,
  borderBottom: '2px solid #3461a9',
};

const tdStyle: React.CSSProperties = {
  padding: '8px 14px',
  borderBottom: '1px solid #eee',
  color: '#333',
};

export default App;