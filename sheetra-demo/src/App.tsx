import React from 'react';
import { ExportBuilder } from 'sheetra';
import './App.css';

const parts = [
  { part_number: 'P001', part_name: 'Widget', current_stock: 100 },
  { part_number: 'P002', part_name: 'Gadget', current_stock: 50 },
];

const instances = [
  { serial_number: 'S001', part_number: 'P001', status: 'Active', location: 'A1' },
  { serial_number: 'S002', part_number: 'P001', status: 'Inactive', location: 'A2' },
  { serial_number: 'S003', part_number: 'P002', status: 'Active', location: 'B1' },
];

function App() {
  const handleSimpleExportCsv = () => {
    const builder = ExportBuilder.create('Users')
      .addHeaderRow(['Name', 'Age'])
      .addDataRows([
        ['John', 30],
        ['Jane', 25]
      ]);
    builder.download({ filename: 'users.csv', format: 'csv' });
  };

  const handleSimpleExportXlsx = () => {
    const builder = ExportBuilder.create('Users')
      .addHeaderRow(['Name', 'Age'])
      .addDataRows([
        ['John', 30],
        ['Jane', 25]
      ]);
    builder.download({ filename: 'users.xlsx', format: 'xlsx' });
  };

  const handleAdvancedExportCsv = () => {
    ExportBuilder.create('Inventory')
      .addSection({ name: 'Parts' })
      .addHeaderRow(['Part Number', 'Part Name', 'Current Stock'])
      .addDataRows(parts.map(p => [p.part_number, p.part_name, p.current_stock]))
      .addSection({ name: 'Instances' })
      .addHeaderRow(['Serial Number', 'Part Number', 'Status', 'Location'])
      .addDataRows(instances.map(i => [i.serial_number, i.part_number, i.status, i.location]))
      .download({ filename: 'inventory.csv', format: 'csv' });
  };

  const handleAdvancedExportXlsx = () => {
    ExportBuilder.create('Inventory')
      .addSection({ name: 'Parts' })
      .addHeaderRow(['Part Number', 'Part Name', 'Current Stock'])
      .addDataRows(parts.map(p => [p.part_number, p.part_name, p.current_stock]))
      .addSection({ name: 'Instances' })
      .addHeaderRow(['Serial Number', 'Part Number', 'Status', 'Location'])
      .addDataRows(instances.map(i => [i.serial_number, i.part_number, i.status, i.location]))
      .download({ filename: 'inventory.xlsx', format: 'xlsx' });
  };

  return (
    <div
      className="App"
      style={{
        fontFamily: 'Tahoma, Geneva, Verdana, sans-serif',
        padding: 32,
        background: '#f4f4f4',
        minHeight: '100vh',
        minWidth: '100vw',
        width: '100vw',
        boxSizing: 'border-box',
      }}
    >
      <h1 style={{ color: '#222', marginBottom: 24, fontWeight: 'bold', fontSize: 28, borderBottom: '2px solid #bbb', paddingBottom: 8 }}>
        Sheetra Classic Demo
      </h1>
      <div style={{ marginBottom: 32 }}>
        <button onClick={handleSimpleExportCsv} style={classicButtonStyle}>Export Users (CSV)</button>
        <button onClick={handleSimpleExportXlsx} style={{ ...classicButtonStyle, marginLeft: 12 }}>Export Users (XLSX)</button>
        <button onClick={handleAdvancedExportCsv} style={{ ...classicButtonStyle, marginLeft: 12 }}>Export Inventory (CSV)</button>
        <button onClick={handleAdvancedExportXlsx} style={{ ...classicButtonStyle, marginLeft: 12 }}>Export Inventory (XLSX)</button>
      </div>
      <div style={{ display: 'flex', gap: 48, flexWrap: 'wrap', width: '100%' }}>
        <div>
          <h2 style={classicSectionTitleStyle}>Users</h2>
          <table style={classicTableStyle}>
            <thead>
              <tr>
                <th style={classicThTdStyle}>Name</th>
                <th style={classicThTdStyle}>Age</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td style={classicThTdStyle}>John</td>
                <td style={classicThTdStyle}>30</td>
              </tr>
              <tr>
                <td style={classicThTdStyle}>Jane</td>
                <td style={classicThTdStyle}>25</td>
              </tr>
            </tbody>
          </table>
        </div>
        <div>
          <h2 style={classicSectionTitleStyle}>Parts</h2>
          <table style={classicTableStyle}>
            <thead>
              <tr>
                <th style={classicThTdStyle}>Part Number</th>
                <th style={classicThTdStyle}>Part Name</th>
                <th style={classicThTdStyle}>Current Stock</th>
              </tr>
            </thead>
            <tbody>
              {parts.map((p) => (
                <tr key={p.part_number}>
                  <td style={classicThTdStyle}>{p.part_number}</td>
                  <td style={classicThTdStyle}>{p.part_name}</td>
                  <td style={classicThTdStyle}>{p.current_stock}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <h2 style={{ ...classicSectionTitleStyle, marginTop: 32 }}>Instances</h2>
          <table style={classicTableStyle}>
            <thead>
              <tr>
                <th style={classicThTdStyle}>Serial Number</th>
                <th style={classicThTdStyle}>Part Number</th>
                <th style={classicThTdStyle}>Status</th>
                <th style={classicThTdStyle}>Location</th>
              </tr>
            </thead>
            <tbody>
              {instances.map((i) => (
                <tr key={i.serial_number}>
                  <td style={classicThTdStyle}>{i.serial_number}</td>
                  <td style={classicThTdStyle}>{i.part_number}</td>
                  <td style={classicThTdStyle}>{i.status}</td>
                  <td style={classicThTdStyle}>{i.location}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

const classicButtonStyle: React.CSSProperties = {
  background: '#e0e0e0',
  color: '#222',
  border: '1px solid #bbb',
  borderRadius: 4,
  padding: '10px 22px',
  fontSize: 15,
  fontWeight: 600,
  cursor: 'pointer',
  boxShadow: '0 1px 2px #ccc',
  marginBottom: 8,
};

const classicSectionTitleStyle: React.CSSProperties = {
  color: '#333',
  marginBottom: 10,
  marginTop: 0,
  fontSize: 18,
  fontWeight: 700,
  letterSpacing: 0.2,
  borderLeft: '4px solid #bbb',
  paddingLeft: 8,
};

const classicTableStyle: React.CSSProperties = {
  borderCollapse: 'collapse',
  width: 380,
  background: '#fff',
  border: '1px solid #bbb',
  borderRadius: 4,
  marginBottom: 16,
};

const classicThTdStyle: React.CSSProperties = {
  border: '1px solid #bbb',
  padding: '8px 12px',
  textAlign: 'left',
  fontSize: 15,
  background: '#f9f9f9',
  color: '#000', // Ensure table text is black
};

export default App;