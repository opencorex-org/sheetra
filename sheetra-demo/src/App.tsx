import React from 'react';
import { ExportBuilder } from '../../../sheetra';
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
  const handleSimpleExport = () => {
    const data = [
      { name: 'John', age: 30 },
      { name: 'Jane', age: 25 }
    ];
    ExportBuilder.create('Users')
      .addHeaderRow(['Name', 'Age'])
      .addDataRows(data)
      .download({ filename: 'users.xlsx' });
  };

  const handleAdvancedExport = () => {
    const builder = ExportBuilder.create('Inventory');
    builder.addSection({
      title: 'Parts',
      level: 0,
      data: parts,
      fields: ['part_number', 'part_name', 'current_stock']
    });
    builder.addSection({
      title: 'Instances',
      level: 0,
      data: instances,
      groupBy: 'part_number',
      fields: ['serial_number', 'status', 'location']
    });
    builder.download({ filename: 'inventory.xlsx' });
  };

  return (
    <div className="App">
      <h1>Welcome to Sheetra Demo</h1>
      <button onClick={handleSimpleExport}>Export Users (Simple)</button>
      <button onClick={handleAdvancedExport} style={{ marginLeft: 8 }}>Export Inventory (Advanced)</button>
    </div>
  );
}

export default App;
