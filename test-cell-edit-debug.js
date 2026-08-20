const { CellValueType } = require('@univerjs/core');
const ExcelJS = require('exceljs');

console.log('CellValueType Enum values:', CellValueType);

// Let's inspect how ExcelJS cell values map to Univer ICellData
function normalizeCell(v, f) {
  const cellInfo = {};
  if (f !== undefined && f !== null && f !== '') {
    cellInfo.f = f;
  }

  if (v !== undefined && v !== null && v !== '') {
    cellInfo.v = v;
    if (typeof v === 'number') {
      cellInfo.t = CellValueType ? CellValueType.NUMBER : 2; // 2
    } else if (typeof v === 'boolean') {
      cellInfo.t = CellValueType ? CellValueType.BOOLEAN : 3; // 3
    } else {
      cellInfo.t = CellValueType ? CellValueType.STRING : 1; // 1
    }
  } else if (f) {
    // If cell has formula but no value yet, type can still be defined or omitted
  }

  return cellInfo;
}

console.log('Normalized Number Cell (100):', normalizeCell(100, undefined));
console.log('Normalized String Cell ("Hello"):', normalizeCell("Hello", undefined));
console.log('Normalized Formula Cell (100, "=A1+B1"):', normalizeCell(100, "=A1+B1"));
