const { CellValueType } = require('@univerjs/core');

// Test sheet structure comparison
const nativeSheetSpec = {
  id: 'sheet-1',
  name: 'Sheet1',
  type: 2, // SheetType.GRID (CRITICAL)
  status: 1,
  rowCount: 100,
  columnCount: 30,
  cellData: {
    '0': { '0': { v: 100, t: 2 }, '1': { v: 200, t: 2 } }
  }
};

const importedSheetSpecWithoutType = {
  id: 'sheet-imported-0-12345',
  name: 'Sheet1',
  // MISSING 'type: 2'!
  status: 1,
  rowCount: 100,
  columnCount: 30,
  cellData: {
    '0': { '0': { v: 100, t: 2 }, '1': { v: 200, t: 2 } }
  }
};

console.log('--- SHEET SPECIFICATION AUDIT ---');
console.log('Native Sheet Spec has "type":', nativeSheetSpec.type);
console.log('Imported Sheet Spec has "type":', importedSheetSpecWithoutType.type);
if (!importedSheetSpecWithoutType.type) {
  console.log('⚠️ WARNING: Imported Sheet Spec is MISSING "type: 2" (SheetType.GRID).');
  console.log('In Univer 0.25+, FormulaEngine and FormulaDataModel filter sheets by (type === 2). Sheets missing "type: 2" are ignored by FormulaEngine!');
}
