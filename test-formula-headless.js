const { createUniver, LocaleType } = require('@univerjs/presets');
const { UniverSheetsCorePreset } = require('@univerjs/preset-sheets-core');
const { UniverFormulaEnginePlugin } = require('@univerjs/engine-formula');
const { UniverSheetsFormulaPlugin } = require('@univerjs/sheets-formula');
const { ICommandService, IUniverInstanceService } = require('@univerjs/core');
const { SetRangeValuesCommand } = require('@univerjs/sheets');

const { univer, univerAPI: api } = createUniver({
  locale: LocaleType.EN_US,
  presets: [
    UniverSheetsCorePreset({
      notExecuteFormula: false,
    }),
  ],
});

async function runTest() {
  console.log("1. Creating Workbook...");
  const wb = api.createUniverSheet({
    id: 'test-wb',
    sheets: {
      'sheet-1': {
        id: 'sheet-1',
        name: 'Sheet1',
        type: 2, // SheetType.GRID
        rowCount: 100,
        columnCount: 100,
        cellData: {}
      }
    }
  });

  console.log("Workbook created. Ready to test formulas.");

  // Type 100 in A1
  api.executeCommand(SetRangeValuesCommand.id, {
    unitId: 'test-wb',
    subUnitId: 'sheet-1',
    range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    value: { v: 100, t: 2 } // NUMBER
  });

  // Type 200 in A2
  api.executeCommand(SetRangeValuesCommand.id, {
    unitId: 'test-wb',
    subUnitId: 'sheet-1',
    range: { startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 },
    value: { v: 200, t: 2 }
  });

  // Type =A1+A2 in A3
  console.log("Typing =A1+A2 in A3...");
  api.executeCommand(SetRangeValuesCommand.id, {
    unitId: 'test-wb',
    subUnitId: 'sheet-1',
    range: { startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 },
    value: { f: '=A1+A2' }
  });

  // Wait a bit for formula async calculation
  await new Promise(r => setTimeout(r, 1000));

  const sheet = wb.getActiveSheet();
  const v1 = sheet.getRange(0, 0, 1, 1).getValue();
  const v2 = sheet.getRange(1, 0, 1, 1).getValue();
  const v3 = sheet.getRange(2, 0, 1, 1).getValue();

  console.log("A1 VALUE:", v1);
  console.log("A2 VALUE:", v2);
  console.log("A3 VALUE:", v3);

  // Type =1+1 in B1
  console.log("Typing =1+1 in B1...");
  api.executeCommand(SetRangeValuesCommand.id, {
    unitId: 'test-wb',
    subUnitId: 'sheet-1',
    range: { startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 },
    value: { f: '=1+1' }
  });

  await new Promise(r => setTimeout(r, 1000));
  const b1 = sheet.getRange(0, 1, 1, 1).getValue();
  console.log("B1 VALUE:", b1);
  
  process.exit(0);
}

runTest();
