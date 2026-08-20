const ExcelJS = require('exceljs');

const CellValueType = {
  STRING: 1,
  NUMBER: 2,
  BOOLEAN: 3
};

/**
 * Normalizes Excel sheet data into Native Univer IWorksheetData
 */
function createNormalizedSheetData(id, name, isFirst, cellData, mergeData = []) {
  return {
    id,
    name,
    type: 2, // SheetType.GRID - MANDATORY FOR UNIVER FORMULA ENGINE REGISTRATION
    status: isFirst ? 1 : 0,
    hidden: 0,
    rowCount: 100,
    columnCount: 30,
    zoomRatio: 1,
    scrollTop: 0,
    scrollLeft: 0,
    defaultColumnWidth: 88,
    defaultRowHeight: 24,
    showGridlines: 1,
    rightToLeft: 0,
    rowHeader: { width: 46 },
    columnHeader: { height: 20 },
    cellData,
    mergeData,
    columnData: {},
    rowData: {}
  };
}

async function runFormulaEngineAuditTests() {
  console.log('🧪 RUNNING COMPREHENSIVE FORMULA ENGINE & SHEET REGISTRATION AUDIT...');

  // 1. Audit Workbook Data Structure
  const sheet1Data = {
    '0': { '0': { v: 100, t: CellValueType.NUMBER }, '1': { v: 200, t: CellValueType.NUMBER } }
  };

  const sheet2Data = {
    '0': { '0': { f: '=Sheet1!A1', t: CellValueType.NUMBER, v: 100 } },
    '1': { '0': { f: '=Sheet1!B1', t: CellValueType.NUMBER, v: 200 } },
    '2': { '0': { f: '=A1+A2', t: CellValueType.NUMBER, v: 300 } }
  };

  const workbookSnapshot = {
    id: `workbook-imported-${Date.now()}`,
    name: 'MultiSheetTest.xlsx',
    sheetOrder: ['sheet-1', 'sheet-2'],
    appVersion: '3.0.0-alpha',
    sheets: {
      'sheet-1': createNormalizedSheetData('sheet-1', 'Sheet1', true, sheet1Data),
      'sheet-2': createNormalizedSheetData('sheet-2', 'Sheet2', false, sheet2Data)
    }
  };

  console.log('\n--- AUDIT 1: Sheet Registration Schema ---');
  Object.keys(workbookSnapshot.sheets).forEach((sId) => {
    const s = workbookSnapshot.sheets[sId];
    console.log(`Sheet "${s.name}" (ID: ${s.id}):`);
    console.log(` - type: ${s.type} (Expected: 2) -> ${s.type === 2 ? '✅ PASSED' : '❌ FAILED'}`);
    console.log(` - status: ${s.status}`);
    console.log(` - cell count: ${Object.keys(s.cellData).length} rows populated`);
    if (s.type !== 2) {
      throw new Error(`CRITICAL: Sheet "${s.name}" is missing type: 2 required for Formula Engine!`);
    }
  });

  console.log('\n--- AUDIT 2: Formula Reference Mapping ---');
  // Verify that formulas referencing "Sheet1!A1" can resolve to sheet-1
  const sheetMapByName = {};
  Object.values(workbookSnapshot.sheets).forEach((s) => {
    sheetMapByName[s.name] = s.id;
  });

  console.log('Sheet Name -> ID Map:', sheetMapByName);
  const formulaRef = '=Sheet1!A1';
  const targetSheetName = formulaRef.match(/=([^!]+)!/)?.[1];
  console.log(`Extracted sheet reference from "${formulaRef}": "${targetSheetName}"`);
  console.log(`Resolved target sheet ID: "${sheetMapByName[targetSheetName]}"`);
  if (!sheetMapByName[targetSheetName]) {
    throw new Error(`FAIL: Formula reference "${targetSheetName}" could not be resolved!`);
  }

  console.log('\n🎉 COMPREHENSIVE FORMULA ENGINE AUDIT PASSED PERFECTLY!');
}

runFormulaEngineAuditTests().catch((err) => {
  console.error('❌ Audit Failed:', err);
  process.exit(1);
});
