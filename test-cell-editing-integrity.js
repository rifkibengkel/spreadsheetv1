const ExcelJS = require('exceljs');

// Enum values from @univerjs/core CellValueType
const CellValueType = {
  STRING: 1,
  NUMBER: 2,
  BOOLEAN: 3,
  FORCE_STRING: 4
};

/**
 * Normalizes any cell raw value and formula into standard Univer ICellData model.
 */
function normalizeUniverCell(rawV, rawF) {
  const cellInfo = {};

  let f = undefined;
  if (rawF && typeof rawF === 'string' && rawF.trim().length > 0) {
    const trimmed = rawF.trim();
    f = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
    cellInfo.f = f;
  }

  let v = rawV;
  if (v !== undefined && v !== null && v !== '') {
    if (typeof v === 'number') {
      cellInfo.v = v;
      cellInfo.t = CellValueType.NUMBER;
    } else if (typeof v === 'boolean') {
      cellInfo.v = v;
      cellInfo.t = CellValueType.BOOLEAN;
    } else if (typeof v === 'string') {
      // If string is purely numeric and no formula, convert if needed or keep string
      const trimmedVal = v.trim();
      const numVal = Number(trimmedVal);
      if (!isNaN(numVal) && trimmedVal !== '' && !f) {
        cellInfo.v = numVal;
        cellInfo.t = CellValueType.NUMBER;
      } else {
        cellInfo.v = v;
        cellInfo.t = CellValueType.STRING;
      }
    } else {
      cellInfo.v = String(v);
      cellInfo.t = CellValueType.STRING;
    }
  }

  return cellInfo;
}

/**
 * Simulates Univer cell editor mutation (SetRangeValuesCommand) on a target cell.
 */
function simulateCellEdit(oldCell, editInput) {
  // If user enters a new formula (starts with =)
  if (typeof editInput === 'string' && editInput.trim().startsWith('=')) {
    return {
      f: editInput.trim(),
      v: null,
      t: null
    };
  }

  // If user enters a number or string value
  const numInput = Number(editInput);
  if (!isNaN(numInput) && editInput !== '' && typeof editInput !== 'boolean') {
    return {
      v: numInput,
      t: CellValueType.NUMBER,
      f: null
    };
  }

  if (typeof editInput === 'boolean') {
    return {
      v: editInput,
      t: CellValueType.BOOLEAN,
      f: null
    };
  }

  return {
    v: String(editInput),
    t: CellValueType.STRING,
    f: null
  };
}

async function runCellEditingTests() {
  console.log('🧪 RUNNING CELL EDITING INTEGRITY TESTS...');

  // TEST A: Native vs Imported Cell Normalization
  console.log('\n--- TEST A: Native vs Imported Normalization ---');
  const importedA1 = normalizeUniverCell(100, undefined);
  console.log('Imported A1 Normalized:', importedA1);
  if (importedA1.v !== 100 || importedA1.t !== CellValueType.NUMBER) {
    throw new Error('FAIL: Imported cell normalization failed for number 100!');
  }

  // TEST B: Editing Imported Number Cell (100 -> 200 -> 500 -> "TEST" -> 12345)
  console.log('\n--- TEST B: Sequential Editing on Imported Cell ---');
  let currentCell = normalizeUniverCell(100, undefined);
  console.log('Initial (Imported):', currentCell);

  currentCell = simulateCellEdit(currentCell, 200);
  console.log('After Edit 100 -> 200:', currentCell);
  if (currentCell.v !== 200 || currentCell.t !== CellValueType.NUMBER) {
    throw new Error('FAIL: Edit 100 -> 200 produced invalid state!');
  }

  currentCell = simulateCellEdit(currentCell, 500);
  console.log('After Edit 200 -> 500:', currentCell);
  if (currentCell.v !== 500 || currentCell.t !== CellValueType.NUMBER) {
    throw new Error('FAIL: Edit 200 -> 500 produced invalid state!');
  }

  currentCell = simulateCellEdit(currentCell, 'TEST');
  console.log('After Edit 500 -> "TEST":', currentCell);
  if (currentCell.v !== 'TEST' || currentCell.t !== CellValueType.STRING) {
    throw new Error('FAIL: Edit 500 -> "TEST" produced invalid state!');
  }

  currentCell = simulateCellEdit(currentCell, 12345);
  console.log('After Edit "TEST" -> 12345:', currentCell);
  if (currentCell.v !== 12345 || currentCell.t !== CellValueType.NUMBER) {
    throw new Error('FAIL: Edit "TEST" -> 12345 produced invalid state!');
  }

  // TEST C: Formula Cell Editing (Imported =A1+B1 -> edit to 999)
  console.log('\n--- TEST C: Imported Formula Editing ---');
  let formulaCell = normalizeUniverCell(300, '=A1+B1');
  console.log('Initial Formula Cell:', formulaCell);

  formulaCell = simulateCellEdit(formulaCell, 999);
  console.log('After Edit Formula -> 999:', formulaCell);
  if (formulaCell.v !== 999 || formulaCell.f !== null) {
    throw new Error('FAIL: Edit Formula -> 999 failed to clear old formula!');
  }

  // TEST D: Value -> Formula Editing (Imported 100 -> edit to =A1+B1)
  console.log('\n--- TEST D: Value -> Formula Editing ---');
  let valToFormCell = normalizeUniverCell(100, undefined);
  valToFormCell = simulateCellEdit(valToFormCell, '=A1+B1');
  console.log('After Edit Value 100 -> =A1+B1:', valToFormCell);
  if (valToFormCell.f !== '=A1+B1') {
    throw new Error('FAIL: Edit Value -> Formula failed to set formula string!');
  }

  console.log('\n🎉 ALL CELL EDITING INTEGRITY TESTS PASSED PERFECTLY!');
}

runCellEditingTests().catch((err) => {
  console.error('❌ Test Failed:', err);
  process.exit(1);
});
