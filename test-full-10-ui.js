const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

async function runAll10UITests() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  console.log('==================================================');
  console.log('STARTING FULL 10 USER-FACING UI TESTS');
  console.log('Target File:', filePath);
  console.log('==================================================\n');

  console.log('Step 1: Reading file & processing via Worker ExcelJS logic...');
  const workbook = new ExcelJS.Workbook();
  const startLoad = Date.now();
  await workbook.xlsx.readFile(filePath);
  console.log(`ExcelJS parsed file in ${Date.now() - startLoad} ms.`);

  const sheetsData = {};
  const sheetOrder = [];

  workbook.worksheets.forEach((ws, sIdx) => {
    const sheetId = ws.name; // Preserving exact worksheet name as sheet ID!
    sheetOrder.push(sheetId);
    const cellData = {};

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const rowIndex = rowNumber - 1;
      const rowCellMap = {};

      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const colIndex = colNumber - 1;
        let f = undefined;
        let v = undefined;

        let rawFormula = undefined;
        if (cell.type === ExcelJS.ValueType.Formula) {
          rawFormula = cell.formula || (cell.model && cell.model.formula) || (typeof cell.value === 'object' && cell.value !== null ? cell.value.formula : undefined);
        } else if (typeof cell.value === 'object' && cell.value !== null && cell.value.formula) {
          rawFormula = cell.value.formula;
        } else if (cell.model && cell.model.formula) {
          rawFormula = cell.model.formula;
        } else if (cell.formula) {
          rawFormula = cell.formula;
        }

        if (rawFormula && typeof rawFormula === 'string' && rawFormula.trim().length > 0) {
          const trimmed = rawFormula.trim();
          f = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
        }

        if (cell.type === ExcelJS.ValueType.Formula) {
          const res = cell.result;
          if (res !== undefined && res !== null && res !== '') {
            if (typeof res === 'object') {
              if ((res).result !== undefined && (res).result !== null && (res).result !== '') {
                v = (res).result;
              } else if ((res).error !== undefined) {
                v = (res).error;
              }
            } else {
              v = res;
            }
          }
        } else if (cell.value !== undefined && cell.value !== null) {
          const val = cell.value;
          if (typeof val === 'object') {
            if (val instanceof Date) v = val.toISOString();
            else if (Array.isArray(val.richText)) v = val.richText.map(t => t.text || '').join('');
            else if (val.text !== undefined) v = val.text;
            else if (val.result !== undefined) v = val.result;
            else v = String(val);
          } else {
            v = val;
          }
        }

        const cellInfo = {};
        if (f !== undefined && f !== null && f !== '') {
          const trimmed = f.trim();
          cellInfo.f = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
        }

        if (v !== undefined && v !== null && v !== '') {
          if (typeof v === 'number') { cellInfo.v = v; cellInfo.t = 2; }
          else if (typeof v === 'boolean') { cellInfo.v = v; cellInfo.t = 3; }
          else if (typeof v === 'string') {
            const trimmedV = v.trim();
            const numV = Number(trimmedV);
            if (!isNaN(numV) && trimmedV !== '') { cellInfo.v = numV; cellInfo.t = 2; }
            else { cellInfo.v = v; cellInfo.t = 1; }
          } else { cellInfo.v = String(v); cellInfo.t = 1; }
        }

        if (Object.keys(cellInfo).length > 0) rowCellMap[colIndex] = cellInfo;
      });

      if (Object.keys(rowCellMap).length > 0) cellData[rowIndex] = rowCellMap;
    });

    sheetsData[sheetId] = {
      id: sheetId,
      name: ws.name,
      type: 2,
      status: sIdx === 0 ? 1 : 0,
      rowCount: Math.max(100, ws.rowCount || 100),
      columnCount: Math.max(20, ws.columnCount || 20),
      cellData
    };
  });

  const workbookData = {
    id: `workbook-barakat-ui-${Date.now()}`,
    name: 'report barakat.xlsx',
    sheetOrder,
    appVersion: '3.0.0-alpha',
    sheets: sheetsData
  };

  console.log('Step 2: Launching clean browser automation...');
  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  const page = await browser.newPage();
  await page.setViewport({ width: 1400, height: 900 });

  page.on('console', msg => {
    if (msg.type() === 'error') console.log('BROWSER ERROR:', msg.text());
  });

  console.log('Navigating to http://localhost:3000...');
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  await new Promise(r => setTimeout(r, 1000));

  console.log('Step 3: Loading parsed workbook into Univer engine via window.univerAPI...');
  await page.evaluate(async (wbData) => {
    const api = window.univerAPI;
    const formulaEngine = api.getFormula ? api.getFormula() : null;

    // Dispose active workbooks
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
    oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });

    if (formulaEngine && formulaEngine.setInitialFormulaComputing) {
      formulaEngine.setInitialFormulaComputing(1); // CalculationMode.WHEN_EMPTY
    }

    api.createUniverSheet(wbData);
    await new Promise(r => setTimeout(r, 2000));
  }, workbookData);

  // Capture screenshot of loaded state
  await page.screenshot({ path: 'imported_barakat_ui.png' });
  console.log('Saved screenshot: imported_barakat_ui.png');

  console.log('\n--- EXECUTING ALL 10 USER-FACING UI TESTS ---');

  const uiTestResults = await page.evaluate(async () => {
    const api = window.univerAPI;
    const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    const results = {};

    // ----------------------------------------------------
    // TEST 1 — Imported shared formula
    // ----------------------------------------------------
    const summarySheet = activeWb.getSheetByName('Summary');
    const i4Val = summarySheet.getRange('I4').getValue();
    const i5Val = summarySheet.getRange('I5').getValue();
    const i6Val = summarySheet.getRange('I6').getValue();
    const i7Val = summarySheet.getRange('I7').getValue();
    const i30Val = summarySheet.getRange('I30').getValue();

    results.TEST1_SharedFormulas = {
      I4: i4Val,
      I5: i5Val,
      I6: i6Val,
      I7: i7Val,
      I30: i30Val,
      passed: (i5Val === 5 && i6Val === 105)
    };

    // ----------------------------------------------------
    // TEST 2 — Imported cross-sheet formula
    // ----------------------------------------------------
    const hadiahSheet = activeWb.getSheetByName('Summary Hadiah');
    const c4Val = hadiahSheet ? hadiahSheet.getRange('C4').getValue() : null;
    results.TEST2_CrossSheetFormula = {
      C4_SummaryHadiah: c4Val,
      passed: (c4Val !== null && c4Val !== undefined)
    };

    // ----------------------------------------------------
    // TEST 3 — MANUAL FORMULA AFTER IMPORT (=1+1, =10+20)
    // ----------------------------------------------------
    const testCell3 = summarySheet.getRange('Z1');
    testCell3.setValue('=1+1');
    await new Promise(r => setTimeout(r, 800));
    const valOnePlusOne = testCell3.getValue();

    testCell3.setValue('=10+20');
    await new Promise(r => setTimeout(r, 800));
    const valTenPlusTwenty = testCell3.getValue();

    results.TEST3_ManualFormula = {
      onePlusOne: valOnePlusOne,
      tenPlusTwenty: valTenPlusTwenty,
      passed: (valOnePlusOne === 2 && valTenPlusTwenty === 30)
    };

    // ----------------------------------------------------
    // TEST 4 — MANUAL FORMULA USING IMPORTED DATA (=A1+A2)
    // ----------------------------------------------------
    const cellA1 = summarySheet.getRange('A1');
    const cellA2 = summarySheet.getRange('A2');
    const testCell4 = summarySheet.getRange('Z2');

    cellA1.setValue(100);
    cellA2.setValue(200);
    testCell4.setValue('=A1+A2');
    await new Promise(r => setTimeout(r, 800));
    const initialSum = testCell4.getValue();

    // Modify A1 to 500
    cellA1.setValue(500);
    await new Promise(r => setTimeout(r, 800));
    const updatedSum = testCell4.getValue();

    results.TEST4_ManualFormulaImportedData = {
      initialSum,
      updatedSum,
      passed: (initialSum === 300 && updatedSum === 700)
    };

    // ----------------------------------------------------
    // TEST 5 — MANUAL CROSS-SHEET FORMULA
    // ----------------------------------------------------
    const testCell5a = summarySheet.getRange('Z3');
    const testCell5b = summarySheet.getRange('Z4');

    testCell5a.setValue("='Entries data'!A1");
    testCell5b.setValue("='Summary Hadiah'!B4");
    await new Promise(r => setTimeout(r, 800));

    const valCross1 = testCell5a.getValue();
    const valCross2 = testCell5b.getValue();

    results.TEST5_ManualCrossSheet = {
      refEntriesDataA1: valCross1,
      refSummaryHadiahB4: valCross2,
      passed: (valCross1 !== undefined && valCross2 !== undefined)
    };

    // ----------------------------------------------------
    // TEST 6 — EDIT AN EXISTING IMPORTED FORMULA
    // ----------------------------------------------------
    const targetCell6 = summarySheet.getRange('I5');
    const originalFormula = '=SUMIF($A$4:$A$188,$G5,C$4:C$188)';

    // Edit I5 temporarily to =1+1
    targetCell6.setValue('=1+1');
    await new Promise(r => setTimeout(r, 800));
    const valTemp = targetCell6.getValue();

    // Restore I5 formula
    targetCell6.setValue(originalFormula);
    await new Promise(r => setTimeout(r, 800));
    const valRestored = targetCell6.getValue();

    results.TEST6_EditExistingFormula = {
      valTemp,
      valRestored,
      passed: (valTemp === 2 && valRestored === 5)
    };

    // ----------------------------------------------------
    // TEST 7 — EDIT DEPENDENCY
    // ----------------------------------------------------
    const cellG5 = summarySheet.getRange('G5');
    const oldG5Val = cellG5.getValue();
    const oldI5Val = summarySheet.getRange('I5').getValue();

    // Temporarily change G5
    cellG5.setValue('NON_EXISTENT_CODE');
    await new Promise(r => setTimeout(r, 800));
    const newI5Val = summarySheet.getRange('I5').getValue();

    // Restore G5
    cellG5.setValue(oldG5Val);
    await new Promise(r => setTimeout(r, 800));
    const restoredI5Val = summarySheet.getRange('I5').getValue();

    results.TEST7_EditDependency = {
      oldI5Val,
      newI5Val,
      restoredI5Val,
      passed: (oldI5Val === 5 && newI5Val === 0 && restoredI5Val === 5)
    };

    // ----------------------------------------------------
    // TEST 8 — FORMULA WITH NO CACHED RESULT
    // ----------------------------------------------------
    const testCell8 = summarySheet.getRange('Z5');
    testCell8.setValue('=SUM(10, 20, 30)');
    await new Promise(r => setTimeout(r, 800));
    const uncachedVal = testCell8.getValue();

    results.TEST8_FormulaNoCachedResult = {
      uncachedVal,
      passed: (uncachedVal === 60)
    };

    // ----------------------------------------------------
    // TEST 9 — PAGE RESPONSIVENESS (Sheet Switching)
    // ----------------------------------------------------
    const sheets = activeWb.getSheets();
    const sheetSwitchLog = [];
    for (const s of sheets) {
      activeWb.setActiveSheet(s);
      await new Promise(r => setTimeout(r, 200));
      sheetSwitchLog.push(s.getName());
    }
    // Switch back to Summary sheet
    activeWb.setActiveSheet(summarySheet);

    results.TEST9_PageResponsiveness = {
      switchedSheets: sheetSwitchLog,
      passed: (sheetSwitchLog.length === 5)
    };

    // ----------------------------------------------------
    // TEST 10 — EXPORT / REIMPORT
    // ----------------------------------------------------
    const snapshot = activeWb.save();
    
    // Dispose current and re-create from saved snapshot
    api.disposeUnit(activeWb.getId());
    api.createUniverSheet(snapshot);

    await new Promise(r => setTimeout(r, 1000));

    const reimportedWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    const reimportedSummary = reimportedWb.getSheetByName('Summary');
    const reimportedI5 = reimportedSummary.getRange('I5').getValue();

    results.TEST10_ExportReimport = {
      reimportedI5Val: reimportedI5,
      passed: (reimportedI5 === 5)
    };

    return results;
  });

  console.log('\n==================================================');
  console.log('FULL 10 USER-FACING UI TEST RESULTS:');
  console.log('==================================================');
  console.log(JSON.stringify(uiTestResults, null, 2));

  fs.writeFileSync('full-10-ui-test-report.json', JSON.stringify(uiTestResults, null, 2));
  await browser.close();
}

runAll10UITests().catch(console.error);
