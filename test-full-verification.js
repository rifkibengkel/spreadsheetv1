const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

async function runFullVerification() {
  console.log('==================================================');
  console.log('RUNNING FULL FORMULA & PERFORMANCE VERIFICATION');
  console.log('==================================================\n');

  // Step 1: Create Test XLSX Files
  console.log('--- 1. Generating Test XLSX Files ---');
  
  // File A: XLSX WITH Cached Results
  const wbA = new ExcelJS.Workbook();
  const sA1 = wbA.addWorksheet('Sheet1');
  sA1.getCell('A1').value = 100;
  sA1.getCell('A2').value = 200;
  sA1.getCell('A3').value = { formula: 'A1+A2', result: 300 };
  const pathA = path.join(__dirname, 'test-with-cache.xlsx');
  await wbA.xlsx.writeFile(pathA);

  // File B: XLSX WITHOUT Cached Results
  const wbB = new ExcelJS.Workbook();
  const sB1 = wbB.addWorksheet('Sheet1');
  sB1.getCell('A1').value = 100;
  sB1.getCell('A2').value = 200;
  sB1.getCell('A3').value = { formula: 'A1+A2' }; // Result is undefined
  const pathB = path.join(__dirname, 'test-no-cache-verify.xlsx');
  await wbB.xlsx.writeFile(pathB);

  // File C/D: Multi-Sheet & Cross-Sheet Formulas
  const wbCD = new ExcelJS.Workbook();
  const s1 = wbCD.addWorksheet('Sheet1');
  s1.getCell('A1').value = 100;
  s1.getCell('A2').value = 200;

  const s2 = wbCD.addWorksheet('Sheet2');
  s2.getCell('A1').value = { formula: 'Sheet1!A1' };
  s2.getCell('A2').value = { formula: 'Sheet1!A1+Sheet1!A2' };
  const pathCD = path.join(__dirname, 'test-cross-sheet.xlsx');
  await wbCD.xlsx.writeFile(pathCD);

  console.log('Created test Excel files cleanly.\n');

  // Helper function to simulate parsing via worker logic
  async function parseExcelFile(filePath) {
    const fileBuffer = fs.readFileSync(filePath);
    const wbRead = new ExcelJS.Workbook();
    await wbRead.xlsx.load(fileBuffer);
    const sheetsData = {};
    const sheetOrder = [];

    wbRead.worksheets.forEach((ws, sIdx) => {
      const sheetId = ws.name.toLowerCase().replace(/[^a-z0-9]/g, '-') || `sheet-${sIdx + 1}`;
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
          }

          if (rawFormula && typeof rawFormula === 'string' && rawFormula.trim().length > 0) {
            const trimmed = rawFormula.trim();
            f = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
          }

          if (cell.type === ExcelJS.ValueType.Formula) {
            const res = cell.result;
            if (res !== undefined && res !== null) {
              if (typeof res === 'object') {
                if ((res).result !== undefined && (res).result !== null && (res).result !== '') {
                  v = (res).result;
                } else if ((res).error !== undefined) {
                  v = (res).error;
                }
              } else if (res !== '') {
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
        rowCount: 50,
        columnCount: 20,
        cellData
      };
    });

    return {
      id: `workbook-imported-${Date.now()}`,
      name: path.basename(filePath),
      sheetOrder,
      appVersion: '3.0.0-alpha',
      sheets: sheetsData
    };
  }

  const dataA = await parseExcelFile(pathA);
  const dataB = await parseExcelFile(pathB);
  const dataCD = await parseExcelFile(pathCD);

  console.log('Parsed Data B sheet1 A3 before calculation:', JSON.stringify(dataB.sheets['sheet1']?.cellData?.[2]?.[0]));

  // Step 2: Browser Verification
  console.log('\n--- 2. Launching Headless Browser for Test Evaluation ---');
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  await new Promise(r => setTimeout(r, 1000));

  const testResults = await page.evaluate(async (dA, dB, dCD) => {
    const api = window.univerAPI;
    const formulaEngine = api.getFormula ? api.getFormula() : null;

    function resetWb(wbData) {
      const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
      oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });
      if (formulaEngine && formulaEngine.setInitialFormulaComputing) {
        formulaEngine.setInitialFormulaComputing(1); // CalculationMode.WHEN_EMPTY
      }
      api.createUniverSheet(wbData);
      return api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    }

    // CASE A: External XLSX WITH Cached Formula
    const wbA = resetWb(dA);
    const sheetA = wbA.getActiveSheet();
    const caseA_val = sheetA.getRange(2, 0, 1, 1).getValue();
    const caseA_model = wbA.save().sheets[dA.sheetOrder[0]]?.cellData?.[2]?.[0];

    // CASE B: External XLSX WITHOUT Cached Formula Result
    const wbB = resetWb(dB);
    const sheetB = wbB.getActiveSheet();
    const caseB_immVal = sheetB.getRange(2, 0, 1, 1).getValue();
    const caseB_immModel = wbB.save().sheets[dB.sheetOrder[0]]?.cellData?.[2]?.[0];

    // Wait 2 seconds for worker calculation to settle
    await new Promise(r => setTimeout(r, 2000));
    const caseB_settledVal = sheetB.getRange(2, 0, 1, 1).getValue();
    const caseB_settledModel = wbB.save().sheets[dB.sheetOrder[0]]?.cellData?.[2]?.[0];

    // CASE C & D: Cross-Sheet Formulas
    const wbCD = resetWb(dCD);
    await new Promise(r => setTimeout(r, 2000));
    const sheetCD_S2 = wbCD.getSheetByName('Sheet2');
    const caseC_initialVal = sheetCD_S2.getRange(0, 0, 1, 1).getValue(); // =Sheet1!A1
    const caseD_initialVal = sheetCD_S2.getRange(1, 0, 1, 1).getValue(); // =Sheet1!A1+Sheet1!A2

    // Edit dependency cell Sheet1!A1 to 500
    const sheetCD_S1 = wbCD.getSheetByName('Sheet1');
    sheetCD_S1.getRange(0, 0, 1, 1).setValue(500);

    // Wait for reactive calculation update
    await new Promise(r => setTimeout(r, 1500));
    const caseC_updatedVal = sheetCD_S2.getRange(0, 0, 1, 1).getValue(); // Should be 500
    const caseD_updatedVal = sheetCD_S2.getRange(1, 0, 1, 1).getValue(); // Should be 700 (500+200)

    // CASE E: Manual formula after import
    const rangeA3 = sheetCD_S1.getRange(2, 0, 1, 1);
    rangeA3.setValue({ f: '=A1+A2' }); // 500 + 200
    await new Promise(r => setTimeout(r, 1500));
    const caseE_val = rangeA3.getValue();

    // Modify A1 to 1000
    sheetCD_S1.getRange(0, 0, 1, 1).setValue(1000);
    await new Promise(r => setTimeout(r, 1500));
    const caseE_updatedVal = rangeA3.getValue(); // Should be 1200 (1000+200)

    return {
      CaseA: { displayedVal: caseA_val, cellModel: caseA_model },
      CaseB: {
        beforeCalc: { displayedVal: caseB_immVal, cellModel: caseB_immModel },
        afterCalc: { displayedVal: caseB_settledVal, cellModel: caseB_settledModel }
      },
      CaseC: { initial: caseC_initialVal, afterDependencyEdit: caseC_updatedVal },
      CaseD: { initial: caseD_initialVal, afterDependencyEdit: caseD_updatedVal },
      CaseE: { initialFormula: caseE_val, afterDependencyEdit: caseE_updatedVal }
    };
  }, dataA, dataB, dataCD);

  fs.writeFileSync('full-verification-results.json', JSON.stringify(testResults, null, 2));
  console.log('--- 3. FULL VERIFICATION RESULTS ---');
  console.log(JSON.stringify(testResults, null, 2));

  await browser.close();
}

runFullVerification().catch(console.error);
