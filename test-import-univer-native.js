const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

async function testImportNative() {
  console.log('--- 1. Creating Excel File Without Cached Formula Results ---');
  const wb = new ExcelJS.Workbook();
  const s1 = wb.addWorksheet('Sheet1');
  s1.getCell('A1').value = 100;
  s1.getCell('A2').value = 200;
  // Intentionally set formula ONLY, result = undefined
  s1.getCell('A3').value = { formula: 'A1+A2' };

  const s2 = wb.addWorksheet('Sheet2');
  s2.getCell('A1').value = { formula: 'Sheet1!A1' };
  s2.getCell('A2').value = { formula: 'Sheet1!A2+50' };

  const excelPath = path.join(__dirname, 'test-no-cache.xlsx');
  await wb.xlsx.writeFile(excelPath);
  console.log('Created test-no-cache.xlsx at:', excelPath);

  console.log('\n--- 2. Parsing with Custom Worker Logic ---');
  const fileBuffer = fs.readFileSync(excelPath);
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
            v = res;
          }
        } else if (cell.value !== undefined && cell.value !== null) {
          v = cell.value;
        }

        const cellInfo = {};
        if (f !== undefined) cellInfo.f = f;
        if (v !== undefined) {
          if (typeof v === 'number') { cellInfo.v = v; cellInfo.t = 2; }
          else if (typeof v === 'boolean') { cellInfo.v = v; cellInfo.t = 3; }
          else { cellInfo.v = String(v); cellInfo.t = 1; }
        }

        if (Object.keys(cellInfo).length > 0) {
          rowCellMap[colIndex] = cellInfo;
        }
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

  const parsedWorkbookData = {
    id: `workbook-imported-${Date.now()}`,
    name: 'test-no-cache.xlsx',
    sheetOrder,
    appVersion: '3.0.0-alpha',
    sheets: sheetsData
  };

  console.log('Parsed workbookData sheet1 A3:', JSON.stringify(sheetsData['sheet1']?.cellData?.[2]?.[0]));
  console.log('Parsed workbookData sheet2 A1:', JSON.stringify(sheetsData['sheet2']?.cellData?.[0]?.[0]));

  console.log('\n--- 3. Testing Import in Browser via replaceUniverWorkbook ---');
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  await new Promise(r => setTimeout(r, 1000));

  const importResult = await page.evaluate(async (wbData) => {
    const api = window.univerAPI;
    
    // Dispose old units
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
    oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });

    // Set calculation mode to WHEN_EMPTY (1) just to be sure
    const formulaEngine = api.getFormula ? api.getFormula() : null;
    if (formulaEngine && formulaEngine.setInitialFormulaComputing) {
      formulaEngine.setInitialFormulaComputing(1);
    }

    // Call createUniverSheet (which replaceUniverWorkbook does)
    api.createUniverSheet(wbData);

    const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];

    // Immediate state
    const immA3 = activeWb.getSheetByName('Sheet1')?.getRange(2, 0, 1, 1).getValue();
    const immS2A1 = activeWb.getSheetByName('Sheet2')?.getRange(0, 0, 1, 1).getValue();
    const immS2A2 = activeWb.getSheetByName('Sheet2')?.getRange(1, 0, 1, 1).getValue();

    // Wait 2 seconds for worker calculation to return
    await new Promise(r => setTimeout(r, 2000));

    const setA3 = activeWb.getSheetByName('Sheet1')?.getRange(2, 0, 1, 1).getValue();
    const setS2A1 = activeWb.getSheetByName('Sheet2')?.getRange(0, 0, 1, 1).getValue();
    const setS2A2 = activeWb.getSheetByName('Sheet2')?.getRange(1, 0, 1, 1).getValue();

    const snapshot = activeWb.save();

    return {
      immediate: { Sheet1_A3: immA3, Sheet2_A1: immS2A1, Sheet2_A2: immS2A2 },
      settled: { Sheet1_A3: setA3, Sheet2_A1: setS2A1, Sheet2_A2: setS2A2 },
      cellModels: {
        Sheet1_A3: snapshot.sheets[wbData.sheetOrder[0]]?.cellData?.[2]?.[0],
        Sheet2_A1: snapshot.sheets[wbData.sheetOrder[1]]?.cellData?.[0]?.[0],
        Sheet2_A2: snapshot.sheets[wbData.sheetOrder[1]]?.cellData?.[1]?.[0]
      }
    };
  }, parsedWorkbookData);

  fs.writeFileSync('native-import-out.json', JSON.stringify(importResult, null, 2));
  console.log('\n--- 4. IMPORT TEST RESULT ---');
  console.log(JSON.stringify(importResult, null, 2));

  await browser.close();
}

testImportNative().catch(console.error);
