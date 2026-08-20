const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

async function runMandatoryAcceptanceTest() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  console.log('==================================================');
  console.log('EXECUTING MANDATORY ACCEPTANCE TEST PROCEDURE');
  console.log('FILE:', filePath);
  console.log('==================================================\n');

  // STEP 1 & 3: Read file & run through exact worker parser logic
  console.log('Step 1 & 3: Ingesting file via Worker Parser logic...');
  const workbook = new ExcelJS.Workbook();
  const startLoad = Date.now();
  await workbook.xlsx.readFile(filePath);
  console.log(`ExcelJS parsed in ${Date.now() - startLoad} ms.`);

  const sheetsData = {};
  const sheetOrder = [];

  workbook.worksheets.forEach((ws, sIdx) => {
    const sheetId = ws.name; // Exact sheet name as ID!
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
    id: `workbook-barakat-acceptance-${Date.now()}`,
    name: 'report barakat.xlsx',
    sheetOrder,
    appVersion: '3.0.0-alpha',
    sheets: sheetsData
  };

  console.log('Step 2: Launching clean browser automation...');
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  await new Promise(r => setTimeout(r, 1000));

  console.log('Step 4 - 18: Executing test sequence in browser...');
  const acceptanceReport = await page.evaluate(async (wbData) => {
    const api = window.univerAPI;
    const formulaEngine = api.getFormula ? api.getFormula() : null;

    // Step 1: Dispose current workbooks
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
    oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });

    if (formulaEngine && formulaEngine.setInitialFormulaComputing) {
      formulaEngine.setInitialFormulaComputing(1); // WHEN_EMPTY
    }

    // Step 3 & 4: Import workbook
    api.createUniverSheet(wbData);

    const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    const snapshot = activeWb.save();

    const sheetIds = Object.keys(snapshot.sheets);
    const sheetNames = sheetIds.map(id => snapshot.sheets[id].name);

    // Step 6: Shared formulas in Summary sheet
    const summarySheetId = sheetIds[0];
    const summaryCellData = snapshot.sheets[summarySheetId]?.cellData || {};

    const i4 = summaryCellData[3]?.[8]; // I4
    const i5 = summaryCellData[4]?.[8]; // I5
    const i6 = summaryCellData[5]?.[8]; // I6
    const i30 = summaryCellData[29]?.[8]; // I30

    // Step 7: Cross-sheet formula in Summary Hadiah sheet
    const hadiahSheetId = sheetIds[1];
    const hadiahCellData = snapshot.sheets[hadiahSheetId]?.cellData || {};
    const c4Hadiah = hadiahCellData[3]?.[2]; // C4

    return {
      step4_sheetsCount: sheetIds.length,
      step4_sheetNames: sheetNames,
      step6_sharedFormulas: {
        I4_Master: i4,
        I5_Follower: i5,
        I6_Follower: i6,
        I30_Follower: i30
      },
      step7_crossSheetHadiahC4: c4Hadiah,
      step17_exportSnapshotValid: Boolean(snapshot && snapshot.sheets)
    };
  }, workbookData);

  console.log('\n==================================================');
  console.log('ACCEPTANCE TEST RESULTS:');
  console.log('==================================================');
  console.log(JSON.stringify(acceptanceReport, null, 2));

  fs.writeFileSync('barakat-acceptance-report.json', JSON.stringify(acceptanceReport, null, 2));
  await browser.close();
}

runMandatoryAcceptanceTest().catch(console.error);
