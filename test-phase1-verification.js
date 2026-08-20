const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

async function testPhase1Verification() {
  console.log('=== PHASE 1 VERIFICATION FOR report barakat.xlsx ===');
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';

  // 1. Parse using node ExcelJS to simulate worker with immediate release
  const fileBuf = fs.readFileSync(filePath);
  const startMem = process.memoryUsage().heapUsed;

  let workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(fileBuf);
  const parseMem = process.memoryUsage().heapUsed;

  const sheetOrder = [];
  const sheetsData = {};

  workbook.worksheets.forEach((ws, sIdx) => {
    sheetOrder.push(ws.name);
    const cellData = {};
    let maxRow = 0;
    let maxCol = 0;

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const rowIndex = rowNumber - 1;
      if (rowIndex > maxRow) maxRow = rowIndex;

      const rowCellMap = {};
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const colIndex = colNumber - 1;
        if (colIndex > maxCol) maxCol = colIndex;

        let v = undefined;
        let f = undefined;

        if (cell.type === ExcelJS.ValueType.Formula) {
          f = cell.formula ? (cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`) : undefined;
          if (cell.result !== undefined && cell.result !== null && cell.result !== '') {
            v = typeof cell.result === 'object' ? cell.result.result : cell.result;
          }
        } else if (cell.value !== undefined && cell.value !== null) {
          v = typeof cell.value === 'object' ? cell.value.text || cell.value.result : cell.value;
        }

        const cellInfo = {};
        if (f) cellInfo.f = f;
        if (v !== undefined && v !== null && v !== '') {
          cellInfo.v = typeof v === 'number' ? v : String(v);
          cellInfo.t = typeof v === 'number' ? 2 : 1;
        }
        if (Object.keys(cellInfo).length > 0) rowCellMap[colIndex] = cellInfo;
      });
      if (Object.keys(rowCellMap).length > 0) cellData[rowIndex] = rowCellMap;
    });

    sheetsData[ws.name] = {
      id: ws.name,
      name: ws.name,
      type: 2,
      cellData,
      rowCount: Math.max(maxRow + 50, 100),
      columnCount: Math.max(maxCol + 10, 30)
    };
  });

  // PHASE 1 REQUIREMENT: Immediate release of ExcelJS workbook
  workbook = null;
  if (global.gc) global.gc();
  const postReleaseMem = process.memoryUsage().heapUsed;

  console.log(`- Worker Parsing Memory Spike: +${((parseMem - startMem) / 1024 / 1024).toFixed(2)} MB`);
  console.log(`- Worker Memory AFTER Immediate ExcelJS Release: ${((postReleaseMem) / 1024 / 1024).toFixed(2)} MB (reclaimed ${((parseMem - postReleaseMem) / 1024 / 1024).toFixed(2)} MB!)`);

  const workbookData = {
    id: `phase1-imported-${Date.now()}`,
    name: 'report barakat.xlsx',
    sheetOrder,
    sheets: sheetsData
  };

  // Launch browser to test browser UI import and export performance
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();

  page.on('console', msg => console.log('[BROWSER]', msg.text()));

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Testing browser import & snapshot creation...');
  const browserResults = await page.evaluate(async (wbData) => {
    const api = window.univerAPI;
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
    oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

    const wb = api.createUniverSheet(wbData);
    await new Promise(r => setTimeout(r, 2000));

    // Check key formulas
    const sHadiah = wb.getSheetByName('Summary Hadiah');
    const fA1 = sHadiah ? sHadiah.getRange(0, 0, 1, 1).getFormula() : null;
    const fC4 = sHadiah ? sHadiah.getRange(3, 2, 1, 1).getFormula() : null;

    const snap = wb.save();

    return {
      workbookId: wb.getId(),
      formulaA1: fA1,
      formulaC4: fC4,
      snapshotA1: snap.sheets['Summary Hadiah']?.cellData?.[0]?.[0],
      snapshotC4: snap.sheets['Summary Hadiah']?.cellData?.[3]?.[2],
    };
  }, workbookData);

  console.log('Browser Results Phase 1:');
  console.log(JSON.stringify(browserResults, null, 2));

  await browser.close();
  console.log('=== PHASE 1 VERIFICATION COMPLETE ===');
}

testPhase1Verification().catch(console.error);
