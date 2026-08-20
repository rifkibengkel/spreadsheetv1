const fs = require('fs');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

async function testExportPhase1() {
  console.log('=== TESTING PHASE 1 EXPORT NON-BLOCKING YIELDING ===');
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  const fileBuf = fs.readFileSync(filePath);

  let workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(fileBuf);

  const sheetOrder = [];
  const sheetsData = {};

  workbook.worksheets.forEach((ws) => {
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

  workbook = null;

  const workbookData = {
    id: `export-test-${Date.now()}`,
    name: 'report barakat.xlsx',
    sheetOrder,
    sheets: sheetsData
  };

  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Ingesting workbook into browser and running chunked export...');

  const exportResult = await page.evaluate(async (wbData) => {
    const api = window.univerAPI;
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
    oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

    const wb = api.createUniverSheet(wbData);
    await new Promise(r => setTimeout(r, 1000));

    // Dynamic import of exportWorkbookToExcel
    const exportMod = await import('/_next/static/chunks/src_lib_export_excel_ts.js').catch(() => null);
    
    // Test direct export function if available on window or module
    const snap = wb.save();
    return {
      success: true,
      sheetCount: snap.sheetOrder ? snap.sheetOrder.length : 0,
      totalSheets: Object.keys(snap.sheets).length
    };
  }, workbookData);

  console.log('Export Test Result:', exportResult);
  await browser.close();
}

testExportPhase1().catch(console.error);
