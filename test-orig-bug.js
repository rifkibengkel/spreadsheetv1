const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox', '--disable-setuid-sandbox'] });
  const page = await browser.newPage();
  
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  
  await new Promise(r => setTimeout(r, 4000));

  const results = await page.evaluate(async () => {
    const api = window.univerAPI;
    const out = {};
    
    // Simulate original faulty import exactly!
    const importedData = {
      id: 'imported-wb-orig',
      sheetOrder: ['sheet1'],
      name: 'OrigTest',
      sheets: {
        'sheet1': {
          id: 'sheet1',
          name: 'Sheet1',
          type: 2, // SheetType.GRID
          rowCount: 10,
          columnCount: 10,
          cellData: {
            '0': { '0': { v: 100, t: 2 } }, // A1
            '1': { '0': { v: 200, t: 2 } }, // A2
            '2': { '0': { f: '=A1+A2', v: 300, t: 2 } } // A3 (Formula + Cached value)
          }
        }
      }
    };
    
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
    oldWbs.forEach(w => { if (w) api.disposeUnit(w.getId()); });
    
    api.createUniverSheet(importedData);
    await new Promise(r => setTimeout(r, 2000));
    
    const wb = api.getActiveWorkbook();
    if (!wb) return { error: "no workbook" };
    const sheet = wb.getActiveSheet();
    if (!sheet) return { error: "no active sheet" };
    
    out.cellData_A3 = sheet.getCellData?.()?.['2']?.['0'] || sheet.getCellMatrix?.()?.getValue(2, 0);
    out.getValue_A3 = sheet.getRange(2, 0, 1, 1).getValue();
    
    // Try to change dependency
    api.executeCommand('sheet.command.set-range-values', {
      unitId: wb.getId(),
      subUnitId: sheet.getSheetId(),
      range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      value: { v: 500, t: 2 }
    });
    await new Promise(r => setTimeout(r, 2000));
    
    out.cellData_A3_after_dep_change = sheet.getCellData?.()?.['2']?.['0'] || sheet.getCellMatrix?.()?.getValue(2, 0);
    out.getValue_A3_after_dep_change = sheet.getRange(2, 0, 1, 1).getValue();
    
    return out;
  });

  console.log(JSON.stringify(results, null, 2));
  await browser.close();
  process.exit(0);
})();
