const puppeteer = require('puppeteer');
const fs = require('fs');

(async () => {
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox', '--disable-setuid-sandbox'] });
  const page = await browser.newPage();
  
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  
  // Wait to settle
  await new Promise(r => setTimeout(r, 4000));

  const results = await page.evaluate(async () => {
    const api = window.univerAPI;
    const out = {};
    
    const importedData = {
      id: 'imported-wb',
      sheetOrder: ['sheet1', 'sheet2'],
      name: 'ImportedTest',
      sheets: {
        'sheet1': {
          id: 'sheet1',
          name: 'Sheet1',
          type: 2, // SheetType.GRID
          rowCount: 100,
          columnCount: 30,
          cellData: {
            '0': { '0': { v: 100, t: 2 } }, // A1
            '1': { '0': { v: 200, t: 2 } }, // A2
            '2': { '0': { f: '=A1+A2', v: 300, t: 2 } } // A3
          }
        },
        'sheet2': {
          id: 'sheet2',
          name: 'Sheet2',
          type: 2,
          rowCount: 100,
          columnCount: 30,
          cellData: {
            '0': { '0': { f: '=Sheet1!A1', v: 100, t: 2 } } // A1 references Sheet1
          }
        }
      }
    };
    
    // Switch completely to the new workbook
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
    oldWbs.forEach(wb => api.disposeUnit(wb.getId()));
    
    api.createUniverSheet(importedData);
    
    // Wait for Univer to set active workbook
    await new Promise(r => setTimeout(r, 2000));
    const wb = api.getAllWorkbooks ? api.getAllWorkbooks()[0] : api.getActiveWorkbook();
    if (!wb) return { error: "Failed to get workbook" };
    
    let a3Data = wb.getActiveSheet().getCellData()['2']?.['0'];
    let a3Val = wb.getActiveSheet().getRange(2, 0, 1, 1).getValue();
    
    out.initial_import = {
      raw_v_after_import: a3Data?.v,
      api_value_after_import: a3Val,
      ui_probably_visible: (a3Data?.v === 300 || a3Data?.v === '300' || a3Val === 300)
    };
    
    // CRITICAL TEST: Dependency Change
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: wb.getId(),
      subUnitId: 'sheet1',
      range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      value: { v: 500, t: 2 } // Change A1 to 500
    });
    
    await new Promise(r => setTimeout(r, 2000));
    
    a3Data = wb.getActiveSheet().getCellData()['2']?.['0'];
    a3Val = wb.getActiveSheet().getRange(2, 0, 1, 1).getValue();
    
    out.after_dependency_change = {
      raw_v: a3Data?.v,
      api_value: a3Val
    };
    
    // CRITICAL TEST 2: Manual Edit of Formula
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: wb.getId(),
      subUnitId: 'sheet1',
      range: { startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 },
      value: { f: '=A1*A2' }
    });
    
    await new Promise(r => setTimeout(r, 2000));
    a3Data = wb.getActiveSheet().getCellData()['2']?.['0'];
    a3Val = wb.getActiveSheet().getRange(2, 0, 1, 1).getValue();
    
    out.after_manual_formula_edit = {
      raw_v: a3Data?.v,
      api_value: a3Val
    };
    
    return out;
  });

  console.log(JSON.stringify(results, null, 2));
  await browser.close();
  process.exit(0);
})();
