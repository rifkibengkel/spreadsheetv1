const puppeteer = require('puppeteer');

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
    
    // 1. Check BEFORE import
    let wb = api.getActiveWorkbook();
    if (!wb) {
       // if no active, try getAll
       const all = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
       wb = all[0];
    }
    
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: wb.getId(),
      subUnitId: wb.getActiveSheet().getSheetId(),
      range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      value: { v: 100, t: 2 }
    });
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: wb.getId(),
      subUnitId: wb.getActiveSheet().getSheetId(),
      range: { startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 },
      value: { v: 200, t: 2 }
    });
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: wb.getId(),
      subUnitId: wb.getActiveSheet().getSheetId(),
      range: { startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 },
      value: { f: '=A1+A2' }
    });
    
    await new Promise(r => setTimeout(r, 2000));
    out.before_import_A3 = wb.getActiveSheet().getRange(2, 0, 1, 1).getValue(); // should be 300
    
    // Check injector state before import
    // Facade uses internal Injector! 
    const injector = api._injector || (window.univer && window.univer._injector);
    out.injector_exists_before = !!injector;
    
    // 2. Perform Excel Import (Mimic)
    const importedData = {
      id: 'imported-wb',
      sheetOrder: ['sheet1'],
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
            // NO FORMULAS IN EXCEL DATA
          }
        }
      }
    };
    
    // IMPORTANT: dispose previously
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
    oldWbs.forEach(w => { if (w) api.disposeUnit(w.getId()); });
    
    api.createUniverSheet(importedData);
    
    await new Promise(r => setTimeout(r, 2000));
    const all2 = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
    const newWb = all2.find(w => w.getId() === 'imported-wb') || api.getActiveWorkbook();
    
    // 3. Test manually typing a formula after import
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: newWb.getId(),
      subUnitId: newWb.getActiveSheet().getSheetId(),
      range: { startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 },
      value: { f: '=A1+A2' }
    });
    
    await new Promise(r => setTimeout(r, 2000));
    out.after_import_manual_A3 = newWb.getActiveSheet().getRange(2, 0, 1, 1).getValue(); // ??
    out.after_import_manual_A3_f = newWb.getActiveSheet().getCellData()['2']?.['0']?.f; // should be '=A1+A2'
    
    // 4. Test adding a completely new sheet after import
    // Unfortunately Facade api might not easily add a sheet but let's try
    await api.executeCommand('sheet.command.insert-sheet', {
      unitId: newWb.getId(),
      index: 1,
      sheet: { id: 'sheet2', name: 'Sheet2', type: 2 }
    });
    await new Promise(r => setTimeout(r, 500));
    
    // change active sheet
    await api.executeCommand('sheet.command.set-worksheet-active', {
      unitId: newWb.getId(),
      subUnitId: 'sheet2'
    });
    
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: newWb.getId(),
      subUnitId: 'sheet2',
      range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      value: { v: 100, t: 2 }
    });
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: newWb.getId(),
      subUnitId: 'sheet2',
      range: { startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 },
      value: { v: 200, t: 2 }
    });
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: newWb.getId(),
      subUnitId: 'sheet2',
      range: { startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 },
      value: { f: '=A1+A2' }
    });
    
    await new Promise(r => setTimeout(r, 2000));
    out.after_import_new_sheet_A3 = newWb.getActiveSheet().getRange(2, 0, 1, 1).getValue(); // ??

    return out;
  });

  console.log(JSON.stringify(results, null, 2));
  await browser.close();
  process.exit(0);
})();
