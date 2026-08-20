const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({ headless: 'new' });
  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');
  await new Promise(r => setTimeout(r, 1000));
  
  const results = await page.evaluate(async () => {
    const api = window.univerAPI;
    const testCases = [
      { name: 'Case 1: f, v, t', cell: { f: '=A1+A2', v: 300, t: 2 } },
      { name: 'Case 2: f only', cell: { f: '=A1+A2' } },
      { name: 'Case 3: f with v null', cell: { f: '=A1+A2', v: null } },
      { name: 'Case 4: f without equals', cell: { f: 'A1+A2' } },
      { name: 'Case 5: cross-sheet matching sheetId and sheetName', cell: { f: '=Sheet1!A1' }, sheetId: 'Sheet1', sheetName: 'Sheet1' },
      { name: 'Case 6: cross-sheet slugified sheetId mismatch (sheetId=sheet1, sheetName=Sheet1)', cell: { f: '=Sheet1!A1' }, sheetId: 'sheet1', sheetName: 'Sheet1' },
    ];
    
    const out = {};
    for (let tc of testCases) {
      const targetSheetId = tc.sheetId || 'sheet1';
      const targetSheetName = tc.sheetName || 'Sheet1';
      const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
      oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });
      
      const cellCopy = JSON.parse(JSON.stringify(tc.cell));
      const wbData = {
        id: 'wb-' + Date.now(),
        name: 'Wb',
        sheetOrder: [targetSheetId],
        appVersion: '3.0.0-alpha',
        sheets: {
          [targetSheetId]: {
            id: targetSheetId,
            name: targetSheetName,
            type: 2,
            status: 1,
            rowCount: 10, columnCount: 10,
            cellData: {
              '0': { '0': { v: 100, t: 2 } },
              '1': { '0': { v: 200, t: 2 } },
              '2': { '0': cellCopy }
            }
          }
        }
      };
      
      api.createUniverSheet(wbData);
      await new Promise(r => setTimeout(r, 2000));
      const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
      const sheet = activeWb.getActiveSheet();
      
      const snap = activeWb.save();
      const cellInModel = snap.sheets[targetSheetId]?.cellData?.[2]?.[0];
      const displayedVal = sheet.getRange(2, 0, 1, 1).getValue();
      
      out[tc.name] = {
        inputCell: tc.cell,
        sheetId: targetSheetId,
        sheetName: targetSheetName,
        modelCell: cellInModel,
        displayedValue: displayedVal
      };
    }
    return out;
  });
  
  fs.writeFileSync('test-cases-out.json', JSON.stringify(results, null, 2));
  console.log('Saved test-cases-out.json');
  await browser.close();
})();
