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
      { name: 'Numeric Result as Number (t=2, v=300)', cell: { f: '=A1+A2', v: 300, t: 2 } },
      { name: 'Numeric Result as String (t=1, v="300")', cell: { f: '=A1+A2', v: '300', t: 1 } },
      { name: 'Formula with String Type (t=1, no v)', cell: { f: '=A1+A2', t: 1 } },
      { name: 'Formula with undefined t and v', cell: { f: '=A1+A2' } },
      { name: 'Formula with t=2 and no v', cell: { f: '=A1+A2', t: 2 } }
    ];
    
    const out = {};
    for (let tc of testCases) {
      const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
      oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });
      
      const wbData = {
        id: 'wb-' + Date.now(),
        name: 'Wb',
        sheetOrder: ['sheet1'],
        appVersion: '3.0.0-alpha',
        sheets: {
          'sheet1': {
            id: 'sheet1', name: 'Sheet1', type: 2, status: 1, rowCount: 10, columnCount: 10,
            cellData: {
              '0': { '0': { v: 100, t: 2 } },
              '1': { '0': { v: 200, t: 2 } },
              '2': { '0': JSON.parse(JSON.stringify(tc.cell)) }
            }
          }
        }
      };
      
      api.createUniverSheet(wbData);
      await new Promise(r => setTimeout(r, 2000));
      const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
      const sheet = activeWb.getActiveSheet();
      
      const snap = activeWb.save();
      const cellInModel = snap.sheets['sheet1']?.cellData?.[2]?.[0];
      const displayedVal = sheet.getRange(2, 0, 1, 1).getValue();
      
      out[tc.name] = {
        inputCell: tc.cell,
        modelCell: cellInModel,
        displayedValue: displayedVal
      };
    }
    return out;
  });
  
  fs.writeFileSync('types-out.json', JSON.stringify(results, null, 2));
  console.log('Saved types-out.json');
  await browser.close();
})();
