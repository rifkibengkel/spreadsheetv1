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
    const formulaEngine = api.getFormula ? api.getFormula() : null;

    const modes = [
      { name: 'FORCED (0)', mode: 0 },
      { name: 'WHEN_EMPTY (1)', mode: 1 },
      { name: 'NO_CALCULATION (2)', mode: 2 }
    ];

    const out = {};

    for (const m of modes) {
      if (formulaEngine && formulaEngine.setInitialFormulaComputing) {
        formulaEngine.setInitialFormulaComputing(m.mode);
      }

      const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
      oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });

      // Create workbook with formula WITHOUT cached value v
      const wbData = {
        id: `wb-mode-${m.mode}-${Date.now()}`,
        name: `WbMode${m.mode}`,
        sheetOrder: ['sheet1'],
        appVersion: '3.0.0-alpha',
        sheets: {
          'sheet1': {
            id: 'sheet1',
            name: 'Sheet1',
            type: 2,
            status: 1,
            rowCount: 10,
            columnCount: 10,
            cellData: {
              '0': { '0': { v: 100, t: 2 } }, // A1 = 100
              '1': { '0': { v: 200, t: 2 } }, // A2 = 200
              '2': { '0': { f: '=A1+A2' } }   // A3 = =A1+A2 (NO v, NO t)
            }
          }
        }
      };

      api.createUniverSheet(wbData);
      const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];

      // Immediate read
      const immValue = activeWb.getActiveSheet().getRange(2, 0, 1, 1).getValue();

      // Wait 1.5 seconds for calculation event
      await new Promise(r => setTimeout(r, 1500));

      const setVal = activeWb.getActiveSheet().getRange(2, 0, 1, 1).getValue();
      const cellInModel = activeWb.save().sheets['sheet1']?.cellData?.[2]?.[0];

      out[m.name] = {
        modeSetting: m.mode,
        immediateValue: immValue,
        settledValue: setVal,
        cellInModel
      };
    }

    return out;
  });

  fs.writeFileSync('initial-computing-out.json', JSON.stringify(results, null, 2));
  console.log('Saved initial-computing-out.json');
  await browser.close();
})();
