const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({ headless: 'new' });
  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  const res = await page.evaluate(async () => {
    const api = window.univerAPI;
    const formulaEngine = api.getFormula ? api.getFormula() : null;

    if (formulaEngine && formulaEngine.setInitialFormulaComputing) {
      formulaEngine.setInitialFormulaComputing(2);
    }

    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
    oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });

    const wbData = {
      id: `wb-cross-${Date.now()}`,
      name: 'CrossTest',
      sheetOrder: ['sheet1', 'sheet2'],
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
            '0': { '0': { v: 10, t: 2 } }, // A1 = 10
            '1': { '0': { v: 20, t: 2 } }, // A2 = 20
            '2': { '0': { f: '=A1+A2' } }  // A3 = 30
          }
        },
        'sheet2': {
          id: 'sheet2',
          name: 'Sheet2',
          type: 2,
          status: 0,
          rowCount: 10,
          columnCount: 10,
          cellData: {
            '0': { '1': { v: 50, t: 2 } },                 // B1 = 50
            '0': { '2': { f: '=Sheet1!A1+Sheet2!B1' } },  // C1
            '0': { '3': { f: '=Sheet1!A1+B1' } }          // D1
          }
        }
      }
    };

    api.createUniverSheet(wbData);
    const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    
    // Execute calculation
    formulaEngine.executeCalculation();
    await new Promise(r => setTimeout(r, 2000));

    const sheet1 = activeWb.getSheetByName('Sheet1');
    const sheet2 = activeWb.getSheetByName('Sheet2');

    return {
      sheet1_A3: sheet1.getRange(2, 0, 1, 1).getValue(),
      sheet2_C1_name_ref: sheet2.getRange(0, 2, 1, 1).getValue(),
      sheet2_D1_simple_cross: sheet2.getRange(0, 3, 1, 1).getValue(),
      sheet2_saved: activeWb.save().sheets['sheet2']?.cellData
    };
  });

  console.log('Cross sheet results:', res);
  await browser.close();
})();
