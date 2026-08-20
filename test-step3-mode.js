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

    // 1. Set mode 2 (NO_CALCULATION)
    if (formulaEngine && formulaEngine.setInitialFormulaComputing) {
      formulaEngine.setInitialFormulaComputing(2);
    }

    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
    oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });

    const wbData = {
      id: `wb-test-${Date.now()}`,
      name: 'TestWb',
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
            '0': { '1': { v: 50, t: 2 } },            // B1 = 50
            '0': { '2': { f: '=Sheet1!A1+Sheet2!B1' } } // C1 = 60
          }
        }
      }
    };

    api.createUniverSheet(wbData);

    const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    const s1 = activeWb.getSheetByName('Sheet1');
    const s2 = activeWb.getSheetByName('Sheet2');

    const valBeforeCalc = {
      s1_A3: s1.getRange(2, 0, 1, 1).getValue(),
      s2_C1: s2.getRange(0, 2, 1, 1).getValue()
    };

    // Test Call 1: executeCalculation() while mode is still 2
    formulaEngine.executeCalculation();
    await new Promise(r => setTimeout(r, 1000));
    const valAfterExecCalcMode2 = {
      s1_A3: s1.getRange(2, 0, 1, 1).getValue(),
      s2_C1: s2.getRange(0, 2, 1, 1).getValue()
    };

    // Test Call 2: Change setInitialFormulaComputing to 0 (FORCED) or 1 (WHEN_EMPTY) then executeCalculation()
    if (formulaEngine.setInitialFormulaComputing) {
      formulaEngine.setInitialFormulaComputing(0); // FORCED
    }
    formulaEngine.executeCalculation();
    await new Promise(r => setTimeout(r, 1500));
    const valAfterModeResetAndExecCalc = {
      s1_A3: s1.getRange(2, 0, 1, 1).getValue(),
      s2_C1: s2.getRange(0, 2, 1, 1).getValue()
    };

    return {
      valBeforeCalc,
      valAfterExecCalcMode2,
      valAfterModeResetAndExecCalc,
      sheet1_saved: activeWb.save().sheets['sheet1']?.cellData?.[2]?.[0],
      sheet2_saved: activeWb.save().sheets['sheet2']?.cellData?.[0]?.[2]
    };
  });

  fs.writeFileSync('step3-mode-test.json', JSON.stringify(res, null, 2));
  console.log('Mode test result:', JSON.stringify(res, null, 2));
  await browser.close();
})();
