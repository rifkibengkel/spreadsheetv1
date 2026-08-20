const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({ headless: 'new' });
  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');
  await new Promise(r => setTimeout(r, 500));

  const results = await page.evaluate(async () => {
    const api = window.univerAPI;
    const formulaEngine = api.getFormula ? api.getFormula() : null;

    // Set NO_CALCULATION
    if (formulaEngine && formulaEngine.setInitialFormulaComputing) {
      formulaEngine.setInitialFormulaComputing(2); // NO_CALCULATION
    }

    // Clear old workbooks
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
    oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });

    const wbData = {
      id: `wb-step3-${Date.now()}`,
      name: 'Step3Wb',
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
            '0': { '0': { v: 10, t: 2 } },  // A1 = 10
            '1': { '0': { v: 20, t: 2 } },  // A2 = 20
            '2': { '0': { f: '=A1+A2' } }   // A3 = =A1+A2 (Same-sheet formula)
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
            '0': { '2': { f: '=Sheet1!A1+Sheet2!B1' } } // C1 = =Sheet1!A1+Sheet2!B1 (Cross-sheet formula)
          }
        }
      }
    };

    // 1. Create workbook with NO_CALCULATION
    api.createUniverSheet(wbData);
    const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    const sheet1 = activeWb.getSheetByName('Sheet1');
    const sheet2 = activeWb.getSheetByName('Sheet2');

    // Check A: Initial values immediately after createUniverSheet
    const initialSameSheetVal = sheet1.getRange(2, 0, 1, 1).getValue();
    const initialCrossSheetVal = sheet2.getRange(0, 2, 1, 1).getValue();

    // Wait 1 second to ensure no automatic calculation happens
    await new Promise(r => setTimeout(r, 1000));

    const settledBeforeCalcSameSheetVal = sheet1.getRange(2, 0, 1, 1).getValue();
    const settledBeforeCalcCrossSheetVal = sheet2.getRange(0, 2, 1, 1).getValue();

    // Check B: Formula definitions preservation
    const sheet1Save = activeWb.save().sheets['sheet1'];
    const sheet2Save = activeWb.save().sheets['sheet2'];
    const sameSheetFormulaDef = sheet1Save?.cellData?.[2]?.[0]?.f;
    const crossSheetFormulaDef = sheet2Save?.cellData?.[0]?.[2]?.f;

    // Check D: Trigger native executeCalculation
    let calcError = null;
    try {
      if (formulaEngine && typeof formulaEngine.executeCalculation === 'function') {
        formulaEngine.executeCalculation();
      }
    } catch (e) {
      calcError = e.message;
    }

    // Wait for calculation to finish
    if (formulaEngine && formulaEngine.onCalculationResultApplied) {
      try {
        await formulaEngine.onCalculationResultApplied(5000);
      } catch (e) {
        // timeout or error
      }
    } else {
      await new Promise(r => setTimeout(r, 1500));
    }

    // Check E & F: Post-calculation values
    const postCalcSameSheetVal = sheet1.getRange(2, 0, 1, 1).getValue();
    const postCalcCrossSheetVal = sheet2.getRange(0, 2, 1, 1).getValue();

    return {
      initialSameSheetVal,
      initialCrossSheetVal,
      settledBeforeCalcSameSheetVal,
      settledBeforeCalcCrossSheetVal,
      sameSheetFormulaDef,
      crossSheetFormulaDef,
      calcError,
      postCalcSameSheetVal,
      postCalcCrossSheetVal
    };
  });

  fs.writeFileSync('step3-results.json', JSON.stringify(results, null, 2));
  console.log('Step 3 results:', results);
  await browser.close();
})();
