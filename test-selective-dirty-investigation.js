const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Starting Selective Dirty Investigation (Synthetic Workbook) ===');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=4096']
  });

  const page = await browser.newPage();

  page.on('console', msg => console.log('[PAGE CONSOLE]', msg.text()));

  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('1. Page loaded, window.univerAPI ready.');

  const result = await page.evaluate(async () => {
    const api = window.univerAPI;
    const formulaEngine = api.getFormula ? api.getFormula() : null;

    // A. Set CalculationMode to NO_CALCULATION (2)
    if (formulaEngine && formulaEngine.setInitialFormulaComputing) {
      formulaEngine.setInitialFormulaComputing(2);
    }

    // Dispose old workbooks
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
    oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });

    const unitId = `wb-tiny-${Date.now()}`;
    const wbData = {
      id: unitId,
      name: 'TinyWorkbook',
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
            '0': { '0': { v: 100, t: 2 } }, // A1 = 100
            '1': { '0': { v: 200, t: 2 } }, // A2 = 200
            '2': { '0': { f: '=A1+A2' } }   // A3 = =A1+A2 (no v)
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
            '0': { '0': { f: '=Sheet1!A3' } } // A1 = =Sheet1!A3 (no v)
          }
        }
      }
    };

    api.createUniverSheet(wbData);

    const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    const s1 = activeWb.getSheetByName('Sheet1');
    const s2 = activeWb.getSheetByName('Sheet2');

    // Monitor command executions
    const capturedCommands = [];
    const commandService = formulaEngine?._commandService || api;
    
    let disposable = null;
    if (commandService && commandService.onCommandExecuted) {
      disposable = commandService.onCommandExecuted((cmd) => {
        if (cmd && cmd.id) {
          capturedCommands.push({ id: cmd.id });
        }
      });
    }

    // B. Check initial values
    const s1_A3_init_v = s1.getRange(2, 0, 1, 1).getValue(); // A3
    const s2_A1_init_v = s2.getRange(0, 0, 1, 1).getValue(); // A1

    // Question 1: Test what commands occur when user edits cell A1 (setValue 100)
    const editCommandsCapturedBefore = capturedCommands.length;
    s1.getRange(0, 0, 1, 1).setValue(100);
    await new Promise(r => setTimeout(r, 800));
    const userEditCommands = capturedCommands.slice(editCommandsCapturedBefore);

    // C. Test triggering native dirty calculation for target range ONLY: Sheet1!A1:A3
    const triggerStartCmdId = 'formula.mutation.set-trigger-formula-calculation-start';
    
    let targetedTriggerSuccess = false;
    let targetedTriggerError = null;

    try {
      await api.executeCommand(triggerStartCmdId, {
        forceCalculation: false,
        dirtyRanges: [
          {
            unitId: unitId,
            sheetId: 'sheet1',
            range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 }
          }
        ]
      }, { onlyLocal: true });
      targetedTriggerSuccess = true;
    } catch (err) {
      targetedTriggerError = err.message || String(err);
    }

    // Wait for calculation to settle
    await new Promise(r => setTimeout(r, 1200));

    // D & E: Check updated values
    const s1_A3_after_v = s1.getRange(2, 0, 1, 1).getValue();
    const s2_A1_after_v = s2.getRange(0, 0, 1, 1).getValue();

    if (disposable && disposable.dispose) disposable.dispose();

    return {
      s1_A3_init_v,
      s2_A1_init_v,
      userEditCommands,
      targetedTriggerSuccess,
      targetedTriggerError,
      s1_A3_after_v,
      s2_A1_after_v,
      allCapturedCommandsSample: capturedCommands.slice(0, 10)
    };
  });

  const jsonOutput = JSON.stringify(result, null, 2);
  console.log('Test Results:', jsonOutput);
  fs.writeFileSync('investigation-synthetic-results.json', jsonOutput);

  await browser.close();
})();
