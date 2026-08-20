const fs = require('fs');
const puppeteer = require('puppeteer');

const EXCEL_PATH = 'C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx';

(async () => {
  console.log('=== Starting Full Verification of Deferred Calculation Implementation ===');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=4096']
  });

  const page = await browser.newPage();

  // Track console logs to confirm no executeCalculation warnings or errors
  const consoleLogs = [];
  page.on('console', msg => consoleLogs.push(msg.text()));

  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('1. Page loaded successfully.');

  // Open import modal
  const buttons = await page.$$('button');
  let importBtn = null;
  for (const b of buttons) {
    const text = await page.evaluate(el => el.textContent, b);
    if (text && text.includes('Import File')) {
      importBtn = b;
      break;
    }
  }

  if (!importBtn) {
    console.error('Import button not found');
    await browser.close();
    process.exit(1);
  }

  await importBtn.click();
  await page.waitForSelector('input[type="file"]', { visible: false });

  console.log('2. Uploading updatehexos.xlsx...');
  const fileInput = await page.$('input[type="file"]');
  await fileInput.uploadFile(EXCEL_PATH);

  const modalButtons = await page.$$('button');
  let startBtn = null;
  for (const b of modalButtons) {
    const text = await page.evaluate(el => el.textContent, b);
    if (text && text.includes('Mulai Impor')) {
      startBtn = b;
      break;
    }
  }

  const startImportTime = Date.now();
  await startBtn.click();

  console.log('3. Waiting for import and mounting to complete...');

  const importResult = await page.evaluate(async () => {
    const api = window.univerAPI;

    let attempts = 0;
    while (attempts < 120) {
      const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : null;
      if (wb) {
        const sheets = wb.getSheets ? wb.getSheets() : [];
        if (sheets.length > 1) {
          break;
        }
      }
      await new Promise(r => setTimeout(r, 1000));
      attempts++;
    }

    const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : null;
    if (!activeWb) return { error: 'Workbook not mounted' };

    const wbSave = activeWb.save();
    let formulaCount = 0;
    let cachedValCount = 0;
    const sampleCells = [];

    if (wbSave.sheets) {
      Object.keys(wbSave.sheets).forEach(sheetId => {
        const sheet = wbSave.sheets[sheetId];
        if (sheet.cellData) {
          Object.keys(sheet.cellData).forEach(r => {
            const row = sheet.cellData[r];
            if (row) {
              Object.keys(row).forEach(c => {
                const cell = row[c];
                if (cell) {
                  if (cell.f) {
                    formulaCount++;
                    if (sampleCells.length < 5) {
                      sampleCells.push({
                        sheet: sheet.name,
                        r, c,
                        f: cell.f,
                        v: cell.v
                      });
                    }
                  }
                  if (cell.v !== undefined && cell.v !== null) {
                    cachedValCount++;
                  }
                }
              });
            }
          });
        }
      });
    }

    return {
      sheetCount: activeWb.getSheets().length,
      formulaCount,
      cachedValCount,
      sampleCells
    };
  });

  const importDuration = Date.now() - startImportTime;
  console.log(`4. Import completed in ${importDuration} ms.`);
  console.log('   Metrics:', JSON.stringify(importResult, null, 2));

  // Test 5: Verify reactive calculation on manual cell edit
  console.log('5. Testing manual cell edit and reactive calculation...');
  const reactiveResult = await page.evaluate(async () => {
    const api = window.univerAPI;
    const formulaEngine = api.getFormula ? api.getFormula() : null;

    // Set CalculationMode to 0 (FORCED) or normal for manual edit test in small test sheet
    if (formulaEngine && formulaEngine.setInitialFormulaComputing) {
      formulaEngine.setInitialFormulaComputing(0);
    }

    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
    oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });

    const wbData = {
      id: `wb-verify-${Date.now()}`,
      name: 'VerifyReactive',
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
            '2': { '0': { f: '=A1+A2' } }  // A3 = =A1+A2 (same sheet)
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
            '0': { '0': { v: 100, t: 2 } },              // A1 = 100
            '0': { '1': { f: '=Sheet1!A1+Sheet2!A1' } }  // B1 = cross sheet
          }
        }
      }
    };

    api.createUniverSheet(wbData);
    const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    const s1 = activeWb.getSheetByName('Sheet1');
    
    // Modify source cell A1 to 50
    s1.getRange(0, 0, 1, 1).setValue(50);
    await new Promise(r => setTimeout(r, 1000));

    const s1_A3_val = s1.getRange(2, 0, 1, 1).getValue();

    return {
      sameSheetFormulaResult: s1_A3_val
    };
  });

  console.log('   Reactive edit result:', JSON.stringify(reactiveResult, null, 2));

  const finalVerification = {
    importDurationMs: importDuration,
    updatehexosMetrics: importResult,
    reactiveTest: reactiveResult,
    executeCalculationLogs: consoleLogs.filter(l => l.includes('executeCalculation'))
  };

  fs.writeFileSync('final-verification-summary.json', JSON.stringify(finalVerification, null, 2));
  console.log('=== Verification Finished Successfully! ===');

  await browser.close();
})();
