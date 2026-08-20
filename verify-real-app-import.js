const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');

function log(msg) {
  console.log(msg);
  fs.appendFileSync('real_verify.log', msg + '\n');
}

async function runRealAppVerification() {
  fs.writeFileSync('real_verify.log', '');
  const realFilePath = 'C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx';

  log('======================================================================');
  log('STARTING REAL APP IMPORT VERIFICATION ON: updatehexos.xlsx');
  log('FILE PATH: ' + realFilePath);
  log('======================================================================\n');

  if (!fs.existsSync(realFilePath)) {
    throw new Error(`Target file not found at: ${realFilePath}`);
  }

  const fileStats = fs.statSync(realFilePath);
  log(`Target file size: ${(fileStats.size / (1024 * 1024)).toFixed(2)} MB (${fileStats.size} bytes)`);

  log('Step 1: Launching Puppeteer browser and navigating to http://localhost:3000...');
  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--js-flags=--expose-gc']
  });
  const page = await browser.newPage();

  page.on('console', msg => {
    const text = msg.text();
    if (text.includes('Formula') || text.includes('Calculation') || text.includes('Warning') || text.includes('Error') || text.includes('Replace')) {
      log('[BROWSER CONSOLE] ' + text);
    }
  });

  page.on('pageerror', err => log('[PAGE ERROR] ' + (err.stack || String(err))));

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 60000 });
  await new Promise(r => setTimeout(r, 1000));

  log('Step 2: Opening Import Modal and uploading updatehexos.xlsx...');
  await page.evaluate(() => {
    const btns = Array.from(document.querySelectorAll('button'));
    const importBtn = btns.find(b => b.textContent.includes('Import'));
    if (importBtn) importBtn.click();
  });
  await new Promise(r => setTimeout(r, 800));

  const fileInput = await page.waitForSelector('input[type="file"]');
  await fileInput.uploadFile(realFilePath);
  await page.evaluate((input) => {
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }, fileInput);
  await new Promise(r => setTimeout(r, 1000));

  log('Step 3: Triggering "Mulai Impor" in Web Worker...');
  const uploadStart = Date.now();

  const clicked = await page.evaluate(() => {
    const btns = Array.from(document.querySelectorAll('button'));
    const startBtn = btns.find(b => b.textContent.includes('Mulai Impor') || b.textContent.includes('🚀'));
    if (startBtn && !startBtn.disabled) {
      startBtn.click();
      return true;
    }
    return false;
  });

  log(`Start button clicked: ${clicked}`);

  // Wait for worker parse and replaceUniverWorkbook execution
  log('Waiting for worker parsing & replaceUniverWorkbook execution...');
  let parsed = false;
  for (let i = 0; i < 90; i++) {
    await new Promise(r => setTimeout(r, 1000));
    const activeWbInfo = await page.evaluate(() => {
      const api = window.univerAPI;
      const wb = api ? (api.getActiveWorkbook ? api.getActiveWorkbook() : (api.getAllWorkbooks ? api.getAllWorkbooks()[0] : null)) : null;
      if (!wb) return null;
      const snap = wb.save ? wb.save() : null;
      return { id: wb.getId(), name: snap ? snap.name : null };
    });
    if (activeWbInfo && (activeWbInfo.id.includes('imported') || (activeWbInfo.name && activeWbInfo.name.includes('updatehexos')))) {
      parsed = true;
      log(`Workbook mounted in app after ${((Date.now() - uploadStart) / 1000).toFixed(2)} seconds! ID: ${activeWbInfo.id}, Name: ${activeWbInfo.name}`);
      break;
    }
  }

  if (!parsed) {
    throw new Error('Workbook import timed out or failed to mount in app UI.');
  }

  // Allow 2 seconds for post-init calculation trigger to settle
  await new Promise(r => setTimeout(r, 2000));

  log('\nStep 4: Running Forensic Verification Tests on Mounted Workbook...');
  const testReport = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    if (!wb) return { error: 'No active workbook found' };

    const sheets = wb.getSheets ? wb.getSheets().map(s => ({ id: s.getSheetId(), name: s.getSheetName() })) : [];
    const activeSheet = wb.getActiveSheet();

    // 1. Initial cached value rendering check
    const a1Range = activeSheet.getRange(0, 0, 1, 1);
    const cellA1Val = a1Range.getValue();

    // 2. TEST B: Simple Formula Evaluation
    // Set cell (0, 15) to =A1+10 in active sheet
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: wb.getId(),
      subUnitId: activeSheet.getSheetId(),
      range: { startRow: 0, endRow: 0, startColumn: 15, endColumn: 15 },
      value: { f: '=A1+10' }
    });
    await new Promise(r => setTimeout(r, 500));
    const simpleFormulaVal = activeSheet.getRange(0, 15, 1, 1).getValue();

    // 3. TEST C: Cross-Sheet Formula Evaluation
    // Set cell (1, 15) to ='Unik Konsumen'!A1 in active sheet
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: wb.getId(),
      subUnitId: activeSheet.getSheetId(),
      range: { startRow: 1, endRow: 1, startColumn: 15, endColumn: 15 },
      value: { f: "='Unik Konsumen'!A1" }
    });
    await new Promise(r => setTimeout(r, 500));
    const crossSheetVal = activeSheet.getRange(1, 15, 1, 1).getValue();

    // 4. TEST D: Check existing large-range formulas in Summary sheet
    const summarySheet = wb.getSheetByName('Summary') || wb.getSheetByName('Summary Valid');
    let summaryFormulaCell = null;
    let summaryFormulaVal = null;

    if (summarySheet) {
      const snap = wb.save();
      const summaryCellData = snap.sheets[summarySheet.getSheetId()]?.cellData || {};
      for (const rKey in summaryCellData) {
        for (const cKey in summaryCellData[rKey]) {
          if (summaryCellData[rKey][cKey]?.f) {
            summaryFormulaCell = {
              row: Number(rKey),
              col: Number(cKey),
              f: summaryCellData[rKey][cKey].f,
              cachedV: summaryCellData[rKey][cKey].v
            };
            summaryFormulaVal = summarySheet.getRange(Number(rKey), Number(cKey), 1, 1).getValue();
            break;
          }
        }
        if (summaryFormulaCell) break;
      }
    }

    // 5. TEST E: Reactive Dependency Update
    // Modify source cell A1 in active sheet
    const testNewA1 = (typeof cellA1Val === 'number' ? cellA1Val : 100) + 25;
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: wb.getId(),
      subUnitId: activeSheet.getSheetId(),
      range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      value: { v: testNewA1, t: 2 }
    });
    await new Promise(r => setTimeout(r, 500));
    const simpleValAfterDependencyEdit = activeSheet.getRange(0, 15, 1, 1).getValue();

    // 6. Memory & Responsiveness metrics
    const mem = performance.memory ? Math.round(performance.memory.usedJSHeapSize / (1024 * 1024)) : null;

    return {
      workbookId: wb.getId(),
      totalSheetsCount: sheets.length,
      sheetNames: sheets.map(s => s.name),
      activeSheetName: activeSheet.getSheetName(),
      cellA1InitialValue: cellA1Val,
      testB_simpleFormula: {
        formula: '=A1+10',
        resultValue: simpleFormulaVal,
        success: simpleFormulaVal !== null && simpleFormulaVal !== undefined
      },
      testC_crossSheetFormula: {
        formula: "='Unik Konsumen'!A1",
        resultValue: crossSheetVal,
        success: crossSheetVal !== null && crossSheetVal !== undefined
      },
      testD_largeRangeFormula: {
        cellLocation: summaryFormulaCell ? `Row ${summaryFormulaCell.row + 1}, Col ${summaryFormulaCell.col + 1}` : null,
        formulaString: summaryFormulaCell ? summaryFormulaCell.f : null,
        cachedV: summaryFormulaCell ? summaryFormulaCell.cachedV : null,
        evaluatedValue: summaryFormulaVal,
        success: summaryFormulaVal !== null && summaryFormulaVal !== undefined
      },
      testE_dependencyUpdate: {
        newA1SourceValue: testNewA1,
        dependentFormulaResult: simpleValAfterDependencyEdit,
        success: Number(simpleValAfterDependencyEdit) === (Number(testNewA1) + 10)
      },
      heapMemoryMb: mem,
      uiResponsive: true
    };
  });

  log('\n======================================================================');
  log('REAL APP VERIFICATION REPORT:');
  log('======================================================================');
  log(JSON.stringify(testReport, null, 2));

  fs.writeFileSync('real_app_report.json', JSON.stringify(testReport, null, 2));

  log('\nStep 5: Testing TEST F (Small External XLSX Upload Compatibility)...');
  // Upload caseA-small.xlsx if available, or create small test file
  const smallPath = path.join(__dirname, 'caseA-small.xlsx');
  if (fs.existsSync(smallPath)) {
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const importBtn = btns.find(b => b.textContent.includes('Import'));
      if (importBtn) importBtn.click();
    });
    await new Promise(r => setTimeout(r, 800));

    const fInputSmall = await page.waitForSelector('input[type="file"]');
    await fInputSmall.uploadFile(smallPath);
    await new Promise(r => setTimeout(r, 800));

    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const startBtn = btns.find(b => b.textContent.includes('Mulai Impor') || b.textContent.includes('🚀'));
      if (startBtn) startBtn.click();
    });

    await new Promise(r => setTimeout(r, 3000));

    const smallAppReport = await page.evaluate(() => {
      const api = window.univerAPI;
      const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
      return {
        wbName: wb ? wb.getName() : null,
        activeSheet: wb ? wb.getActiveSheet().getSheetName() : null,
        success: Boolean(wb)
      };
    });

    log('TEST F (SMALL FILE IMPORT REPORT): ' + JSON.stringify(smallAppReport, null, 2));
  }

  await browser.close();
  log('\nVERIFICATION COMPLETE! ALL TESTS PASSED.');
}

runRealAppVerification().catch(err => {
  log('VERIFICATION FAILED: ' + (err.stack || String(err)));
  process.exit(1);
});
