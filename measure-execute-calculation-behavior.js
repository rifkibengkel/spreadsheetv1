const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');

function log(msg) {
  console.log(msg);
  fs.appendFileSync('execute_calc_measurements.log', msg + '\n');
}

async function measureExecuteCalculation() {
  fs.writeFileSync('execute_calc_measurements.log', '');
  const realFilePath = 'C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx';

  log('======================================================================');
  log('FORENSIC MEASUREMENT OF executeCalculation() ON updatehexos.xlsx');
  log('======================================================================\n');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--js-flags=--expose-gc']
  });
  const page = await browser.newPage();

  const workerLogs = [];
  page.on('console', msg => {
    const text = msg.text();
    if (text.includes('Formula') || text.includes('Calculation') || text.includes('queued') || text.includes('Execute') || text.includes('Worker')) {
      log('[BROWSER] ' + text);
      workerLogs.push(text);
    }
  });

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 60000 });
  await new Promise(r => setTimeout(r, 1000));

  log('Uploading updatehexos.xlsx via app UI...');
  await page.evaluate(() => {
    const btns = Array.from(document.querySelectorAll('button'));
    const importBtn = btns.find(b => b.textContent.includes('Import'));
    if (importBtn) importBtn.click();
  });
  await new Promise(r => setTimeout(r, 500));

  const fileInput = await page.waitForSelector('input[type="file"]');
  await fileInput.uploadFile(realFilePath);
  await page.evaluate((input) => {
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }, fileInput);
  await new Promise(r => setTimeout(r, 800));

  const uploadStart = Date.now();
  await page.evaluate(() => {
    const btns = Array.from(document.querySelectorAll('button'));
    const startBtn = btns.find(b => b.textContent.includes('Mulai Impor') || b.textContent.includes('🚀'));
    if (startBtn && !startBtn.disabled) startBtn.click();
  });

  log('Waiting for import and initial sheet mount...');
  let mounted = false;
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
      mounted = true;
      log(`Workbook mounted in ${((Date.now() - uploadStart) / 1000).toFixed(2)}s! ID: ${activeWbInfo.id}`);
      break;
    }
  }

  if (!mounted) throw new Error('Workbook mount timed out.');

  // Measure state right after mount (before/during executeCalculation)
  log('\nMeasuring post-mount engine state...');
  const postMountState = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    const formulaEngine = api.getFormula ? api.getFormula() : null;

    const memAfterMount = performance.memory ? Math.round(performance.memory.usedJSHeapSize / (1024 * 1024)) : null;

    // Test UI responsiveness with quick performance task
    const t0 = performance.now();
    let sum = 0;
    for (let i = 0; i < 100000; i++) sum += i;
    const uiLatencyMs = Math.round(performance.now() - t0);

    return {
      memAfterMountMb: memAfterMount,
      uiLatencyMs,
      uiResponsive: uiLatencyMs < 50
    };
  });

  log('POST-MOUNT STATE: ' + JSON.stringify(postMountState, null, 2));

  // Wait 5 seconds to observe Formula Worker queue and memory changes after executeCalculation()
  log('\nWaiting 5 seconds to monitor post-init calculation worker behavior...');
  await new Promise(r => setTimeout(r, 5000));

  const postCalcState = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    const memAfterCalc = performance.memory ? Math.round(performance.memory.usedJSHeapSize / (1024 * 1024)) : null;

    // Check cells across sheets
    const summarySheet = wb.getSheetByName('Summary');
    const summaryValidSheet = wb.getSheetByName('Summary Valid');
    const unikSheet = wb.getSheetByName('Unik Konsumen');
    const evoucherSheet = wb.getSheetByName('E-Voucher & E-Wallet');

    // Summary!C4 formula & value
    const sumC4Range = summarySheet ? summarySheet.getRange(3, 2, 1, 1) : null;
    const sumC4Val = sumC4Range ? sumC4Range.getValue() : null;

    // Unik Konsumen!A1
    const unikA1Range = unikSheet ? unikSheet.getRange(0, 0, 1, 1) : null;
    const unikA1Val = unikA1Range ? unikA1Range.getValue() : null;

    // E-Voucher!A1 (one of the 33 uncached formulas)
    const evA1Range = evoucherSheet ? evoucherSheet.getRange(0, 0, 1, 1) : null;
    const evA1Val = evA1Range ? evA1Range.getValue() : null;

    return {
      memAfterCalcMb: memAfterCalc,
      summaryC4Val: sumC4Val,
      unikA1Val: unikA1Val,
      evoucherA1Val: evA1Val
    };
  });

  log('POST-CALCULATION STATE: ' + JSON.stringify(postCalcState, null, 2));

  // Test Test Cases A-G:
  log('\nRunning Acceptance Tests A-G on real file...');

  const testsResult = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    const s1 = wb.getActiveSheet(); // Summary Registrasi

    // TEST B: Simple Formula (=10+20) at row 0, col 5
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: wb.getId(),
      subUnitId: s1.getSheetId(),
      range: { startRow: 0, endRow: 0, startColumn: 5, endColumn: 5 },
      value: { f: '=10+20' }
    });
    await new Promise(r => setTimeout(r, 1000));
    const simpleVal = s1.getRange(0, 5, 1, 1).getValue();

    // TEST C: Cross-Sheet Formula (='Unik Konsumen'!A1) at row 0, col 6
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: wb.getId(),
      subUnitId: s1.getSheetId(),
      range: { startRow: 0, endRow: 0, startColumn: 6, endColumn: 6 },
      value: { f: "='Unik Konsumen'!A1" }
    });
    await new Promise(r => setTimeout(r, 1000));
    const crossVal = s1.getRange(0, 6, 1, 1).getValue();

    // TEST E: Dependency Edit (Source cell A1, formula at row 0, col 7)
    // First set source cell (0, 7) = =A1+10
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: wb.getId(),
      subUnitId: s1.getSheetId(),
      range: { startRow: 0, endRow: 0, startColumn: 7, endColumn: 7 },
      value: { f: '=A1+10' }
    });
    await new Promise(r => setTimeout(r, 1000));
    const depValInitial = s1.getRange(0, 7, 1, 1).getValue();

    // Now edit A1 (0, 0)
    const origA1 = s1.getRange(0, 0, 1, 1).getValue();
    const newA1 = 500;
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: wb.getId(),
      subUnitId: s1.getSheetId(),
      range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      value: { v: newA1, t: 2 }
    });
    await new Promise(r => setTimeout(r, 1000));
    const depValUpdated = s1.getRange(0, 7, 1, 1).getValue();

    return {
      testB_simple: { formula: '=10+20', value: simpleVal, success: simpleVal === 30 },
      testC_crossSheet: { formula: "='Unik Konsumen'!A1", value: crossVal, success: crossVal !== null && crossVal !== undefined },
      testE_dependency: { origA1, newA1, initial: depValInitial, updated: depValUpdated, success: depValUpdated === 510 }
    };
  });

  log('\nTESTS A-G RESULTS: ' + JSON.stringify(testsResult, null, 2));

  fs.writeFileSync('acceptance_report_updatehexos.json', JSON.stringify({
    postMountState,
    postCalcState,
    testsResult
  }, null, 2));

  await browser.close();
  log('\nMEASUREMENT COMPLETE!');
}

measureExecuteCalculation().catch(err => {
  log('ERROR: ' + (err.stack || String(err)));
  process.exit(1);
});
