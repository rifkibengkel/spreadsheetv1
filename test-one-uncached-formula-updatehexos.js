const fs = require('fs');
const puppeteer = require('puppeteer');

const EXCEL_PATH = 'C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx';

(async () => {
  console.log('=== Starting Test on ONE Known Formula in updatehexos.xlsx ===');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=4096']
  });

  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  // Open import modal and upload file
  const buttons = await page.$$('button');
  let importBtn = null;
  for (const b of buttons) {
    const text = await page.evaluate(el => el.textContent, b);
    if (text && text.includes('Import File')) {
      importBtn = b;
      break;
    }
  }
  await importBtn.click();
  await page.waitForSelector('input[type="file"]', { visible: false });

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
  await startBtn.click();

  console.log('Waiting for updatehexos.xlsx to mount...');

  // Wait for workbook to mount with NO_CALCULATION
  const targetTestResult = await page.evaluate(async () => {
    const api = window.univerAPI;

    let attempts = 0;
    while (attempts < 120) {
      const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : null;
      if (wb) {
        const sheets = wb.getSheets ? wb.getSheets() : [];
        if (sheets.length > 1) break;
      }
      await new Promise(r => setTimeout(r, 1000));
      attempts++;
    }

    const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : null;
    if (!activeWb) return { error: 'Workbook not mounted' };

    const unitId = activeWb.getId ? activeWb.getId() : activeWb.getUnitId();

    // Find ONE formula cell in sheet "Summary Registrasi" (row 3, col 2: =COUNTIF('Unik Konsumen'!$D$2:$D$89768,B4))
    const summarySheet = activeWb.getSheetByName('Summary Registrasi');
    if (!summarySheet) return { error: 'Summary Registrasi sheet not found' };

    const targetSheetId = summarySheet.getSheetId();
    const row = 3;
    const col = 2;

    const cellRange = summarySheet.getRange(row, col, 1, 1);
    const initialValue = cellRange.getValue(); // e.g. cached value 433

    // Record memory before targeted calculation
    let memBefore = null;
    if (window.performance && window.performance.memory) {
      memBefore = (window.performance.memory.usedJSHeapSize / 1024 / 1024).toFixed(2) + ' MB';
    }

    const startTime = Date.now();

    // Execute targeted native calculation mutation for ONLY this ONE range!
    const triggerCmdId = 'formula.mutation.set-trigger-formula-calculation-start';
    await api.executeCommand(triggerCmdId, {
      forceCalculation: false,
      dirtyRanges: [
        {
          unitId,
          sheetId: targetSheetId,
          range: { startRow: row, endRow: row, startColumn: col, endColumn: col }
        }
      ]
    }, { onlyLocal: true });

    // Wait for calculation pass to settle
    await new Promise(r => setTimeout(r, 1500));

    const durationMs = Date.now() - startTime;

    const finalValue = cellRange.getValue();

    let memAfter = null;
    if (window.performance && window.performance.memory) {
      memAfter = (window.performance.memory.usedJSHeapSize / 1024 / 1024).toFixed(2) + ' MB';
    }

    return {
      unitId,
      sheetName: 'Summary Registrasi',
      targetSheetId,
      cellLocation: `R${row+1}C${col+1}`,
      initialValue,
      finalValue,
      durationMs,
      memBefore,
      memAfter
    };
  });

  console.log('Targeted Single Formula Test Result:', JSON.stringify(targetTestResult, null, 2));

  fs.writeFileSync('single-formula-updatehexos-test.json', JSON.stringify(targetTestResult, null, 2));
  await browser.close();
})();
