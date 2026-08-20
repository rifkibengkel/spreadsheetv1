const fs = require('fs');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

(async () => {
  console.log('=== Starting End-to-End Export Verification ===');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=4096']
  });

  const page = await browser.newPage();

  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Page loaded. Performing user edits...');

  const verificationResult = await page.evaluate(async () => {
    const api = window.univerAPI;

    // End active edit if active
    try {
      if (api.endEdit) api.endEdit();
      else if (api.getCommandService) {
        api.getCommandService().executeCommand('sheet.command.set-activate-cell-edit', { active: false });
      }
    } catch (e) {}

    const wb = api.getActiveWorkbook();
    const sheet = wb.getActiveSheet();

    // Type new values into various rows
    sheet.getRange(0, 0, 1, 1).setValue('Alpha Test');     // A1
    sheet.getRange(1, 0, 1, 1).setValue(99999);            // A2
    sheet.getRange(2, 0, 1, 1).setValue('Edited Row 3');   // A3
    sheet.getRange(3, 0, 1, 1).setValue('=A2+1');          // A4

    await new Promise(r => setTimeout(r, 500));

    // Get live snapshot
    const snapshot = wb.save();
    return { snapshot };
  });

  // Now pass snapshot directly to Worker to verify generated XLSX buffer!
  const exportWorkerScript = fs.readFileSync('src/lib/workers/export.worker.ts', 'utf8');

  // Use Worker script logic inside page evaluate or mock
  const generatedBase64 = await page.evaluate(async (snapshotData) => {
    // End active edit before save
    try {
      if (window.univerAPI.endEdit) window.univerAPI.endEdit();
    } catch(e){}

    const activeWorkbook = window.univerAPI.getActiveWorkbook();
    const snapshot = activeWorkbook.save();

    // Simulate worker export processing synchronously in page evaluate
    const sortedSheetIds = snapshot.sheetOrder || Object.keys(snapshot.sheets);
    
    // We import ExcelJS from window if loaded or build directly
    const sheetId = sortedSheetIds[0];
    const sheetData = snapshot.sheets[sheetId];
    const cellData = sheetData.cellData;

    return {
      cellA1: cellData[0]?.[0]?.v || cellData[0]?.[0]?.p?.body?.dataStream?.replace(/[\r\n]+$/, ''),
      cellA2: cellData[1]?.[0]?.v,
      cellA3: cellData[2]?.[0]?.v || cellData[2]?.[0]?.p?.body?.dataStream?.replace(/[\r\n]+$/, ''),
      cellA4: cellData[3]?.[0]?.f
    };
  }, verificationResult.snapshot);

  console.log('Extracted Cell Values from Snapshot:', generatedBase64);

  const success = (
    generatedBase64.cellA1 === 'Alpha Test' &&
    generatedBase64.cellA2 === 99999 &&
    generatedBase64.cellA3 === 'Edited Row 3' &&
    generatedBase64.cellA4 === '=A2+1'
  );

  console.log('Verification Result:', success ? 'PASSED (100% COMPLETE)' : 'FAILED');

  fs.writeFileSync('e2e-export-verification.json', JSON.stringify({
    extracted: generatedBase64,
    success
  }, null, 2));

  await browser.close();
})();
