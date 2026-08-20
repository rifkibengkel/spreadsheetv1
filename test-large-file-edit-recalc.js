const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Measuring Large File (updatehexos.xlsx) Edit Recalculation Performance ===');

  const fileBuffer = fs.readFileSync('C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx');
  const base64Data = fileBuffer.toString('base64');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=8192']
  });

  const page = await browser.newPage();

  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Importing updatehexos.xlsx...');

  const importResult = await page.evaluate(async (b64) => {
    const binaryString = atob(b64);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const file = new File([blob], 'updatehexos.xlsx');

    const api = window.univerAPI;
    const startTime = performance.now();
    const result = await window.importExcelToWorkbookDataAsync(file);
    window.replaceUniverWorkbook(api, result.workbookData);

    const importTime = performance.now() - startTime;
    return { importTime };
  }, base64Data);

  console.log(`Import finished in ${(importResult.importTime / 1000).toFixed(2)}s. Testing cell edit recalculation...`);

  const editResult = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook();

    // Select sheet 'Daftar Hexos' or active sheet
    const sheet = wb.getActiveSheet();

    // Measure time for cell edit reactive calculation
    const t0 = performance.now();

    // Edit cell A1
    sheet.getRange(0, 0, 1, 1).setValue(9999);

    const t1 = performance.now();

    await new Promise(r => setTimeout(r, 1500));
    const t2 = performance.now();

    return {
      setValueTimeMs: t1 - t0,
      totalElapsedMs: t2 - t0
    };
  });

  console.log('Edit Recalculation Audit Result:', JSON.stringify(editResult, null, 2));
  fs.writeFileSync('large-edit-recalc-audit.json', JSON.stringify(editResult, null, 2));

  await browser.close();
})();
