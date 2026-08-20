const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Measuring Sequential Edit Recalculation Performance ===');

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

  await page.evaluate(async (b64) => {
    const binaryString = atob(b64);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const file = new File([blob], 'updatehexos.xlsx');

    const api = window.univerAPI;
    const result = await window.importExcelToWorkbookDataAsync(file);
    window.replaceUniverWorkbook(api, result.workbookData);
  }, base64Data);

  console.log('Import finished. Measuring 3 sequential edits...');

  const editTimings = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook();
    const sheet = wb.getActiveSheet();

    const results = [];

    for (let editIndex = 1; editIndex <= 3; editIndex++) {
      const t0 = performance.now();
      sheet.getRange(0, 0, 1, 1).setValue(1000 + editIndex);
      const t1 = performance.now();

      await new Promise(r => setTimeout(r, 800));
      const t2 = performance.now();

      results.push({
        editIndex,
        mutationTimeMs: t1 - t0,
        totalTimeMs: t2 - t0
      });
    }

    return results;
  });

  console.log('Sequential Edit Audit Results:', JSON.stringify(editTimings, null, 2));
  fs.writeFileSync('sequential-edits-audit.json', JSON.stringify(editTimings, null, 2));

  await browser.close();
})();
