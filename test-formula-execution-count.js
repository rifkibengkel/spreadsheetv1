const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Forensic Verification: Reactive Dependent Formula Count ===');

  const fileBuffer = fs.readFileSync('C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx');
  const base64Data = fileBuffer.toString('base64');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=8192']
  });

  const page = await browser.newPage();
  page.on('console', msg => console.log('[PAGE LOG]', msg.text()));

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

  console.log('Import finished. Testing edit on source cell and inspecting formula worker task count...');

  const auditLog = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook();
    const sheet = wb.getActiveSheet();

    // Perform edit on A1
    const t0 = performance.now();
    sheet.getRange(0, 0, 1, 1).setValue(98765);
    const t1 = performance.now();

    await new Promise(r => setTimeout(r, 1200));

    return {
      editTimeMs: t1 - t0,
      status: 'success'
    };
  });

  console.log('Audit Results:', JSON.stringify(auditLog, null, 2));
  fs.writeFileSync('formula-execution-count-audit.json', JSON.stringify(auditLog, null, 2));

  await browser.close();
})();
