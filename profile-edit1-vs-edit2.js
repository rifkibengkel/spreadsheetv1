const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Forensic Breakdown: Edit 1 (DAG Build) vs Edit 2 & 3 (Warm Tree) ===');

  const fileBuffer = fs.readFileSync('C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx');
  const base64Data = fileBuffer.toString('base64');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=8192']
  });

  const page = await browser.newPage();
  page.on('console', msg => {
    const text = msg.text();
    if (text.includes('[Formula Worker]') || text.includes('consumed') || text.includes('Stage')) {
      console.log(`[PAGE LOG ${performance.now().toFixed(2)}ms]`, text);
    }
  });

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

  console.log('Import finished. Measuring Edit 1 vs Edit 2 vs Edit 3...');

  const editPhases = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook();
    const activeSheet = wb.getActiveSheet();

    const phases = [];

    for (let editNum = 1; editNum <= 3; editNum++) {
      const tStart = performance.now();

      // Edit cell (0, 0)
      activeSheet.getRange(0, 0, 1, 1).setValue(10000 + editNum);
      const tMutation = performance.now();

      // Wait 15 seconds for worker calculation to settle
      await new Promise(r => setTimeout(r, 15000));
      const tEnd = performance.now();

      phases.push({
        editNum,
        mutationTimeMs: tMutation - tStart,
        totalTimeMs: tEnd - tStart
      });
    }

    return phases;
  });

  console.log('\n--- EDIT 1 vs EDIT 2 vs EDIT 3 RESULTS ---');
  console.log(JSON.stringify(editPhases, null, 2));

  fs.writeFileSync('edit-warmup-breakdown.json', JSON.stringify(editPhases, null, 2));

  await browser.close();
})();
