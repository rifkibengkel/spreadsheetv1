const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Measuring Reactive Formula Recalculation Performance ===');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=4096']
  });

  const page = await browser.newPage();
  page.on('console', msg => console.log('[PAGE LOG]', msg.text()));

  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Page loaded. Testing reactive edit recalculation timing...');

  const timings = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook();
    const sheet1 = wb.getActiveSheet();

    // 1. Setup small same-sheet formula: A1=100, A2=200, A3==A1+A2
    sheet1.getRange(0, 0, 1, 1).setValue(100);
    sheet1.getRange(1, 0, 1, 1).setValue(200);
    sheet1.getRange(2, 0, 1, 1).setValue('=A1+A2');

    await new Promise(r => setTimeout(r, 500));

    // Measure same-sheet edit recalculation
    const t0 = performance.now();
    sheet1.getRange(0, 0, 1, 1).setValue(150); // Change A1 from 100 to 150

    // Poll until A3 becomes 350 or timeout
    let t1 = 0;
    let valA3 = null;

    for (let i = 0; i < 50; i++) {
      await new Promise(r => setTimeout(r, 20));
      valA3 = sheet1.getRange(2, 0, 1, 1).getValue();
      if (valA3 === 350 || (typeof valA3 === 'object' && valA3?.v === 350)) {
        t1 = performance.now();
        break;
      }
    }

    const sameSheetDelayMs = t1 > 0 ? (t1 - t0) : -1;

    // 2. Setup cross-sheet formula: Sheet1!A1 -> Sheet2!A1==Sheet1!A1
    let sheet2 = wb.getSheetByName('Sheet2');
    if (!sheet2) {
      sheet2 = wb.create('Sheet2', 10, 10);
    }
    sheet2.getRange(0, 0, 1, 1).setValue('=Sheet1!A1');

    await new Promise(r => setTimeout(r, 500));

    const t2 = performance.now();
    sheet1.getRange(0, 0, 1, 1).setValue(700); // Change Sheet1!A1 to 700

    let t3 = 0;
    let valSheet2A1 = null;

    for (let i = 0; i < 50; i++) {
      await new Promise(r => setTimeout(r, 20));
      valSheet2A1 = sheet2.getRange(0, 0, 1, 1).getValue();
      if (valSheet2A1 === 700 || (typeof valSheet2A1 === 'object' && valSheet2A1?.v === 700)) {
        t3 = performance.now();
        break;
      }
    }

    const crossSheetDelayMs = t3 > 0 ? (t3 - t2) : -1;

    return {
      sameSheetDelayMs,
      valA3,
      crossSheetDelayMs,
      valSheet2A1
    };
  });

  console.log('Recalculation Timing Result:', JSON.stringify(timings, null, 2));
  fs.writeFileSync('recalc-timing-audit.json', JSON.stringify(timings, null, 2));

  await browser.close();
})();
