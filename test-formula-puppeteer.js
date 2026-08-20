const puppeteer = require('puppeteer');

async function runTest() {
  console.log("🚀 Starting Puppeteer Test...");
  const browser = await puppeteer.launch({ headless: 'new' });
  const page = await browser.newPage();

const fs = require('fs');

  page.on('console', msg => {
    const text = msg.text();
    if (!text.includes('DOM element is missing') && !text.includes('React DevTools')) {
      fs.appendFileSync('test-error-log.txt', `[Console] ${text}\n`);
    }
  });

  page.on('pageerror', err => {
    fs.appendFileSync('test-error-log.txt', `[PageError] ${err.toString()}\n`);
  });

  console.log("🌐 Navigating to http://localhost:3000 ...");
  await page.goto('http://localhost:3000', { waitUntil: 'networkidle2' });

  console.log("⏳ Waiting for UI to load completely...");
  await page.waitForTimeout(5000); 

  console.log("✅ Wait complete! Injecting formula test...");

  const result = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook();
    const sheet = wb.getActiveSheet();
    const sheetId = sheet.getSheetId();
    const wbId = wb.getId();

    function setValue(row, col, value, type = 2) {
      const SetRangeValuesCommand = window.univerAPI._commands?.get('sheet.command.set-range-values') || 
                                    window.univerAPI.getCommandService?.().getCommand('sheet.command.set-range-values');
      
      return api.executeCommand('sheet.command.set-range-values', {
        unitId: wbId,
        subUnitId: sheetId,
        range: { startRow: row, endRow: row, startColumn: col, endColumn: col },
        value: typeof value === 'string' && value.startsWith('=') 
               ? { f: value } 
               : { v: value, t: type }
      });
    }

    console.log("📝 Typing 100 into A1...");
    setValue(0, 0, 100);
    
    console.log("📝 Typing 200 into A2...");
    setValue(1, 0, 200);

    console.log("📝 Typing =A1+A2 into A3...");
    setValue(2, 0, '=A1+A2');

    // Wait a bit to let formula calculate synchronously or via worker
    await new Promise(r => setTimeout(r, 1000));

    console.log("📝 Typing =1+1 into B1...");
    setValue(0, 1, '=1+1');

    await new Promise(r => setTimeout(r, 1000));

    const a3 = sheet.getRange(2, 0, 1, 1).getValue();
    const b1 = sheet.getRange(0, 1, 1, 1).getValue();

    return { a3, b1 };
  });

  console.log("\n📊 --- TEST RESULTS ---");
  console.log("Result A3 (=A1+A2):", JSON.stringify(result.a3, null, 2));
  console.log("Result B1 (=1+1)  :", JSON.stringify(result.b1, null, 2));

  await browser.close();
}

runTest().catch(console.error);
