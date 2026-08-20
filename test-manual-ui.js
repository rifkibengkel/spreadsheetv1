const puppeteer = require('puppeteer');

async function testManualFormulas() {
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 60000 });
  await new Promise(r => setTimeout(r, 1000));

  const res = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook();
    const sheet = wb.getActiveSheet();

    // Set A1 = 100, A2 = 200, A3 = =A1+A2
    const cellA1 = sheet.getRange(0, 0, 1, 1);
    const cellA2 = sheet.getRange(1, 0, 1, 1);
    const cellA3 = sheet.getRange(2, 0, 1, 1);

    cellA1.setValue(100);
    cellA2.setValue(200);
    cellA3.setValue('=A1+A2');

    await new Promise(r => setTimeout(r, 1500));

    const valA3 = cellA3.getValue();

    return { valA3 };
  });

  console.log('MANUAL FORMULA TEST IN ACTIVE UI:', JSON.stringify(res, null, 2));
  await browser.close();
}

testManualFormulas().catch(console.error);
