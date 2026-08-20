const puppeteer = require('puppeteer');

async function testRangeGet() {
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  await new Promise(r => setTimeout(r, 1000));

  const res = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.createUniverSheet({
      id: 'wb-test-getrange',
      name: 'Test',
      sheetOrder: ['s1'],
      sheets: { s1: { id: 's1', name: 'Sheet1', type: 2, status: 1, rowCount: 50, columnCount: 10, cellData: {} } }
    });
    const sheet = wb.getActiveSheet();

    let stringRangeVal = null;
    let numberRangeVal = null;

    try {
      sheet.getRange('A1').setValue(123);
      stringRangeVal = sheet.getRange('A1').getValue();
    } catch(e) {
      stringRangeVal = e.message;
    }

    try {
      numberRangeVal = sheet.getRange(0, 0).getValue();
    } catch(e) {
      numberRangeVal = e.message;
    }

    return { stringRangeVal, numberRangeVal };
  });

  console.log('RANGE GET RESULT:', JSON.stringify(res, null, 2));
  await browser.close();
}

testRangeGet().catch(console.error);
