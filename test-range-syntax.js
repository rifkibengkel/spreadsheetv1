const puppeteer = require('puppeteer');

async function testRangeSyntax() {
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  await new Promise(r => setTimeout(r, 1000));

  const res = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.createUniverSheet({
      id: 'wb-test-syntax',
      name: 'Test',
      sheetOrder: ['s1', 's2'],
      sheets: {
        s1: { id: 's1', name: 'Sheet1', type: 2, status: 1, rowCount: 50, columnCount: 10, cellData: {} },
        s2: { id: 's2', name: 'Sheet2', type: 2, status: 0, rowCount: 50, columnCount: 10, cellData: { 0: { 0: { v: 999, t: 2 } } } }
      }
    });

    const s1 = wb.getSheetByName('Sheet1');

    const r1 = s1.getRange(0, 0, 1, 1);
    const r2 = s1.getRange(0, 1, 1, 1);
    const r3 = s1.getRange(0, 2, 1, 1);
    const r4 = s1.getRange(0, 3, 1, 1);

    r1.setValue(10);
    r2.setValue(20);
    
    // Formula via setValue string
    r3.setValue('=A1+B1');

    // Cross-sheet formula via setValue string
    r4.setValue("='Sheet2'!A1");

    await new Promise(r => setTimeout(r, 1500));

    return {
      A1: r1.getValue(),
      B1: r2.getValue(),
      C1: r3.getValue(),
      D1: r4.getValue()
    };
  });

  console.log('SYNTAX TEST RESULT:', JSON.stringify(res, null, 2));
  await browser.close();
}

testRangeSyntax().catch(console.error);
