const puppeteer = require('puppeteer');

async function inspectFWorksheet() {
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  await new Promise(r => setTimeout(r, 1000));

  const res = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    const sheet = wb.getActiveSheet ? wb.getActiveSheet() : wb.getSheets()[0];

    const methods = [];
    for (let prop in sheet) {
      methods.push(prop);
    }

    return {
      sheetName: sheet.getSheetName ? sheet.getSheetName() : 'NO_METHOD',
      prototypeMethods: Object.getOwnPropertyNames(Object.getPrototypeOf(sheet)),
      properties: methods
    };
  });

  console.log('FWORKSHEET INSPECTION:', JSON.stringify(res, null, 2));
  await browser.close();
}

inspectFWorksheet().catch(console.error);
