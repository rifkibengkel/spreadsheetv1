const puppeteer = require('puppeteer');

async function test4Params() {
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  await new Promise(r => setTimeout(r, 1000));

  const res = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    const sheet = wb.getActiveSheet();

    let twoParamsRes = null;
    let fourParamsRes = null;

    try {
      sheet.getRange(0, 0).setValue(999);
      twoParamsRes = sheet.getRange(0, 0).getValue();
    } catch(e) {
      twoParamsRes = e.message;
    }

    try {
      fourParamsRes = sheet.getRange(0, 0, 1, 1).getValue();
    } catch(e) {
      fourParamsRes = e.message;
    }

    return { twoParamsRes, fourParamsRes };
  });

  console.log('PARAM TEST RESULT:', JSON.stringify(res, null, 2));
  await browser.close();
}

test4Params().catch(console.error);
