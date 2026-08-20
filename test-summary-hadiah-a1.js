const puppeteer = require('puppeteer');

async function testSummaryHadiahA1() {
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  page.on('console', msg => console.log('[PAGE]', msg.text()));

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  const res = await page.evaluate(async () => {
    const api = window.univerAPI;

    // Clean old units
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
    oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

    const wbData = {
      id: `wb-test-summary-${Date.now()}`,
      name: 'Test Summary Workbook',
      sheetOrder: ['Summary', 'Summary Hadiah'],
      appVersion: '3.0.0-alpha',
      sheets: {
        'Summary': {
          id: 'Summary',
          name: 'Summary',
          type: 2,
          status: 1,
          rowCount: 50,
          columnCount: 20,
          cellData: {
            0: { 0: { v: 123, t: 2 } } // Summary!A1 = 123
          }
        },
        'Summary Hadiah': {
          id: 'Summary Hadiah',
          name: 'Summary Hadiah',
          type: 2,
          status: 0,
          rowCount: 50,
          columnCount: 20,
          cellData: {
            0: { 0: { f: '=Summary!A1' } } // Summary Hadiah!A1 = =Summary!A1 (NO v)
          }
        }
      }
    };

    console.log('Creating Univer Sheet...');
    const wb = api.createUniverSheet(wbData);

    console.log('Waiting 2 seconds for formula engine...');
    await new Promise(r => setTimeout(r, 2000));

    const sHadiah = wb.getSheetByName('Summary Hadiah');
    const range = sHadiah.getRange(0, 0, 1, 1);
    const value = range.getValue();
    const formula = range.getFormula ? range.getFormula() : undefined;

    const snapCell = wb.save().sheets['Summary Hadiah']?.cellData?.[0]?.[0];

    return {
      value,
      formula,
      snapCell
    };
  });

  console.log('SUMMARY HADIAH A1 RESULT:');
  console.log(JSON.stringify(res, null, 2));

  await browser.close();
}

testSummaryHadiahA1().catch(console.error);
