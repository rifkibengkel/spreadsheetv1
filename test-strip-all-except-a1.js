const fs = require('fs');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

async function testStripAllExceptA1() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  const buf = fs.readFileSync(filePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);

  const sheetOrder = [];
  const sheetsData = {};

  workbook.worksheets.forEach((ws, sIdx) => {
    const sheetId = ws.name;
    sheetOrder.push(sheetId);

    sheetsData[sheetId] = {
      id: sheetId,
      name: ws.name,
      type: 2,
      status: sIdx === 0 ? 1 : 0,
      rowCount: 100,
      columnCount: 30,
      cellData: {}
    };
  });

  // Only put Summary A1 = 123 and Summary Hadiah A1 = =Summary!A1
  sheetsData['Summary'].cellData = { 0: { 0: { v: 123, t: 2 } } };
  sheetsData['Summary Hadiah'].cellData = { 0: { 0: { f: '=Summary!A1' } } };

  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  page.on('console', msg => console.log('[BROWSER]', msg.text()));

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  const res = await page.evaluate(async (data) => {
    const api = window.univerAPI;
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
    oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

    const wb = api.createUniverSheet(data);
    await new Promise(r => setTimeout(r, 2000));

    const sHadiah = wb.getSheetByName('Summary Hadiah');
    const range = sHadiah.getRange(0, 0, 1, 1);
    return {
      value: range.getValue(),
      formula: range.getFormula ? range.getFormula() : null,
      cellData: wb.save().sheets['Summary Hadiah']?.cellData?.[0]?.[0]
    };
  }, { id: `tb-all5-${Date.now()}`, name: 'All 5 Sheets minimal', sheetOrder, appVersion: '3.0.0-alpha', sheets: sheetsData });

  console.log('ALL 5 SHEETS MINIMAL A1 RESULT:');
  console.log(JSON.stringify(res, null, 2));

  await browser.close();
}

testStripAllExceptA1().catch(console.error);
