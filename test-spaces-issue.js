const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

async function testSpacesIssue() {
  console.log('=== Creating external XLSX with spaces in sheet names ===');
  const wb = new ExcelJS.Workbook();
  const s1 = wb.addWorksheet('Entries data');
  s1.getCell('A1').value = 100;
  s1.getCell('A2').value = 200;

  const s2 = wb.addWorksheet('Summary Hadiah');
  s2.getCell('A1').value = { formula: "'Entries data'!A1" };
  s2.getCell('A2').value = { formula: "SUM('Entries data'!A1:A2)" };

  const testFilePath = path.join(__dirname, 'test-external-spaces.xlsx');
  await wb.xlsx.writeFile(testFilePath);
  console.log('Saved to:', testFilePath);

  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();

  page.on('console', msg => console.log('[BROWSER LOG]', msg.text()));

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Uploading file...');
  await page.evaluate(() => {
    const btns = Array.from(document.querySelectorAll('button'));
    const importBtn = btns.find(b => b.textContent.includes('Import'));
    if (importBtn) importBtn.click();
  });
  await new Promise(r => setTimeout(r, 1000));

  const fileInput = await page.waitForSelector('input[type="file"]');
  await fileInput.uploadFile(testFilePath);
  await new Promise(r => setTimeout(r, 1000));

  await page.evaluate(() => {
    const btns = Array.from(document.querySelectorAll('button'));
    const startBtn = btns.find(b => b.textContent.includes('Mulai Impor') || b.textContent.includes('🚀'));
    if (startBtn) startBtn.click();
  });

  console.log('Waiting 5 seconds for import...');
  await new Promise(r => setTimeout(r, 5000));

  const result = await page.evaluate(async () => {
    try {
      const api = window.univerAPI;
      const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
      if (!wb) return { error: 'No active workbook found' };

      const sheets = wb.getSheets ? wb.getSheets().map(s => s.getSheetName()) : [];
      const s2 = wb.getSheetByName('Summary Hadiah');
      if (!s2) return { error: 'Summary Hadiah not found', sheets };

      const rA1 = s2.getRange(0, 0, 1, 1);
      const rA2 = s2.getRange(1, 0, 1, 1);

      const snapshot = wb.save();
      const s2Snap = snapshot.sheets['Summary Hadiah'] || snapshot.sheets[s2.getSheetId()];

      return {
        sheets,
        A1: {
          value: rA1.getValue(),
          cellData: s2Snap?.cellData?.[0]?.[0]
        },
        A2: {
          value: rA2.getValue(),
          cellData: s2Snap?.cellData?.[1]?.[0]
        }
      };
    } catch (e) {
      return { error: e.message, stack: e.stack };
    }
  });

  fs.writeFileSync('spaces-test-results.json', JSON.stringify(result, null, 2));
  console.log('SPACES TEST RESULTS WRITTEN TO spaces-test-results.json');
  console.log(JSON.stringify(result, null, 2));

  await browser.close();
}

testSpacesIssue().catch(console.error);
