const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

async function reproduceExactIssue() {
  try {
    console.log('=== Creating minimal external XLSX (no cached formula result) ===');
    const wb = new ExcelJS.Workbook();
    const s1 = wb.addWorksheet('Sheet1');
    s1.getCell('A1').value = 100;
    s1.getCell('A2').value = 200;

    const s2 = wb.addWorksheet('Sheet2');
    s2.getCell('A1').value = { formula: 'Sheet1!A1' };
    s2.getCell('A2').value = { formula: 'Sheet1!A2' };
    s2.getCell('A3').value = { formula: 'A1+A2' };

    const testFilePath = path.join(__dirname, 'test-external-minimal.xlsx');
    await wb.xlsx.writeFile(testFilePath);
    console.log('Saved external XLSX to:', testFilePath);

    const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
    const page = await browser.newPage();

    page.on('console', msg => console.log('[BROWSER]', msg.text()));
    page.on('pageerror', err => console.log('[PAGE ERROR]', err));

    await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction('window.univerAPI !== undefined');

    console.log('Opening import modal...');
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const importBtn = btns.find(b => b.textContent.includes('Import'));
      if (importBtn) importBtn.click();
    });
    await new Promise(r => setTimeout(r, 1000));

    console.log('Uploading file...');
    const fileInput = await page.waitForSelector('input[type="file"]');
    await fileInput.uploadFile(testFilePath);
    await new Promise(r => setTimeout(r, 1000));

    console.log('Clicking start import button...');
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const startBtn = btns.find(b => b.textContent.includes('Mulai Impor') || b.textContent.includes('🚀'));
      if (startBtn) startBtn.click();
    });

    console.log('Waiting 5 seconds for worker and Univer sheet replacement...');
    await new Promise(r => setTimeout(r, 5000));

    const result = await page.evaluate(async () => {
      try {
        const api = window.univerAPI;
        const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
        if (!wb) return { error: 'No active workbook found' };

        const sheets = wb.getSheets ? wb.getSheets().map(s => s.getSheetName()) : [];
        const s2 = wb.getSheetByName('Sheet2');
        if (!s2) return { error: 'Sheet2 not found in imported workbook', sheets };

        const rA1 = s2.getRange(0, 0, 1, 1);
        const rA2 = s2.getRange(1, 0, 1, 1);
        const rA3 = s2.getRange(2, 0, 1, 1);

        const snapshot = wb.save();
        const sheet2Snap = snapshot.sheets['Sheet2'] || snapshot.sheets[s2.getSheetId()];

        return {
          activeWorkbookId: wb.getId(),
          sheets,
          Sheet2_A1: {
            value: rA1.getValue(),
            cellData: sheet2Snap?.cellData?.[0]?.[0]
          },
          Sheet2_A2: {
            value: rA2.getValue(),
            cellData: sheet2Snap?.cellData?.[1]?.[0]
          },
          Sheet2_A3: {
            value: rA3.getValue(),
            cellData: sheet2Snap?.cellData?.[2]?.[0]
          }
        };
      } catch (e) {
        return { error: e.message, stack: e.stack };
      }
    });

    fs.writeFileSync('reproduce-results.json', JSON.stringify(result, null, 2));
    console.log('REPRODUCE TEST RESULTS WRITTEN TO reproduce-results.json');
    console.log(JSON.stringify(result, null, 2));

    await browser.close();
  } catch (outerErr) {
    console.error('OUTER ERR:', outerErr);
  }
}

reproduceExactIssue();
