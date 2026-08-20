const fs = require('fs');
const puppeteer = require('puppeteer');

async function inspectFacade() {
  try {
    const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
    const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(120000);
    page.setDefaultTimeout(120000);
    
    console.log('Navigating to http://localhost:3000...');
    await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 120000 });
    await page.waitForFunction('window.univerAPI !== undefined', { timeout: 120000 });
    await new Promise(r => setTimeout(r, 1000));

    console.log('Opening import modal...');
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const importBtn = btns.find(b => b.textContent.includes('Import'));
      if (importBtn) importBtn.click();
    });
    await new Promise(r => setTimeout(r, 1000));

    console.log('Uploading file...');
    const fileInput = await page.waitForSelector('input[type="file"]');
    await fileInput.uploadFile(filePath);
    await new Promise(r => setTimeout(r, 1000));

    console.log('Clicking start import...');
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const startBtn = btns.find(b => b.textContent.includes('Mulai Impor') || b.textContent.includes('🚀'));
      if (startBtn) startBtn.click();
    });

    console.log('Waiting 50 seconds for worker import of 28.7 MB file completion...');
    await new Promise(r => setTimeout(r, 50000));

    const res = await page.evaluate(() => {
      const api = window.univerAPI;
      const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : null;
      return {
        hasWb: !!wb,
        sheetCount: wb ? (wb.getSheets ? wb.getSheets().length : 0) : 0,
        sheetNames: wb ? wb.getSheets().map(s => s.getSheetName()) : []
      };
    });

    console.log('FACADE ACTIVE WORKBOOK INSPECTION AFTER 50S:', JSON.stringify(res, null, 2));
    fs.writeFileSync('inspect.log', JSON.stringify(res, null, 2));
    await browser.close();
  } catch(err) {
    console.error('INSPECT ERROR:', err);
    fs.writeFileSync('inspect.log', err.stack || String(err));
  }
}

inspectFacade();
