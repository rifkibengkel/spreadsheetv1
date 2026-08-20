const fs = require('fs');
const puppeteer = require('puppeteer');

async function testUIViaPicker() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  console.log('==================================================');
  console.log('TESTING FILE PICKER IMPORT VIA REAL BROWSER UI');
  console.log('File:', filePath);
  console.log('==================================================\n');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  const page = await browser.newPage();
  await page.setViewport({ width: 1400, height: 900 });

  page.on('console', msg => console.log('BROWSER LOG:', msg.text()));

  console.log('Navigating to http://localhost:3000...');
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  await new Promise(r => setTimeout(r, 1000));

  console.log('Clicking Import button in UI...');
  await page.evaluate(() => {
    const btns = Array.from(document.querySelectorAll('button'));
    const importBtn = btns.find(b => b.textContent.includes('Import'));
    if (importBtn) importBtn.click();
  });
  await new Promise(r => setTimeout(r, 800));

  console.log('Locating file input in modal...');
  const fileInput = await page.waitForSelector('input[type="file"]', { timeout: 5000 });
  
  console.log('Uploading real Excel file report barakat.xlsx to file input...');
  await fileInput.uploadFile(filePath);
  await new Promise(r => setTimeout(r, 800));

  console.log('Clicking Mulai Impor button...');
  await page.evaluate(() => {
    const btns = Array.from(document.querySelectorAll('button'));
    const startBtn = btns.find(b => b.textContent.includes('Mulai Impor') || b.textContent.includes('🚀'));
    if (startBtn) startBtn.click();
  });

  console.log('Waiting for Web Worker import completion...');
  await page.waitForFunction(() => {
    const modalHeading = document.querySelector('h2');
    return !modalHeading || !document.body.innerText.includes('Memproses');
  }, { timeout: 60000 });

  await new Promise(r => setTimeout(r, 3000));

  console.log('Import finished! Taking UI screenshot...');
  await page.screenshot({ path: 'real_picker_import_ui.png' });

  // Evaluate UI formula calculation results directly in browser
  const uiResults = await page.evaluate(async () => {
    const api = window.univerAPI;
    const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    const summarySheet = activeWb.getSheetByName('Summary');
    const hadiahSheet = activeWb.getSheetByName('Summary Hadiah');

    return {
      Summary_I4: summarySheet ? summarySheet.getRange('I4').getValue() : null,
      Summary_I5: summarySheet ? summarySheet.getRange('I5').getValue() : null,
      Summary_I6: summarySheet ? summarySheet.getRange('I6').getValue() : null,
      Hadiah_C4: hadiahSheet ? hadiahSheet.getRange('C4').getValue() : null,
    };
  });

  console.log('\n--- UI FORMULA RESULTS AFTER REAL FILE PICKER IMPORT ---');
  console.log(JSON.stringify(uiResults, null, 2));

  await browser.close();
}

testUIViaPicker().catch(console.error);
