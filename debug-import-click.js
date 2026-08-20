const fs = require('fs');
const puppeteer = require('puppeteer');

async function debugImportClick() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';

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

  console.log('Uploading real file report barakat.xlsx...');
  const fileInput = await page.waitForSelector('input[type="file"]');
  await fileInput.uploadFile(filePath);
  await new Promise(r => setTimeout(r, 1000));

  console.log('Clicking start import button...');
  await page.evaluate(() => {
    const btns = Array.from(document.querySelectorAll('button'));
    const startBtn = btns.find(b => b.textContent.includes('Mulai Impor') || b.textContent.includes('🚀'));
    if (startBtn) startBtn.click();
  });

  console.log('Waiting 15 seconds for import of report barakat.xlsx...');
  await new Promise(r => setTimeout(r, 15000));

  const result = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    if (!wb) return { error: 'No active workbook' };

    const sheets = wb.getSheets ? wb.getSheets().map(s => ({ id: s.getSheetId(), name: s.getSheetName() })) : [];
    
    // Inspect cell Summary Hadiah!A1, Summary Hadiah!C4, Summary!C4
    const sHadiah = wb.getSheetByName('Summary Hadiah');
    const sSummary = wb.getSheetByName('Summary');
    const sEntries = wb.getSheetByName('Entries data');

    const snap = wb.save();

    return {
      workbookId: wb.getId(),
      sheets,
      SummaryHadiah_A1: {
        value: sHadiah ? sHadiah.getRange(0, 0, 1, 1).getValue() : null,
        formula: sHadiah ? (sHadiah.getRange(0, 0, 1, 1).getFormula ? sHadiah.getRange(0, 0, 1, 1).getFormula() : null) : null,
        rawSnapCell: snap.sheets[sHadiah ? sHadiah.getSheetId() : 'Summary Hadiah']?.cellData?.[0]?.[0]
      },
      SummaryHadiah_C4: {
        value: sHadiah ? sHadiah.getRange(3, 2, 1, 1).getValue() : null,
        formula: sHadiah ? (sHadiah.getRange(3, 2, 1, 1).getFormula ? sHadiah.getRange(3, 2, 1, 1).getFormula() : null) : null,
        rawSnapCell: snap.sheets[sHadiah ? sHadiah.getSheetId() : 'Summary Hadiah']?.cellData?.[3]?.[2]
      },
      Summary_C4: {
        value: sSummary ? sSummary.getRange(3, 2, 1, 1).getValue() : null,
        rawSnapCell: snap.sheets[sSummary ? sSummary.getSheetId() : 'Summary']?.cellData?.[3]?.[2]
      }
    };
  });

  fs.writeFileSync('barakat-acceptance-report.json', JSON.stringify(result, null, 2));
  console.log('ACCEPTANCE REPORT WRITTEN TO barakat-acceptance-report.json');
  console.log(JSON.stringify(result, null, 2));

  await browser.close();
}

debugImportClick().catch(console.error);
