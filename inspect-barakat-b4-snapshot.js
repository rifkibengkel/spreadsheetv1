const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Inspecting Snapshot Data for Summary B4, H4, N4 ===');

  const srcPath = 'C:\\Users\\BKKI-1\\Documents\\report barakat.xlsx';
  const fileBuffer = fs.readFileSync(srcPath);
  const base64Data = fileBuffer.toString('base64');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=8192']
  });

  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  const snapshot = await page.evaluate(async (b64) => {
    const binaryString = atob(b64);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const file = new File([blob], 'report barakat.xlsx');

    const api = window.univerAPI;
    const result = await window.importExcelToWorkbookDataAsync(file);
    window.replaceUniverWorkbook(api, result.workbookData);

    const activeWorkbook = api.getActiveWorkbook();
    return activeWorkbook.save();
  }, base64Data);

  await browser.close();

  const summarySheetId = snapshot.sheetOrder[0];
  const summarySheet = snapshot.sheets[summarySheetId];
  const cellData = summarySheet.cellData;

  // Row 3 (which is Row 4 1-indexed, i.e. B4 is row 3 col 1)
  console.log('Row 3 (Row 4 in Excel):', JSON.stringify(cellData[3], null, 2));
})();
