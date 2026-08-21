const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Detailed Inspection of Summary Row 3 (Row 4 in Excel) ===');

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
  const row3 = summarySheet.cellData[3];

  Object.keys(row3).forEach(cIdx => {
    const colNum = Number(cIdx) + 1;
    let temp, colRef = '';
    let cn = colNum;
    while (cn > 0) {
      temp = (cn - 1) % 26;
      colRef = String.fromCharCode(65 + temp) + colRef;
      cn = (cn - temp) / 26 | 0;
    }
    console.log(`Cell ${colRef}4 (col ${cIdx}):`, JSON.stringify(row3[cIdx]));
  });
})();
