const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Forensic Inspection of Univer Shared Formulas in save() Snapshot ===');

  const fileBuffer = fs.readFileSync('C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx');
  const base64Data = fileBuffer.toString('base64');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=8192']
  });

  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Importing updatehexos.xlsx into Univer...');

  const result = await page.evaluate(async (b64) => {
    const binaryString = atob(b64);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const file = new File([blob], 'updatehexos.xlsx');

    const api = window.univerAPI;
    const res = await window.importExcelToWorkbookDataAsync(file);
    window.replaceUniverWorkbook(api, res.workbookData);

    const activeWorkbook = api.getActiveWorkbook();
    const snapshot = activeWorkbook.save();

    // Inspect sheet 8 (or sheets with shared formulas)
    const report = {};

    Object.keys(snapshot.sheets).forEach(sId => {
      const sheet = snapshot.sheets[sId];
      const name = sheet.name || sId;
      const cellData = sheet.cellData || {};

      let hasFAndSi = 0;
      let hasSiOnly = 0;
      let hasFOnly = 0;

      const siSamples = [];

      Object.keys(cellData).forEach(r => {
        Object.keys(cellData[r]).forEach(c => {
          const cell = cellData[r][c];
          if (cell) {
            if (cell.f && cell.si) hasFAndSi++;
            else if (cell.si && !cell.f) hasSiOnly++;
            else if (cell.f && !cell.si) hasFOnly++;

            if (cell.si && siSamples.length < 10) {
              siSamples.push({ r, c, f: cell.f, si: cell.si, v: cell.v });
            }
          }
        });
      });

      report[name] = {
        sheetId: sId,
        hasFAndSi,
        hasSiOnly,
        hasFOnly,
        siSamples
      };
    });

    return report;
  }, base64Data);

  console.log('Shared Formula Snapshot Report:', JSON.stringify(result, null, 2));
  fs.writeFileSync('univer-shared-formula-snapshot-report.json', JSON.stringify(result, null, 2));

  await browser.close();
})();
