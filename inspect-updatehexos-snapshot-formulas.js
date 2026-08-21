const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Forensic Inspection of updatehexos.xlsx Snapshot Formulas ===');

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

  const snapshotData = await page.evaluate(async (b64) => {
    const binaryString = atob(b64);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const file = new File([blob], 'updatehexos.xlsx');

    const api = window.univerAPI;
    const result = await window.importExcelToWorkbookDataAsync(file);
    window.replaceUniverWorkbook(api, result.workbookData);

    const activeWorkbook = api.getActiveWorkbook();
    const snapshot = activeWorkbook.save();

    // Collect formula samples from each sheet
    const sheetSamples = {};
    Object.keys(snapshot.sheets).forEach(sheetId => {
      const sheet = snapshot.sheets[sheetId];
      const name = sheet.name || sheetId;
      const cellData = sheet.cellData || {};
      const formulaSamples = [];

      Object.keys(cellData).forEach(r => {
        Object.keys(cellData[r]).forEach(c => {
          const cell = cellData[r][c];
          if (cell && (cell.f || cell.si)) {
            if (formulaSamples.length < 15) {
              formulaSamples.push({
                r,
                c,
                f: cell.f,
                v: cell.v,
                si: cell.si,
                t: cell.t
              });
            }
          }
        });
      });

      sheetSamples[name] = {
        sheetId,
        formulaSampleCount: formulaSamples.length,
        samples: formulaSamples
      };
    });

    return {
      sheetOrder: snapshot.sheetOrder,
      sheetSamples
    };
  }, base64Data);

  console.log('Snapshot Formula Samples Report:', JSON.stringify(snapshotData, null, 2));
  fs.writeFileSync('updatehexos-snapshot-formula-samples.json', JSON.stringify(snapshotData, null, 2));

  await browser.close();
})();
