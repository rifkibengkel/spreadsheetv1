const fs = require('fs');
const puppeteer = require('puppeteer');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Exporting updatehexos.xlsx from Univer and Inspecting XML ===');

  const fileBuffer = fs.readFileSync('C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx');
  const base64Data = fileBuffer.toString('base64');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=8192']
  });

  const page = await browser.newPage();
  page.on('console', msg => console.log('[PAGE LOG]', msg.text()));

  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Importing updatehexos.xlsx into Univer...');

  await page.evaluate(async (b64) => {
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
  }, base64Data);

  console.log('Exporting workbook from Univer snapshot...');

  const exportedB64 = await page.evaluate(async () => {
    const api = window.univerAPI;
    const blob = await window.exportWorkbookToExcel(api);
    const arrayBuffer = await blob.arrayBuffer();
    const bytes = new Uint8Array(arrayBuffer);
    let binary = '';
    for (let i = 0; i < bytes.byteLength; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
  });

  const exportedBuffer = Buffer.from(exportedB64, 'base64');
  fs.writeFileSync('exported-from-univer.xlsx', exportedBuffer);
  console.log(`Saved exported-from-univer.xlsx (${(exportedBuffer.length / (1024*1024)).toFixed(2)} MB). Unzipping for XML inspection...`);

  await browser.close();

  // Inspect XML files in exported-from-univer.xlsx
  const directory = await unzipper.Open.file('exported-from-univer.xlsx');

  const xmlDetails = [];

  for (const file of directory.files) {
    if (file.path === 'xl/workbook.xml' || file.path.startsWith('xl/worksheets/sheet')) {
      const content = await file.buffer();
      const xmlStr = content.toString('utf8');

      // Extract sample formula tags <f...>...</f>
      const formulaMatches = [];
      const regex = /<c r="([A-Z0-9]+)"[^>]*>(?:<f([^>]*)>([^<]*)<\/f>)?(?:<v>([^<]*)<\/v>)?/g;
      let m;
      let count = 0;
      while ((m = regex.exec(xmlStr)) !== null && count < 15) {
        if (m[2] !== undefined || m[3] !== undefined) {
          formulaMatches.push({
            cellRef: m[1],
            fAttrs: m[2],
            fText: m[3],
            vText: m[4]
          });
          count++;
        }
      }

      xmlDetails.push({
        path: file.path,
        sampleFormulas: formulaMatches
      });
    }
  }

  console.log('XML Inspection Report:', JSON.stringify(xmlDetails, null, 2));
  fs.writeFileSync('exported-univer-xml-inspection.json', JSON.stringify(xmlDetails, null, 2));
})();
