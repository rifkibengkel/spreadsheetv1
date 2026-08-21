const fs = require('fs');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Export Direct Forensic XML Inspection ===');

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

  const snapshot = await page.evaluate(async (b64) => {
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
    return activeWorkbook.save();
  }, base64Data);

  await browser.close();

  console.log('Snapshot received. Creating ExcelJS Workbook DOM in Node.js...');

  const exportWb = new ExcelJS.Workbook();
  exportWb.creator = 'Univer Spreadsheets App';
  exportWb.lastModifiedBy = 'Univer Spreadsheets App';

  const sortedSheetIds = snapshot.sheetOrder || Object.keys(snapshot.sheets);

  for (let sIdx = 0; sIdx < sortedSheetIds.length; sIdx++) {
    const sheetId = sortedSheetIds[sIdx];
    const sheetData = snapshot.sheets[sheetId];
    if (!sheetData) continue;

    const ws = exportWb.addWorksheet(sheetData.name || sheetId);
    const cellDataRaw = sheetData.cellData || {};
    const rowKeys = Object.keys(cellDataRaw).map(Number).sort((a, b) => a - b);

    for (let rIdx = 0; rIdx < rowKeys.length; rIdx++) {
      const rowIndex = rowKeys[rIdx];
      const cols = cellDataRaw[rowIndex];
      if (cols) {
        const colKeys = Object.keys(cols).map(Number).sort((a, b) => a - b);
        for (let cIdx = 0; cIdx < colKeys.length; cIdx++) {
          const colIndex = colKeys[cIdx];
          const cellPay = cols[colIndex];
          if (!cellPay) continue;

          const cell = ws.getCell(rowIndex + 1, colIndex + 1);
          if (cellPay.f) {
            let parsedFormula = cellPay.f;
            if (parsedFormula.startsWith('=')) {
              parsedFormula = parsedFormula.substring(1);
            }
            const formulaRes = cellPay.v !== undefined && cellPay.v !== null 
              ? cellPay.v 
              : (cellPay.p?.body?.dataStream ? cellPay.p.body.dataStream.replace(/[\r\n]+$/, '') : undefined);

            cell.value = {
              formula: parsedFormula,
              result: formulaRes,
            };
          } else if (cellPay.v !== undefined && cellPay.v !== null) {
            cell.value = cellPay.v;
          } else if (cellPay.p && cellPay.p.body && typeof cellPay.p.body.dataStream === 'string') {
            const rawText = cellPay.p.body.dataStream.replace(/[\r\n]+$/, '');
            if (rawText.length > 0) {
              cell.value = rawText;
            }
          }
        }
      }
    }
  }

  console.log('Writing exported buffer...');
  const arrayBuffer = await exportWb.xlsx.writeBuffer();
  fs.writeFileSync('node-exported-updatehexos.xlsx', arrayBuffer);

  console.log(`Saved node-exported-updatehexos.xlsx (${(arrayBuffer.length / (1024*1024)).toFixed(2)} MB). Unzipping for XML inspection...`);

  // Unzip and inspect XML
  const directory = await unzipper.Open.file('node-exported-updatehexos.xlsx');

  const xmlDetails = [];
  for (const file of directory.files) {
    if (file.path === 'xl/workbook.xml' || file.path.startsWith('xl/worksheets/sheet')) {
      const content = await file.buffer();
      const xmlStr = content.toString('utf8');

      const formulaMatches = [];
      const regex = /<c r="([A-Z0-9]+)"[^>]*>(?:<f([^>]*)>([^<]*)<\/f>)?(?:<v>([^<]*)<\/v>)?/g;
      let m;
      let count = 0;
      while ((m = regex.exec(xmlStr)) !== null && count < 10) {
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

  console.log('XML Forensic Report:', JSON.stringify(xmlDetails, null, 2));
  fs.writeFileSync('node-exported-xml-report.json', JSON.stringify(xmlDetails, null, 2));
})();
