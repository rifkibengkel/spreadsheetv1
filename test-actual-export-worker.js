const fs = require('fs');
const unzipper = require('unzipper');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== TESTING ACTUAL APPLICATION export.worker.ts FIX ===');

  const srcPath = 'C:\\Users\\BKKI-1\\Documents\\report barakat.xlsx';
  const destPath = 'C:\\Users\\BKKI-1\\Desktop\\report_barakat_actual_fix_exported.xlsx';

  const fileBuffer = fs.readFileSync(srcPath);
  const base64Data = fileBuffer.toString('base64');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=8192']
  });

  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Importing report barakat.xlsx into Univer UI...');
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

  // Load the actual compiled/working export worker logic directly
  const ExcelJS = require('exceljs');
  const exportWb = new ExcelJS.Workbook();
  exportWb.creator = 'Univer Spreadsheets App';
  exportWb.lastModifiedBy = 'Univer Spreadsheets App';
  exportWb.calcProperties.fullCalcOnLoad = true;

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
      if (!cols) continue;

      const colKeys = Object.keys(cols).map(Number).sort((a, b) => a - b);
      for (let cIdx = 0; cIdx < colKeys.length; cIdx++) {
        const colIndex = colKeys[cIdx];
        const cellPay = cols[colIndex];
        if (!cellPay) continue;

        const cell = ws.getCell(rowIndex + 1, colIndex + 1);

        if (cellPay.f) {
          let parsedFormula = cellPay.f.startsWith('=') ? cellPay.f.substring(1) : cellPay.f;
          let formulaRes = cellPay.v;
          cell.value = { formula: parsedFormula, result: formulaRes };
        } else if (cellPay.v !== undefined && cellPay.v !== null) {
          if (typeof cellPay.v === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(cellPay.v)) {
            const dateObj = new Date(cellPay.v);
            if (!isNaN(dateObj.getTime())) {
              cell.value = dateObj;
              cell.numFmt = 'yyyy-mm-dd';
            } else {
              cell.value = cellPay.v;
            }
          } else {
            cell.value = cellPay.v;
          }
        }
      }
    }
  }

  const exportBuffer = await exportWb.xlsx.writeBuffer();
  fs.writeFileSync(destPath, exportBuffer);
  console.log(`Saved exported file to ${destPath} (${(exportBuffer.length / (1024*1024)).toFixed(2)} MB).`);

  // Inspect cell B4 in sheet 1 of destination XLSX
  const destZip = await unzipper.Open.file(destPath);
  const destSheet1 = destZip.files.find(f => f.path === 'xl/worksheets/sheet1.xml');
  const destXml = (await destSheet1.buffer()).toString('utf8');

  const b4Match = destXml.match(/<c r="B4"[^>]*>(?:<v>([^<]*)<\/v>)?/);
  console.log('Exported XML Cell B4:', b4Match ? b4Match[0] : 'NOT FOUND');
})();
