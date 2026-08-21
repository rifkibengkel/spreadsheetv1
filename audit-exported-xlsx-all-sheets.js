const fs = require('fs');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Forensic Audit of All 13 Sheets in Exported XLSX ===');

  const fileBuffer = fs.readFileSync('C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx');
  const base64Data = fileBuffer.toString('base64');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=8192']
  });

  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Importing updatehexos.xlsx...');

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

  console.log('Snapshot captured. Running export worker logic in Node.js...');

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
            parsedFormula = parsedFormula.replace(/(?<!')\b([A-Za-z0-9_]+(?:\s+[A-Za-z0-9_]+)+)\b(?!')!/g, "'$1'!");

            const isArrayFormula = cellPay.t === 1 || 
              /SUM\s*\(\s*IF\b|COUNT\s*\(\s*IF\b|AVERAGE\s*\(\s*IF\b|MODE\.MULT\b/i.test(parsedFormula) ||
              (parsedFormula.startsWith('{') && parsedFormula.endsWith('}'));

            if (parsedFormula.startsWith('{') && parsedFormula.endsWith('}')) {
              parsedFormula = parsedFormula.substring(1, parsedFormula.length - 1);
            }

            const formulaRes = cellPay.v !== undefined && cellPay.v !== null 
              ? cellPay.v 
              : (cellPay.p?.body?.dataStream ? cellPay.p.body.dataStream.replace(/[\r\n]+$/, '') : undefined);

            const cellObj = {
              formula: parsedFormula,
              result: formulaRes,
            };

            if (isArrayFormula) {
              cellObj.shareType = 'array';
              cellObj.ref = cell.address;
            }

            cell.value = cellObj;
          } else if (cellPay.v !== undefined && cellPay.v !== null) {
            cell.value = cellPay.v;
          }
        }
      }
    }
  }

  const arrayBuffer = await exportWb.xlsx.writeBuffer();
  fs.writeFileSync('full-audit-exported.xlsx', arrayBuffer);
  console.log(`Saved full-audit-exported.xlsx (${(arrayBuffer.length / (1024*1024)).toFixed(2)} MB). Unzipping for XML inspection...`);

  const directory = await unzipper.Open.file('full-audit-exported.xlsx');

  const sheetAudit = [];

  for (const file of directory.files) {
    if (file.path.startsWith('xl/worksheets/sheet') && file.path.endsWith('.xml')) {
      const content = await file.buffer();
      const xmlStr = content.toString('utf8');

      const formulaTags = (xmlStr.match(/<f[^>]*>[^<]*<\/f>/g) || []).length;
      const arrayFormulaTags = (xmlStr.match(/<f[^>]*t="array"[^>]*>[^<]*<\/f>/g) || []).length;
      const cachedValues = (xmlStr.match(/<v>[^<]*<\/v>/g) || []).length;

      // Extract 3 sample formulas from this sheet
      const sampleFormulas = [];
      const regex = /<c r="([A-Z0-9]+)"[^>]*>(?:<f([^>]*)>([^<]*)<\/f>)?(?:<v>([^<]*)<\/v>)?/g;
      let m;
      while ((m = regex.exec(xmlStr)) !== null && sampleFormulas.length < 5) {
        if (m[3]) {
          sampleFormulas.push({ cell: m[1], fAttrs: m[2], formula: m[3], v: m[4] });
        }
      }

      sheetAudit.push({
        file: file.path,
        formulaTags,
        arrayFormulaTags,
        cachedValues,
        sampleFormulas
      });
    }
  }

  console.log('Full Sheet Audit Report:', JSON.stringify(sheetAudit, null, 2));
  fs.writeFileSync('full-sheet-audit-report.json', JSON.stringify(sheetAudit, null, 2));
})();
