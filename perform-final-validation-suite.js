const fs = require('fs');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== FULL FORENSIC VALIDATION SUITE FOR EXPORTED XLSX ===');

  const fileBuffer = fs.readFileSync('C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx');
  const base64Data = fileBuffer.toString('base64');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=8192']
  });

  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Step 1: Importing updatehexos.xlsx into Univer...');

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

  console.log('Step 2: Snapshot captured. Exporting XLSX using updated export.worker.ts logic...');

  // Worker logic execution
  const exportWb = new ExcelJS.Workbook();
  exportWb.creator = 'Univer Spreadsheets App';
  exportWb.lastModifiedBy = 'Univer Spreadsheets App';
  exportWb.calcProperties.fullCalcOnLoad = true;

  const sortedSheetIds = snapshot.sheetOrder || Object.keys(snapshot.sheets);

  const availableSheetNames = sortedSheetIds
    .map((id) => snapshot.sheets[id]?.name)
    .filter((name) => typeof name === 'string' && name.length > 0);

  function sanitizeSheetNames(formula) {
    if (!formula) return formula;
    let result = formula;
    for (let i = 0; i < availableSheetNames.length; i++) {
      const name = availableSheetNames[i];
      if (/[\s&\-\.\/\\]/.test(name)) {
        const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const regex = new RegExp(`(?<!')${escaped}(?!')!`, 'g');
        result = result.replace(regex, `'${name}'!`);
      }
    }
    return result;
  }

  const snapshotSheetStats = {};
  let totalSnapshotFormulas = 0;

  sortedSheetIds.forEach(sheetId => {
    const sheet = snapshot.sheets[sheetId];
    if (!sheet) return;
    const name = sheet.name || sheetId;
    const cellData = sheet.cellData || {};

    let formulaCount = 0;
    let masterSharedCount = 0;
    let slaveSharedCount = 0;

    const sharedMap = new Map();
    Object.keys(cellData).forEach(r => {
      Object.keys(cellData[r]).forEach(c => {
        const cell = cellData[r][c];
        if (cell) {
          if (cell.f && cell.si !== undefined && cell.si !== null) {
            sharedMap.set(String(cell.si), cell.f);
          }
        }
      });
    });

    Object.keys(cellData).forEach(r => {
      Object.keys(cellData[r]).forEach(c => {
        const cell = cellData[r][c];
        if (cell) {
          if (cell.f) {
            formulaCount++;
            if (cell.si !== undefined) masterSharedCount++;
          } else if (cell.si !== undefined && sharedMap.has(String(cell.si))) {
            formulaCount++;
            slaveSharedCount++;
          }
        }
      });
    });

    snapshotSheetStats[name] = { formulaCount, masterSharedCount, slaveSharedCount };
    totalSnapshotFormulas += formulaCount;
  });

  for (let sIdx = 0; sIdx < sortedSheetIds.length; sIdx++) {
    const sheetId = sortedSheetIds[sIdx];
    const sheetData = snapshot.sheets[sheetId];
    if (!sheetData || !exportWb) continue;

    const ws = exportWb.addWorksheet(sheetData.name || sheetId);
    const cellDataRaw = sheetData.cellData || {};
    const rowKeys = Object.keys(cellDataRaw).map(Number).sort((a, b) => a - b);

    const sharedFormulaMap = new Map();
    for (let rIdx = 0; rIdx < rowKeys.length; rIdx++) {
      const rowIndex = rowKeys[rIdx];
      const cols = cellDataRaw[rowIndex];
      if (cols) {
        const colKeys = Object.keys(cols).map(Number);
        for (let cIdx = 0; cIdx < colKeys.length; cIdx++) {
          const cellPay = cols[colKeys[cIdx]];
          if (cellPay && cellPay.f && cellPay.si !== undefined && cellPay.si !== null) {
            sharedFormulaMap.set(String(cellPay.si), cellPay.f);
          }
        }
      }
    }

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

          let formulaText = cellPay.f;
          if (!formulaText && cellPay.si !== undefined && cellPay.si !== null) {
            formulaText = sharedFormulaMap.get(String(cellPay.si));
          }

          if (formulaText) {
            let parsedFormula = formulaText;
            if (parsedFormula.startsWith('=')) {
              parsedFormula = parsedFormula.substring(1);
            }
            parsedFormula = sanitizeSheetNames(parsedFormula);

            const isArrayFormula = cellPay.t === 1 || 
              /SUM\s*\(\s*IF\b|COUNT\s*\(\s*IF\b|AVERAGE\s*\(\s*IF\b|MODE\.MULT\b|MAX\s*\(\s*IF\b|MIN\s*\(\s*IF\b/i.test(parsedFormula) ||
              /:\$[A-Z]+\$\d+\s*=|\$[A-Z]+\$\d+:\$[A-Z]+\$\d+\s*=/i.test(parsedFormula) ||
              (parsedFormula.startsWith('{') && parsedFormula.endsWith('}'));

            if (parsedFormula.startsWith('{') && parsedFormula.endsWith('}')) {
              parsedFormula = parsedFormula.substring(1, parsedFormula.length - 1);
            }

            let formulaRes = cellPay.v !== undefined && cellPay.v !== null 
              ? cellPay.v 
              : (cellPay.p?.body?.dataStream ? cellPay.p.body.dataStream.replace(/[\r\n]+$/, '') : undefined);

            if (typeof formulaRes === 'string' && formulaRes.trim() !== '' && !isNaN(Number(formulaRes))) {
              formulaRes = Number(formulaRes);
            }

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

  const exportBuffer = await exportWb.xlsx.writeBuffer();
  fs.writeFileSync('C:\\Users\\BKKI-1\\Desktop\\updatehexos_exported_validated.xlsx', exportBuffer);
  console.log(`Saved exported file to Desktop: C:\\Users\\BKKI-1\\Desktop\\updatehexos_exported_validated.xlsx (${(exportBuffer.length / (1024*1024)).toFixed(2)} MB)`);

  console.log('\nStep 3: Unzipping exported XLSX and verifying XML across all 13 worksheets...');

  const directory = await unzipper.Open.file('C:\\Users\\BKKI-1\\Desktop\\updatehexos_exported_validated.xlsx');

  let totalExportedFormulas = 0;
  let totalExportedArrayFormulas = 0;
  let totalExportedCachedValues = 0;
  let unquotedSheetRefIssueCount = 0;

  const exportedSheetReport = [];

  for (const file of directory.files) {
    if (file.path.startsWith('xl/worksheets/sheet') && file.path.endsWith('.xml')) {
      const content = await file.buffer();
      const xmlStr = content.toString('utf8');

      const formulaCount = (xmlStr.match(/<f[^>]*>/g) || []).length;
      const arrayCount = (xmlStr.match(/<f[^>]*t="array"[^>]*>/g) || []).length;
      const cachedCount = (xmlStr.match(/<v>[^<]*<\/v>/g) || []).length;

      // Check for unquoted sheet names with special characters
      const unquotedSpecial = (xmlStr.match(/<f[^>]*>[^<]*(?:E-Voucher & E-Wallet|Unik Konsumen|Data All|Data Valid|Summary Registrasi|Summary Valid|Registrasi 2025|Entries 2025|Direct Prize|Hadiah Fisik)!/g) || []).length;

      unquotedSheetRefIssueCount += unquotedSpecial;
      totalExportedFormulas += formulaCount;
      totalExportedArrayFormulas += arrayCount;
      totalExportedCachedValues += cachedCount;

      exportedSheetReport.push({
        file: file.path,
        formulaCount,
        arrayCount,
        cachedCount,
        unquotedSpecial
      });
    }
  }

  let calcPrFullCalcOnLoad = false;
  const workbookFile = directory.files.find(f => f.path === 'xl/workbook.xml');
  if (workbookFile) {
    const wbXml = (await workbookFile.buffer()).toString('utf8');
    calcPrFullCalcOnLoad = wbXml.includes('fullCalcOnLoad="1"');
  }

  console.log('\nStep 4: Re-importing exported XLSX back into Univer to verify complete readability...');

  const reimportResult = await page.evaluate(async (b64Exported) => {
    const binaryString = atob(b64Exported);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const file = new File([blob], 'updatehexos_exported_validated.xlsx');

    const api = window.univerAPI;
    const res = await window.importExcelToWorkbookDataAsync(file);
    window.replaceUniverWorkbook(api, res.workbookData);

    const activeWorkbook = api.getActiveWorkbook();
    const snapshotReimported = activeWorkbook.save();

    let reimportedFormulas = 0;
    const sheetCounts = {};

    Object.keys(snapshotReimported.sheets).forEach(sId => {
      const s = snapshotReimported.sheets[sId];
      const name = s.name || sId;
      const cellData = s.cellData || {};
      let count = 0;

      Object.keys(cellData).forEach(r => {
        Object.keys(cellData[r]).forEach(c => {
          if (cellData[r][c] && cellData[r][c].f) count++;
        });
      });

      sheetCounts[name] = count;
      reimportedFormulas += count;
    });

    return {
      sheetCount: Object.keys(snapshotReimported.sheets).length,
      reimportedFormulas,
      sheetCounts
    };
  }, exportBuffer.toString('base64'));

  await browser.close();

  const finalSummary = {
    totalSnapshotFormulas,
    totalExportedFormulas,
    totalExportedArrayFormulas,
    unquotedSheetRefIssueCount,
    calcPrFullCalcOnLoad,
    reimportStatus: reimportResult
  };

  console.log('\n=== FINAL VALIDATION SUMMARY REPORT ===');
  console.log(JSON.stringify(finalSummary, null, 2));

  fs.writeFileSync('final-validation-summary-report.json', JSON.stringify(finalSummary, null, 2));
})();
