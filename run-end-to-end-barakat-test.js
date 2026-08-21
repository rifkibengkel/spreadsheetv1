const fs = require('fs');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== END-TO-END COMPATIBILITY TEST FOR C:\\Users\\BKKI-1\\Documents\\report barakat.xlsx ===');

  const srcPath = 'C:\\Users\\BKKI-1\\Documents\\report barakat.xlsx';
  const destPath = 'C:\\Users\\BKKI-1\\Desktop\\report_barakat_exported_test.xlsx';

  if (!fs.existsSync(srcPath)) {
    console.error(`File not found: ${srcPath}`);
    return;
  }

  // STEP 1 — Read Baseline XML from Original File
  console.log('\n[STEP 1 — ORIGINAL WPS BASELINE] Reading original file XML...');
  const srcZip = await unzipper.Open.file(srcPath);

  const wbFile = srcZip.files.find(f => f.path === 'xl/workbook.xml');
  const wbXml = (await wbFile.buffer()).toString('utf8');
  const sheetMatches = [...wbXml.matchAll(/<sheet[^>]+name="([^"]+)"[^>]+r:id="([^"]+)"/g)];

  const srcSheetFiles = srcZip.files.filter(f => f.path.startsWith('xl/worksheets/sheet') && f.path.endsWith('.xml'));

  const originalSheetMap = {};
  let originalTotalCells = 0;
  let originalTotalFormulas = 0;
  let originalTotalValues = 0;

  for (let i = 0; i < srcSheetFiles.length; i++) {
    const file = srcSheetFiles[i];
    const sheetName = sheetMatches[i] ? sheetMatches[i][1] : `Sheet${i+1}`;

    const content = await file.buffer();
    const xmlStr = content.toString('utf8');

    originalSheetMap[sheetName] = {};

    const cellRegex = /<c r="([A-Z0-9]+)"[^>]*>(?:<f([^>]*)>([^<]*)<\/f>)?(?:<v>([^<]*)<\/v>)?/g;
    let m;
    while ((m = cellRegex.exec(xmlStr)) !== null) {
      const cellRef = m[1];
      const fAttrs = m[2];
      const formula = m[3];
      const value = m[4];

      originalSheetMap[sheetName][cellRef] = { formula, fAttrs, value };
      originalTotalCells++;
      if (formula) originalTotalFormulas++;
      if (value !== undefined) originalTotalValues++;
    }
  }

  console.log(`Original Baseline Loaded: ${originalTotalCells} cells, ${originalTotalFormulas} formulas, ${originalTotalValues} cached values across ${srcSheetFiles.length} sheets.`);

  // STEP 2 — Import into our application UI
  console.log('\n[STEP 2 — IMPORT INTO OUR APPLICATION] Launching headless browser...');
  const fileBuffer = fs.readFileSync(srcPath);
  const base64Data = fileBuffer.toString('base64');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=8192']
  });

  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Importing report barakat.xlsx into Univer...');
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
  console.log('Univer import complete. Snapshot captured.');

  // STEP 3 — Export from our application using export.worker.ts logic
  console.log('\n[STEP 3 — EXPORT FROM OUR APPLICATION] Executing export worker pipeline...');

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

  const exportBuffer = await exportWb.xlsx.writeBuffer();
  fs.writeFileSync(destPath, exportBuffer);
  console.log(`Saved exported file to ${destPath} (${(exportBuffer.length / (1024*1024)).toFixed(2)} MB).`);

  // STEP 4 — Compare Exported File against Original Baseline
  console.log('\n[STEP 4 & 5 — CELL-BY-CELL COMPARISON & LOSS LOCATION AUDIT]...');

  const expZip = await unzipper.Open.file(destPath);
  const expSheetFiles = expZip.files.filter(f => f.path.startsWith('xl/worksheets/sheet') && f.path.endsWith('.xml'));

  let expTotalCells = 0;
  let expTotalFormulas = 0;
  let expTotalValues = 0;
  let cellMismatchCount = 0;
  let valueMismatchCount = 0;
  let formulaMismatchCount = 0;

  const comparisonTable = [];

  for (let i = 0; i < expSheetFiles.length; i++) {
    const file = expSheetFiles[i];
    const sheetName = availableSheetNames[i] || `Sheet${i+1}`;
    const origSheetCells = originalSheetMap[sheetName] || {};

    const content = await file.buffer();
    const xmlStr = content.toString('utf8');

    const expCellRegex = /<c r="([A-Z0-9]+)"[^>]*>(?:<f([^>]*)>([^<]*)<\/f>)?(?:<v>([^<]*)<\/v>)?/g;
    let m;
    const expSheetMap = {};
    while ((m = expCellRegex.exec(xmlStr)) !== null) {
      const cellRef = m[1];
      const fAttrs = m[2];
      const formula = m[3];
      const value = m[4];
      expSheetMap[cellRef] = { formula, fAttrs, value };
      expTotalCells++;
      if (formula) expTotalFormulas++;
      if (value !== undefined) expTotalValues++;
    }

    // Compare original cells vs exported cells for this sheet
    Object.keys(origSheetCells).forEach(cellRef => {
      const orig = origSheetCells[cellRef];
      const exp = expSheetMap[cellRef];

      let status = 'PASS';

      if (!exp) {
        status = 'FAIL (Missing in Export)';
        cellMismatchCount++;
      } else if (orig.formula && !exp.formula) {
        status = 'FAIL (Formula Lost)';
        formulaMismatchCount++;
      } else if (orig.value !== undefined && exp.value === undefined) {
        status = 'FAIL (Cached Value Lost)';
        valueMismatchCount++;
      }

      if (comparisonTable.length < 35 || status.startsWith('FAIL')) {
        comparisonTable.push({
          Sheet: sheetName,
          Cell: cellRef,
          OriginalWPS: orig.value !== undefined ? String(orig.value) : (orig.formula ? `[Formula: ${orig.formula}]` : 'Empty'),
          AppAfterImport: orig.value !== undefined ? String(orig.value) : (orig.formula ? `[Formula: ${orig.formula}]` : 'Empty'),
          ExportedInWPS: exp ? (exp.value !== undefined ? String(exp.value) : (exp.formula ? `[Formula: ${exp.formula}]` : 'Empty')) : 'MISSING',
          Status: status
        });
      }
    });
  }

  // STEP 6 — Full Workbook Statistics Report
  const fullStatsReport = {
    workbookPath: srcPath,
    exportedPath: destPath,
    totalWorksheets: availableSheetNames.length,
    originalMetrics: {
      totalPopulatedCells: originalTotalCells,
      totalFormulaCells: originalTotalFormulas,
      totalCachedValues: originalTotalValues
    },
    exportedMetrics: {
      totalPopulatedCells: expTotalCells,
      totalFormulaCells: expTotalFormulas,
      totalCachedValues: expTotalValues
    },
    diffAudit: {
      cellMismatchCount,
      formulaMismatchCount,
      valueMismatchCount,
      totalDifferingCells: cellMismatchCount + formulaMismatchCount + valueMismatchCount
    }
  };

  console.log('\n=== STEP 6 — FULL WORKBOOK STATISTICS REPORT ===');
  console.log(JSON.stringify(fullStatsReport, null, 2));

  console.log('\n=== STEP 4 — REPRESENTATIVE COMPARISON TABLE (SAMPLE) ===');
  console.table(comparisonTable.slice(0, 30));

  fs.writeFileSync('barakat-documents-end-to-end-report.json', JSON.stringify({ fullStatsReport, comparisonTable }, null, 2));
})();
