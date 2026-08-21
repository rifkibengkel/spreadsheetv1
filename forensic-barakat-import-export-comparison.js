const fs = require('fs');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== FORENSIC TEST: report barakat_07-03-07.xlsx IMPORT vs EXPORT ===');

  const srcPath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat_07-03-07.xlsx';
  const destPath = 'C:\\Users\\BKKI-1\\Desktop\\barakat_exported.xlsx';

  if (!fs.existsSync(srcPath)) {
    console.error(`File not found: ${srcPath}`);
    return;
  }

  const fileBuffer = fs.readFileSync(srcPath);
  const base64Data = fileBuffer.toString('base64');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=8192']
  });

  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('\n[Phase 1] Importing report barakat_07-03-07.xlsx into Univer UI...');

  const snapshot = await page.evaluate(async (b64) => {
    const binaryString = atob(b64);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const file = new File([blob], 'report barakat_07-03-07.xlsx');

    const api = window.univerAPI;
    const result = await window.importExcelToWorkbookDataAsync(file);
    window.replaceUniverWorkbook(api, result.workbookData);

    const activeWorkbook = api.getActiveWorkbook();
    return activeWorkbook.save();
  }, base64Data);

  await browser.close();
  console.log('[Phase 1 Complete] Snapshot captured successfully from Univer.');

  console.log('\n[Phase 2] Exporting snapshot to XLSX using application worker pipeline...');

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

  // Pre-collect stats before export
  let beforePopulatedCells = 0;
  let beforeFormulaCells = 0;
  let beforeSharedMasterCells = 0;
  let beforeSharedSlaveCells = 0;

  const sheetDataBeforeMap = {};

  sortedSheetIds.forEach(sheetId => {
    const sData = snapshot.sheets[sheetId];
    if (!sData) return;
    const name = sData.name || sheetId;
    const cellDataRaw = sData.cellData || {};

    sheetDataBeforeMap[name] = {};

    const sharedMap = new Map();
    Object.keys(cellDataRaw).forEach(r => {
      Object.keys(cellDataRaw[r]).forEach(c => {
        const cell = cellDataRaw[r][c];
        if (cell && cell.f && cell.si !== undefined && cell.si !== null) {
          sharedMap.set(String(cell.si), cell.f);
        }
      });
    });

    Object.keys(cellDataRaw).forEach(r => {
      Object.keys(cellDataRaw[r]).forEach(c => {
        const cell = cellDataRaw[r][c];
        if (!cell) return;

        let formulaText = cell.f;
        if (!formulaText && cell.si !== undefined && cell.si !== null) {
          formulaText = sharedMap.get(String(cell.si));
        }

        const hasVal = cell.v !== undefined && cell.v !== null;
        if (hasVal || formulaText) {
          beforePopulatedCells++;
        }

        if (formulaText) {
          beforeFormulaCells++;
          if (cell.f && cell.si !== undefined) beforeSharedMasterCells++;
          else if (!cell.f && cell.si !== undefined) beforeSharedSlaveCells++;
        }

        // Store cell coordinate (e.g. A1, I5)
        const colNum = Number(c) + 1;
        const rowNum = Number(r) + 1;

        // Convert colNum to letter
        let temp, colRef = '';
        let cn = colNum;
        while (cn > 0) {
          temp = (cn - 1) % 26;
          colRef = String.fromCharCode(65 + temp) + colRef;
          cn = (cn - temp) / 26 | 0;
        }
        const cellRef = `${colRef}${rowNum}`;

        sheetDataBeforeMap[name][cellRef] = {
          v: cell.v,
          f: formulaText,
          si: cell.si,
          t: cell.t
        };
      });
    });
  });

  // Run Export Loop
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
  console.log(`[Phase 2 Complete] Saved exported XLSX to ${destPath} (${(exportBuffer.length / (1024*1024)).toFixed(2)} MB).`);

  console.log('\n[Phase 3] Inspecting Exported XLSX OpenXML and comparing cell by cell against Univer snapshot...');

  const directory = await unzipper.Open.file(destPath);

  let afterPopulatedCells = 0;
  let afterFormulaCells = 0;
  let afterArrayFormulas = 0;
  let afterCachedValues = 0;

  const comparisonTable = [];
  let valuesLostCount = 0;
  let cachedResultsLostCount = 0;

  // Read worksheets from zip
  const sheetFiles = directory.files.filter(f => f.path.startsWith('xl/worksheets/sheet') && f.path.endsWith('.xml'));

  // Get sheet order from workbook.xml
  const wbFile = directory.files.find(f => f.path === 'xl/workbook.xml');
  const wbXmlStr = (await wbFile.buffer()).toString('utf8');

  const sheetNameMatches = [...wbXmlStr.matchAll(/<sheet[^>]+name="([^"]+)"[^>]+r:id="([^"]+)"/g)];
  const sheetIdToNameMap = {};
  sheetNameMatches.forEach(m => {
    sheetIdToNameMap[m[2]] = m[1];
  });

  // Process each worksheet XML
  for (let idx = 0; idx < sheetFiles.length; idx++) {
    const file = sheetFiles[idx];
    const sheetName = availableSheetNames[idx] || `Sheet${idx+1}`;

    const content = await file.buffer();
    const xmlStr = content.toString('utf8');

    const beforeSheetData = sheetDataBeforeMap[sheetName] || {};

    // Match all cell XML tags: <c r="I5" ...>(?:<f...>...</f>)?(?:<v>...</v>)?
    const cellRegex = /<c r="([A-Z0-9]+)"[^>]*>(?:<f([^>]*)>([^<]*)<\/f>)?(?:<v>([^<]*)<\/v>)?/g;
    let match;

    const exportedSheetCellMap = {};

    while ((match = cellRegex.exec(xmlStr)) !== null) {
      const cellRef = match[1];
      const fAttrs = match[2];
      const expFormula = match[3];
      const expValue = match[4];

      exportedSheetCellMap[cellRef] = {
        formula: expFormula,
        fAttrs: fAttrs,
        value: expValue
      };

      if (expFormula || expValue !== undefined) {
        afterPopulatedCells++;
      }
      if (expFormula) {
        afterFormulaCells++;
        if (fAttrs && fAttrs.includes('t="array"')) afterArrayFormulas++;
      }
      if (expValue !== undefined) {
        afterCachedValues++;
      }
    }

    // Compare before vs after for cells in this sheet
    Object.keys(beforeSheetData).forEach(cellRef => {
      const before = beforeSheetData[cellRef];
      const after = exportedSheetCellMap[cellRef];

      let status = 'PASS';

      if (!after) {
        status = 'FAIL (Cell Disappeared)';
        valuesLostCount++;
      } else if (before.f && !after.formula) {
        status = 'FAIL (Formula Lost)';
      } else if (before.v !== undefined && before.v !== null && after.value === undefined) {
        status = 'FAIL (Value/Result Lost)';
        if (before.f) cachedResultsLostCount++;
        else valuesLostCount++;
      }

      // Collect sample representative comparison rows (up to 25 rows for table output)
      if (comparisonTable.length < 30 || status.startsWith('FAIL')) {
        comparisonTable.push({
          Sheet: sheetName,
          Cell: cellRef,
          ValueBefore: before.v !== undefined ? String(before.v) : 'undefined',
          FormulaBefore: before.f || 'None',
          ExportedFormula: after ? (after.formula || 'None') : 'MISSING',
          ExportedValue: after ? (after.value !== undefined ? String(after.value) : 'undefined') : 'MISSING',
          Status: status
        });
      }
    });
  }

  const forensicSummary = {
    beforeStats: {
      populatedCells: beforePopulatedCells,
      formulaCells: beforeFormulaCells,
      sharedMasterCells: beforeSharedMasterCells,
      sharedSlaveCells: beforeSharedSlaveCells
    },
    afterStats: {
      populatedCells: afterPopulatedCells,
      formulaCells: afterFormulaCells,
      arrayFormulas: afterArrayFormulas,
      cachedValues: afterCachedValues
    },
    lossAudit: {
      totalValuesLost: valuesLostCount,
      totalCachedResultsLost: cachedResultsLostCount
    }
  };

  console.log('\n=== FORENSIC SUMMARY REPORT ===');
  console.log(JSON.stringify(forensicSummary, null, 2));

  console.log('\n=== REPRESENTATIVE CELL COMPARISON TABLE (SAMPLE) ===');
  console.table(comparisonTable.slice(0, 25));

  fs.writeFileSync('barakat-forensic-summary.json', JSON.stringify({ forensicSummary, comparisonTable }, null, 2));
})();
