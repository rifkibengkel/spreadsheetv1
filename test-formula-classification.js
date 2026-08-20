const fs = require('fs');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

function isExpensiveFormula(formulaStr) {
  if (!formulaStr || typeof formulaStr !== 'string') return false;
  // Match range row numbers span > 10000 rows, e.g., 2:155810 or 1:50000
  const rangeMatch = formulaStr.match(/:[A-Za-z]*\$?([0-9]{5,})/);
  if (rangeMatch && parseInt(rangeMatch[1], 10) >= 10000) {
    return true;
  }
  return false;
}

async function testFormulaClassification() {
  console.log('=====================================================================');
  console.log('TESTING FORMULA CLASSIFICATION & SCHEDULING LAYER');
  console.log('=====================================================================');

  const barakatPath = 'c:\\Users\\BKKI-1\\.gemini\\antigravity\\scratch\\sheet-V-2\\barakat-forensic.xlsx';
  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(barakatPath, {
    entries: 'emit',
    sharedStrings: 'cache',
    styles: 'cache'
  });

  const sheetOrder = [];
  const sheetsData = {};

  let totalFormulas = 0;
  let expensiveFormulasCount = 0;
  let simpleFormulasCount = 0;

  const classifiedExpensive = [];

  for await (const worksheetReader of workbookReader) {
    const sName = worksheetReader.name;
    sheetOrder.push(sName);
    const cellData = {};
    let maxRow = 0;
    let maxCol = 0;

    for await (const row of worksheetReader) {
      const rIdx = row.number - 1;
      if (rIdx > maxRow) maxRow = rIdx;

      const rowCellMap = {};
      row.eachCell((cell, colNumber) => {
        const cIdx = colNumber - 1;
        if (cIdx > maxCol) maxCol = cIdx;

        let v = undefined;
        let f = undefined;

        if (cell.type === ExcelJS.ValueType.Formula) {
          totalFormulas++;
          const rawF = cell.formula ? (cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`) : undefined;
          
          if (rawF) {
            // Clean malformed WPS sheet references like ='Summary '[Hadiah]#REF'!A1
            f = rawF.replace(/'\[[^'\]]+\]#REF'/g, '');
          }

          if (cell.result !== undefined && cell.result !== null && cell.result !== '') {
            v = typeof cell.result === 'object' ? cell.result.result : cell.result;
          }

          const expensive = isExpensiveFormula(f);
          if (expensive) {
            expensiveFormulasCount++;
            if (classifiedExpensive.length < 5) {
              classifiedExpensive.push({ cell: `${sName}!${cell.address}`, formula: f, initialResult: v });
            }
            // For expensive formulas without cached v, assign default 0 so WHEN_EMPTY does not block worker queue
            if (v === undefined || v === null || v === '') {
              v = 0;
            }
          } else {
            simpleFormulasCount++;
          }
        } else if (cell.value !== undefined && cell.value !== null) {
          v = typeof cell.value === 'object' ? cell.value.text || cell.value.result : cell.value;
        }

        const cellInfo = {};
        if (f) cellInfo.f = f;
        if (v !== undefined && v !== null && v !== '') {
          cellInfo.v = typeof v === 'number' ? v : String(v);
          cellInfo.t = typeof v === 'number' ? 2 : 1;
        }
        if (Object.keys(cellInfo).length > 0) rowCellMap[cIdx] = cellInfo;
      });
      if (Object.keys(rowCellMap).length > 0) cellData[rIdx] = rowCellMap;
    }

    sheetsData[sName] = {
      id: sName,
      name: sName,
      type: 2,
      cellData,
      rowCount: Math.max(maxRow + 50, 100),
      columnCount: Math.max(maxCol + 10, 30)
    };
  }

  console.log(`- Total Formulas Processed: ${totalFormulas}`);
  console.log(`- Simple Formulas: ${simpleFormulasCount}`);
  console.log(`- Classified Expensive Formulas (>10,000 rows): ${expensiveFormulasCount}`);
  console.log('Sample Expensive Formulas:', JSON.stringify(classifiedExpensive, null, 2));

  // Perform pre-resolution for simple formulas in worker logic
  for (let pass = 0; pass < 5; pass++) {
    let resolvedAny = false;
    Object.keys(sheetsData).forEach((sName) => {
      const cData = sheetsData[sName].cellData;
      Object.keys(cData).forEach((rIdxStr) => {
        const rIdx = parseInt(rIdxStr, 10);
        Object.keys(cData[rIdx]).forEach((cIdxStr) => {
          const cIdx = parseInt(cIdxStr, 10);
          const cell = cData[rIdx][cIdx];
          if (cell.f && (cell.v === undefined || cell.v === null || cell.v === '')) {
            let match = cell.f.match(/^=('([^']+)'|([A-Za-z0-9_\s]+))!([A-Za-z]+)([0-9]+)$/i);
            let targetSheetName = match ? (match[2] || match[3]) : null;
            let colStr = match ? match[4].toUpperCase() : null;
            let rowNum = match ? parseInt(match[5], 10) : null;

            if (!match) {
              const sameMatch = cell.f.match(/^=([A-Za-z]+)([0-9]+)$/i);
              if (sameMatch) {
                targetSheetName = sName;
                colStr = sameMatch[1].toUpperCase();
                rowNum = parseInt(sameMatch[2], 10);
              }
            }

            if (targetSheetName && colStr && rowNum) {
              let targetCol = 0;
              for (let i = 0; i < colStr.length; i++) {
                targetCol = targetCol * 26 + (colStr.charCodeAt(i) - 64);
              }
              targetCol -= 1;
              const targetRow = rowNum - 1;

              const targetSheet = sheetsData[targetSheetName];
              if (targetSheet && targetSheet.cellData[targetRow] && targetSheet.cellData[targetRow][targetCol]) {
                const targetCell = targetSheet.cellData[targetRow][targetCol];
                if (targetCell.v !== undefined && targetCell.v !== null && targetCell.v !== '') {
                  cell.v = targetCell.v;
                  cell.t = targetCell.t !== undefined ? targetCell.t : (typeof targetCell.v === 'number' ? 2 : 1);
                  resolvedAny = true;
                }
              }
            }
          }
        });
      });
    });
    if (!resolvedAny) break;
  }

  const workbookData = {
    id: `classification-test-${Date.now()}`,
    name: 'report barakat.xlsx',
    sheetOrder,
    sheets: sheetsData
  };

  // Launch browser session to test formula calculation queue behavior
  console.log('\n--- TESTING BROWSER IMPORT & FORMULA WORKER QUEUE BEHAVIOR ---');
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();

  page.on('console', msg => {
    const txt = msg.text();
    if (txt.includes('[Formula Worker]') || txt.includes('[FORENSIC]')) console.log('[BROWSER]', txt);
  });

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  const targets = [
    { name: 'Simple Same-sheet (Summary!I9)', sheet: 'Summary', row: 8, col: 8 },
    { name: 'Simple Cross-sheet (Summary Hadiah!A1)', sheet: 'Summary Hadiah', row: 0, col: 0 },
    { name: 'Large Cross-sheet (Summary Hadiah!C4)', sheet: 'Summary Hadiah', row: 3, col: 2 },
    { name: 'Shared Follower (Summary Hadiah!C5)', sheet: 'Summary Hadiah', row: 4, col: 2 }
  ];

  const browserEval = await page.evaluate(async (wbData, targetCells) => {
    const api = window.univerAPI;
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
    oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

    const wb = api.createUniverSheet(wbData);

    const initialValues = targetCells.map(t => {
      const sheet = wb.getSheetByName(t.sheet);
      const range = sheet ? sheet.getRange(t.row, t.col, 1, 1) : null;
      return {
        name: t.name,
        formula: range ? range.getFormula() : null,
        value: range ? range.getValue() : null
      };
    });

    await new Promise(r => setTimeout(r, 3000));

    const settledValues = targetCells.map(t => {
      const sheet = wb.getSheetByName(t.sheet);
      const range = sheet ? sheet.getRange(t.row, t.col, 1, 1) : null;
      return {
        name: t.name,
        formula: range ? range.getFormula() : null,
        value: range ? range.getValue() : null
      };
    });

    return { initialValues, settledValues };
  }, workbookData, targets);

  fs.writeFileSync('c:\\Users\\BKKI-1\\.gemini\\antigravity\\scratch\\sheet-V-2\\classification-output.json', JSON.stringify({
    totalFormulas,
    simpleFormulasCount,
    expensiveFormulasCount,
    classifiedExpensive,
    browserEval
  }, null, 2));

  console.log('\nSaved classification verification results to classification-output.json!');
  await browser.close();
}

testFormulaClassification().catch(console.error);
