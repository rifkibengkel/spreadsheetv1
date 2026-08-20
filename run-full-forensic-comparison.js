const ExcelJS = require('exceljs');
const fs = require('fs');
const puppeteer = require('puppeteer');

async function runFullForensicComparison() {
  console.log('=====================================================================');
  console.log('FORENSIC COMPARISON: CASE A (Small) vs CASE B (Real Large) vs CASE C (Re-imported)');
  console.log('=====================================================================');

  // Create Case A: Small external XLSX file
  const wbA = new ExcelJS.Workbook();
  const wsA1 = wbA.addWorksheet('Sheet1');
  wsA1.getCell('A1').value = 100;
  wsA1.getCell('A2').value = 200;

  const wsA2 = wbA.addWorksheet('Sheet2');
  wsA2.getCell('A1').value = { formula: 'Sheet1!A1', result: 100 };
  wsA2.getCell('A2').value = { formula: 'Sheet1!A2', result: 200 };
  wsA2.getCell('A3').value = { formula: 'A1+A2', result: 300 };

  const caseAPath = 'c:\\Users\\BKKI-1\\.gemini\\antigravity\\scratch\\sheet-V-2\\caseA-small.xlsx';
  await wbA.xlsx.writeFile(caseAPath);
  console.log('Created Case A (Small external XLSX file)');

  // 1. Launch Puppeteer
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();

  page.on('console', msg => {
    const txt = msg.text();
    if (txt.includes('[Formula Worker]') || txt.includes('[FORENSIC]')) {
      console.log(txt);
    }
  });

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  // Test Case A in Browser
  console.log('\n--- EVALUATING CASE A (Small External XLSX) ---');
  const bufA = fs.readFileSync(caseAPath);
  const resultCaseA = await page.evaluate(async (arrayBuf) => {
    const api = window.univerAPI;
    const worker = new Worker('/_next/static/chunks/src_lib_workers_excel_worker_ts.js', { type: 'module' });

    return new Promise((resolve) => {
      worker.onmessage = async (e) => {
        if (e.data.type === 'COMPLETE') {
          const wbData = e.data.workbookData;
          const stage3 = wbData.sheets['Sheet2']?.cellData?.[0]?.[0]; // A1 formula

          const wb = api.createUniverSheet(wbData);
          const stage4 = wb.getSheetByName('Sheet2')?.getRange(0, 0, 1, 1).getValue();

          await new Promise(r => setTimeout(r, 2000));
          const stage6 = wb.getSheetByName('Sheet2')?.getRange(0, 0, 1, 1).getValue();
          const snap = wb.save();
          const stage7 = snap.sheets['Sheet2']?.cellData?.[0]?.[0];

          worker.terminate();
          resolve({ stage3, stage4, stage6, stage7 });
        }
      };
      worker.postMessage({ fileArrayBuffer: arrayBuf, fileName: 'caseA-small.xlsx' });
    });
  }, bufA.buffer);
  console.log('Case A Results:', JSON.stringify(resultCaseA, null, 2));

  // Test Case B in Browser (report barakat.xlsx streaming payload)
  console.log('\n--- EVALUATING CASE B (report barakat.xlsx) ---');
  const barakatPath = 'c:\\Users\\BKKI-1\\.gemini\\antigravity\\scratch\\sheet-V-2\\barakat-forensic.xlsx';
  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(barakatPath, {
    entries: 'emit',
    sharedStrings: 'cache',
    styles: 'cache'
  });

  const sheetOrder = [];
  const sheetsData = {};

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
          f = cell.formula ? (cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`) : undefined;
          if (cell.result !== undefined && cell.result !== null && cell.result !== '') {
            v = typeof cell.result === 'object' ? cell.result.result : cell.result;
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

  const workbookDataB = {
    id: `forensic-b-${Date.now()}`,
    name: 'report barakat.xlsx',
    sheetOrder,
    sheets: sheetsData
  };

  const targetsB = [
    { type: 'cross-sheet scalar', sheet: 'Summary Hadiah', row: 0, col: 0, address: 'A1' },
    { type: 'large range formula', sheet: 'Summary Hadiah', row: 3, col: 2, address: 'C4' },
    { type: 'same-sheet formula', sheet: 'Summary', row: 8, col: 8, address: 'I9' },
    { type: 'shared-formula follower', sheet: 'Summary Hadiah', row: 4, col: 2, address: 'C5' }
  ];

  const resultCaseB = await page.evaluate(async (wbData, targets) => {
    const api = window.univerAPI;
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
    oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

    // Stage 3: workbookData immediately before createUniverSheet
    const stage3Before = targets.map(t => {
      const sheet = wbData.sheets[t.sheet];
      const cell = sheet && sheet.cellData[t.row] ? sheet.cellData[t.row][t.col] : null;
      return { address: `${t.sheet}!${t.address}`, cellData: cell };
    });

    // Stage 4: Immediately after createUniverSheet
    const wb = api.createUniverSheet(wbData);

    const stage4AfterCreate = targets.map(t => {
      const sheet = wb.getSheetByName(t.sheet);
      const range = sheet ? sheet.getRange(t.row, t.col, 1, 1) : null;
      return {
        address: `${t.sheet}!${t.address}`,
        value: range ? range.getValue() : null,
        formula: range ? range.getFormula() : null
      };
    });

    // Stage 5 & 6: Formula calculation waiting
    await new Promise(r => setTimeout(r, 4000));

    const stage6AfterCalc = targets.map(t => {
      const sheet = wb.getSheetByName(t.sheet);
      const range = sheet ? sheet.getRange(t.row, t.col, 1, 1) : null;
      const snap = wb.save();
      const rawCell = snap.sheets[t.sheet]?.cellData?.[t.row]?.[t.col];

      return {
        address: `${t.sheet}!${t.address}`,
        formula: range ? range.getFormula() : null,
        facadeValue: range ? range.getValue() : null,
        snapshotCell: rawCell
      };
    });

    return {
      stage3Before,
      stage4AfterCreate,
      stage6AfterCalc
    };
  }, workbookDataB, targetsB);

  fs.writeFileSync('c:\\Users\\BKKI-1\\.gemini\\antigravity\\scratch\\sheet-V-2\\forensic-output.json', JSON.stringify({ resultCaseA, resultCaseB }, null, 2));
  console.log('\nSaved forensic results to forensic-output.json!');

  await browser.close();
}

runFullForensicComparison().catch(console.error);
