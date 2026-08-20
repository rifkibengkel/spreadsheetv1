const fs = require('fs');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

async function testIsolateSheets() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  const buf = fs.readFileSync(filePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);

  // Parse all 5 sheets
  const sheetsData = {};
  const allSheetNames = workbook.worksheets.map(w => w.name);

  workbook.worksheets.forEach((ws, sIdx) => {
    const sheetId = ws.name;
    const cellData = {};
    let maxRow = 0, maxCol = 0;

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const rowIndex = rowNumber - 1;
      if (rowIndex > maxRow) maxRow = rowIndex;

      const rowCellMap = {};
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const colIndex = colNumber - 1;
        if (colIndex > maxCol) maxCol = colIndex;

        let v = undefined, f = undefined, rawFormula = undefined;
        if (cell.type === ExcelJS.ValueType.Formula) {
          rawFormula = cell.formula || (cell.model && cell.model.formula) || (typeof cell.value === 'object' && cell.value !== null ? cell.value.formula : undefined);
        } else if (typeof cell.value === 'object' && cell.value !== null && cell.value.formula) {
          rawFormula = cell.value.formula;
        } else if (cell.model && cell.model.formula) {
          rawFormula = cell.model.formula;
        } else if (cell.formula) {
          rawFormula = cell.formula;
        }

        if (rawFormula && typeof rawFormula === 'string' && rawFormula.trim().length > 0) {
          const trimmed = rawFormula.trim();
          f = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
        }

        if (cell.type === ExcelJS.ValueType.Formula) {
          const res = cell.result;
          if (res !== undefined && res !== null) {
            if (typeof res === 'object') {
              if (res.result !== undefined && res.result !== null && res.result !== '') v = res.result;
              else if (res.error !== undefined) v = res.error;
            } else if (res !== '') v = res;
          }
        } else if (cell.value !== undefined && cell.value !== null) {
          const val = cell.value;
          if (typeof val === 'object') {
            if (val instanceof Date) v = val.toISOString();
            else if (Array.isArray(val.richText)) v = val.richText.map(t => t.text || '').join('');
            else if (val.text !== undefined) v = val.text;
            else if (val.result !== undefined) v = val.result;
            else v = String(val);
          } else {
            v = val;
          }
        }

        const cellInfo = {};
        if (f !== undefined && f !== null && f !== '') cellInfo.f = f;
        if (v !== undefined && v !== null && v !== '') {
          if (typeof v === 'number') { cellInfo.v = v; cellInfo.t = 2; }
          else if (typeof v === 'boolean') { cellInfo.v = v; cellInfo.t = 3; }
          else {
            const trimmedV = String(v).trim();
            const numV = Number(trimmedV);
            if (!isNaN(numV) && trimmedV !== '') { cellInfo.v = numV; cellInfo.t = 2; }
            else { cellInfo.v = String(v); cellInfo.t = 1; }
          }
        }

        if (Object.keys(cellInfo).length > 0) rowCellMap[colIndex] = cellInfo;
      });

      if (Object.keys(rowCellMap).length > 0) cellData[rowIndex] = rowCellMap;
    });

    sheetsData[sheetId] = {
      id: sheetId,
      name: ws.name,
      type: 2,
      status: sIdx === 0 ? 1 : 0,
      rowCount: Math.max(maxRow + 50, 100),
      columnCount: Math.max(maxCol + 10, 30),
      cellData
    };
  });

  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  page.on('console', msg => console.log('[BROWSER]', msg.text()));

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  const combinations = [
    ['Summary', 'Summary Hadiah'],
    ['Summary', 'Summary Hadiah', 'Tren Hadiah'],
    ['Summary', 'Summary Hadiah', 'Tren Hadiah', 'Kode Unik KOL'],
    ['Summary', 'Summary Hadiah', 'Tren Hadiah', 'Kode Unik KOL', 'Entries data']
  ];

  const results = {};

  for (const sheetSubset of combinations) {
    const subKey = sheetSubset.join(' + ');
    console.log(`Testing subset: ${subKey}...`);

    const subSheets = {};
    sheetSubset.forEach((name, idx) => {
      subSheets[name] = { ...sheetsData[name], status: idx === 0 ? 1 : 0 };
    });

    const wbData = {
      id: `wb-sub-${Date.now()}`,
      name: 'Sub Workbook',
      sheetOrder: sheetSubset,
      appVersion: '3.0.0-alpha',
      sheets: subSheets
    };

    const res = await page.evaluate(async (data) => {
      const api = window.univerAPI;
      const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
      oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

      const wb = api.createUniverSheet(data);
      await new Promise(r => setTimeout(r, 2500));

      const sHadiah = wb.getSheetByName('Summary Hadiah');
      const val = sHadiah ? sHadiah.getRange(0, 0, 1, 1).getValue() : null;
      const cellData = wb.save().sheets['Summary Hadiah']?.cellData?.[0]?.[0];

      return { val, cellData };
    }, wbData);

    results[subKey] = res;
  }

  console.log('ISOLATION RESULTS:');
  console.log(JSON.stringify(results, null, 2));

  fs.writeFileSync('isolation-results.json', JSON.stringify(results, null, 2));

  await browser.close();
}

testIsolateSheets().catch(console.error);
