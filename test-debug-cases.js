const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

async function testDebug() {
  console.log('--- Creating XLSX ---');
  const wb = new ExcelJS.Workbook();
  const s1 = wb.addWorksheet('Sheet1');
  s1.getCell('A1').value = 100;
  const s2 = wb.addWorksheet('Sheet2');
  s2.getCell('A1').value = { formula: 'Sheet1!A1' };
  await wb.xlsx.writeFile('test-compare.xlsx');

  const fileBuffer = fs.readFileSync('test-compare.xlsx');
  const wbRead = new ExcelJS.Workbook();
  await wbRead.xlsx.load(fileBuffer);

  const sheetsData = {};
  const sheetOrder = [];

  wbRead.worksheets.forEach((ws, sIdx) => {
    const sheetId = ws.name;
    sheetOrder.push(sheetId);
    const cellData = {};

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const rowIndex = rowNumber - 1;
      const rowCellMap = {};
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const colIndex = colNumber - 1;
        let f = undefined;
        let v = undefined;

        let rawFormula = undefined;
        if (cell.type === ExcelJS.ValueType.Formula) {
          rawFormula = cell.formula || (cell.model && cell.model.formula) || (typeof cell.value === 'object' && cell.value !== null ? cell.value.formula : undefined);
        } else if (typeof cell.value === 'object' && cell.value !== null && cell.value.formula) {
          rawFormula = cell.value.formula;
        }

        if (rawFormula && typeof rawFormula === 'string' && rawFormula.trim().length > 0) {
          const trimmed = rawFormula.trim();
          f = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
        }

        if (cell.type === ExcelJS.ValueType.Formula) {
          const res = cell.result;
          if (res !== undefined && res !== null) v = res;
        } else if (cell.value !== undefined && cell.value !== null) {
          v = cell.value;
        }

        const cellInfo = {};
        if (f !== undefined && f !== null && f !== '') cellInfo.f = f;
        if (v !== undefined && v !== null && v !== '') {
          if (typeof v === 'number') { cellInfo.v = v; cellInfo.t = 2; }
          else if (typeof v === 'boolean') { cellInfo.v = v; cellInfo.t = 3; }
          else { cellInfo.v = String(v); cellInfo.t = 1; }
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
      rowCount: 50,
      columnCount: 20,
      cellData
    };
  });

  const impWbData = {
    id: `workbook-imported-${Date.now()}`,
    name: 'test-compare.xlsx',
    sheetOrder,
    appVersion: '3.0.0-alpha',
    sheets: sheetsData
  };

  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  page.on('console', msg => console.log('[PAGE]', msg.text()));
  page.on('pageerror', err => console.log('[PAGE ERROR]', err.message));

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  const res = await page.evaluate(async (data) => {
    const api = window.univerAPI;

    try {
      // 1. Manual Creation
      console.log('Creating manual workbook...');
      const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
      oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

      const manualWb = api.createUniverSheet({
        id: 'wb-manual',
        name: 'Manual',
        sheetOrder: ['s1', 's2'],
        sheets: {
          s1: { id: 's1', name: 'Sheet1', type: 2, status: 1, cellData: {} },
          s2: { id: 's2', name: 'Sheet2', type: 2, status: 0, cellData: {} }
        }
      });

      const ms1 = manualWb.getSheetByName('Sheet1');
      ms1.getRange(0, 0, 1, 1).setValue(100);

      const ms2 = manualWb.getSheetByName('Sheet2');
      ms2.getRange(0, 0, 1, 1).setValue('=Sheet1!A1');

      await new Promise(r => setTimeout(r, 1500));

      const snapA = manualWb.save();
      const valA = ms2.getRange(0, 0, 1, 1).getValue();

      // 2. Imported Workbook
      console.log('Creating imported workbook...');
      try { api.disposeUnit(manualWb.getId()); } catch(e){}
      const impWb = api.createUniverSheet(data);

      await new Promise(r => setTimeout(r, 1500));

      const snapB = impWb.save();
      const is2 = impWb.getSheetByName('Sheet2');
      const valB_imm = is2 ? is2.getRange(0, 0, 1, 1).getValue() : 'NO_SHEET2';

      // 3. Edit Imported Workbook
      console.log('Editing imported workbook cell...');
      if (is2) {
        is2.getRange(0, 0, 1, 1).setValue('=Sheet1!A1');
      }

      await new Promise(r => setTimeout(r, 1500));

      const snapB_after = impWb.save();
      const valB_after = is2 ? is2.getRange(0, 0, 1, 1).getValue() : 'NO_SHEET2';

      return {
        caseA: { val: valA, snapSheet1: snapA.sheets['s1'], snapSheet2: snapA.sheets['s2'] },
        caseB_imm: { val: valB_imm, snapSheet1: snapB.sheets[data.sheetOrder[0]], snapSheet2: snapB.sheets[data.sheetOrder[1]] },
        caseB_after: { val: valB_after, snapSheet1: snapB_after.sheets[data.sheetOrder[0]], snapSheet2: snapB_after.sheets[data.sheetOrder[1]] }
      };
    } catch (err) {
      console.log('EVAL ERR:', err.stack || err.message);
      return { evalError: err.message, stack: err.stack };
    }
  }, impWbData);

  fs.writeFileSync('comparison-out.json', JSON.stringify(res, null, 2));
  console.log('DONE. Written to comparison-out.json');

  await browser.close();
}

testDebug().catch(console.error);
