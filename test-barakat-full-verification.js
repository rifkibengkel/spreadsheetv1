const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

function preResolveWorkbookFormulas(workbookData) {
  if (!workbookData.sheets) return;
  const sheets = workbookData.sheets;

  for (let pass = 0; pass < 5; pass++) {
    let resolvedAny = false;
    Object.keys(sheets).forEach((sName) => {
      const sheet = sheets[sName];
      if (!sheet || !sheet.cellData) return;
      const cData = sheet.cellData;
      Object.keys(cData).forEach((rIdxStr) => {
        const rIdx = parseInt(rIdxStr, 10);
        const row = cData[rIdx];
        if (!row) return;
        Object.keys(row).forEach((cIdxStr) => {
          const cIdx = parseInt(cIdxStr, 10);
          const cell = row[cIdx];
          if (cell && cell.f && (cell.v === undefined || cell.v === null || cell.v === '')) {
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

              const targetSheet = sheets[targetSheetName];
              if (targetSheet && targetSheet.cellData && targetSheet.cellData[targetRow] && targetSheet.cellData[targetRow][targetCol]) {
                const targetCell = targetSheet.cellData[targetRow][targetCol];
                if (targetCell && targetCell.v !== undefined && targetCell.v !== null && targetCell.v !== '') {
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
}

async function verifyBarakatImport() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  console.log('--- Step 1: Parsing report barakat.xlsx ---');
  const buf = fs.readFileSync(filePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);

  const sheetOrder = [];
  const sheetsData = {};

  workbook.worksheets.forEach((ws, sIdx) => {
    const sheetId = ws.name;
    sheetOrder.push(sheetId);
    const cellData = {};

    let maxRow = 0;
    let maxCol = 0;

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const rowIndex = rowNumber - 1;
      if (rowIndex > maxRow) maxRow = rowIndex;

      const rowCellMap = {};
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const colIndex = colNumber - 1;
        if (colIndex > maxCol) maxCol = colIndex;

        let v = undefined;
        let f = undefined;

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
          if (res !== undefined && res !== null) {
            if (typeof res === 'object') {
              if (res.result !== undefined && res.result !== null && res.result !== '') v = res.result;
              else if (res.error !== undefined) v = res.error;
            } else if (res !== '') {
              v = res;
            }
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
      hidden: 0,
      rowCount: Math.max(maxRow + 50, 100),
      columnCount: Math.max(maxCol + 10, 30),
      cellData
    };
  });

  const parsedWbData = {
    id: `barakat-imported-${Date.now()}`,
    name: 'report barakat.xlsx',
    sheetOrder,
    appVersion: '3.0.0-alpha',
    sheets: sheetsData
  };

  // Run pre-resolution pass as in import pipeline
  preResolveWorkbookFormulas(parsedWbData);

  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();

  page.on('console', msg => console.log('[BROWSER LOG]', msg.text()));

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  const verification = await page.evaluate(async (wbData) => {
    const api = window.univerAPI;
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
    oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

    const wb = api.createUniverSheet(wbData);

    await new Promise(r => setTimeout(r, 4000));

    const checkCells = [
      { sheet: 'Summary', r: 0, c: 0, label: 'Summary!A1' },
      { sheet: 'Summary', r: 3, c: 2, label: 'Summary!C4' },
      { sheet: 'Summary', r: 3, c: 14, label: 'Summary!O4' },
      { sheet: 'Summary', r: 3, c: 15, label: 'Summary!P4' },
      { sheet: 'Summary Hadiah', r: 0, c: 0, label: 'Summary Hadiah!A1' },
      { sheet: 'Summary Hadiah', r: 3, c: 2, label: 'Summary Hadiah!C4' },
      { sheet: 'Kode Unik KOL', r: 2, c: 3, label: 'Kode Unik KOL!D3' },
      { sheet: 'Kode Unik KOL', r: 2, c: 6, label: 'Kode Unik KOL!G3' },
      { sheet: 'Kode Unik KOL', r: 2, c: 7, label: 'Kode Unik KOL!H3' }
    ];

    const results = checkCells.map(item => {
      const sh = wb.getSheetByName(item.sheet);
      if (!sh) return { ...item, error: 'Sheet not found' };
      const range = sh.getRange(item.r, item.c, 1, 1);
      return {
        label: item.label,
        sheet: item.sheet,
        value: range.getValue(),
        formula: range.getFormula ? range.getFormula() : undefined,
        snapshotCell: wb.save().sheets[item.sheet]?.cellData?.[item.r]?.[item.c]
      };
    });

    return results;
  }, parsedWbData);

  fs.writeFileSync('barakat-pipeline-out.json', JSON.stringify(verification, null, 2));
  console.log('SAVED TO barakat-pipeline-out.json');
  console.log(JSON.stringify(verification, null, 2));
  await browser.close();
}

verifyBarakatImport().catch(console.error);
