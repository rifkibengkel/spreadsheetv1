const fs = require('fs');
const ExcelJS = require('exceljs');

async function diffPayload() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  const buf = fs.readFileSync(filePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);

  // Manual construct (from testSummaryHadiahA1 which WORKED)
  const manualWbData = {
    id: 'manual-wb',
    name: 'Manual Workbook',
    sheetOrder: ['Summary', 'Summary Hadiah'],
    appVersion: '3.0.0-alpha',
    sheets: {
      'Summary': {
        id: 'Summary',
        name: 'Summary',
        type: 2,
        status: 1,
        rowCount: 50,
        columnCount: 20,
        cellData: {
          0: { 0: { v: 123, t: 2 } }
        }
      },
      'Summary Hadiah': {
        id: 'Summary Hadiah',
        name: 'Summary Hadiah',
        type: 2,
        status: 0,
        rowCount: 50,
        columnCount: 20,
        cellData: {
          0: { 0: { f: '=Summary!A1' } }
        }
      }
    }
  };

  // Barakat parsed construct for Summary and Summary Hadiah
  const barakatSheets = {};
  ['Summary', 'Summary Hadiah'].forEach((name, sIdx) => {
    const ws = workbook.getWorksheet(name);
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

    barakatSheets[name] = {
      id: name,
      name: name,
      type: 2,
      status: sIdx === 0 ? 1 : 0,
      rowCount: Math.max(maxRow + 50, 100),
      columnCount: Math.max(maxCol + 10, 30),
      cellData
    };
  });

  const barakatWbData = {
    id: 'barakat-2sheets',
    name: 'Barakat 2 Sheets',
    sheetOrder: ['Summary', 'Summary Hadiah'],
    appVersion: '3.0.0-alpha',
    sheets: barakatSheets
  };

  fs.writeFileSync('manualWbData.json', JSON.stringify(manualWbData, null, 2));
  fs.writeFileSync('barakatWbData.json', JSON.stringify(barakatWbData, null, 2));

  console.log('Saved manualWbData.json and barakatWbData.json');
}

diffPayload().catch(console.error);
