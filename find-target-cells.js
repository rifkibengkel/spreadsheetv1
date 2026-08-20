const ExcelJS = require('exceljs');
const path = require('path');

async function findTargetCells() {
  const filePath = 'c:\\Users\\BKKI-1\\.gemini\\antigravity\\scratch\\sheet-V-2\\test-barakat.xlsx';
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(filePath);

  console.log('Worksheets:', workbook.worksheets.map(w => w.name));

  const candidateCells = {
    sameSheet: null,
    sharedFollower: null,
    crossSheet: null,
    largeRange: null
  };

  workbook.worksheets.forEach(ws => {
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (cell.type === ExcelJS.ValueType.Formula) {
          const formula = cell.formula || (cell.value && cell.value.formula) || '';
          const address = cell.address;

          if (!candidateCells.largeRange && formula.includes('155810')) {
            candidateCells.largeRange = { sheet: ws.name, address, formula, result: cell.result, type: cell.type };
          } else if (!candidateCells.crossSheet && formula.includes('!') && !formula.includes('155810')) {
            candidateCells.crossSheet = { sheet: ws.name, address, formula, result: cell.result, type: cell.type };
          } else if (!candidateCells.sharedFollower && cell.sharedFormula) {
            candidateCells.sharedFollower = { sheet: ws.name, address, formula, result: cell.result, type: cell.type };
          } else if (!candidateCells.sameSheet && !formula.includes('!')) {
            candidateCells.sameSheet = { sheet: ws.name, address, formula, result: cell.result, type: cell.type };
          }
        }
      });
    });
  });

  console.log('Candidate Cells Found in report barakat.xlsx:');
  console.log(JSON.stringify(candidateCells, null, 2));
}

findTargetCells().catch(console.error);
