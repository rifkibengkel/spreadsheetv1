const fs = require('fs');
const ExcelJS = require('exceljs');

async function inspectRealFile() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  console.log('Reading file buffer...');
  const buf = fs.readFileSync(filePath);
  console.log('Buffer read, size:', buf.length);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);

  console.log('Sheet names in workbook:');
  workbook.worksheets.forEach((ws, idx) => {
    console.log(`Sheet [${idx}]: name="${ws.name}", id="${ws.id}"`);
  });

  console.log('\nScanning for formulas in sheets:');
  const formulasFound = [];

  workbook.worksheets.forEach((ws) => {
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        let rawFormula = null;
        if (cell.type === ExcelJS.ValueType.Formula) {
          rawFormula = cell.formula || (cell.model && cell.model.formula) || (typeof cell.value === 'object' && cell.value !== null ? cell.value.formula : undefined);
        } else if (typeof cell.value === 'object' && cell.value !== null && cell.value.formula) {
          rawFormula = cell.value.formula;
        }

        if (rawFormula) {
          formulasFound.push({
            sheet: ws.name,
            cell: cell.address,
            formula: rawFormula,
            result: cell.result,
          });
        }
      });
    });
  });

  console.log(`Total formulas found: ${formulasFound.length}`);
  console.log('First 30 formulas:');
  console.log(JSON.stringify(formulasFound.slice(0, 30), null, 2));
}

inspectRealFile().catch(err => {
  console.error('INSPECT ERR:', err);
});
