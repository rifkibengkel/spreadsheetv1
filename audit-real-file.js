const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

async function auditRealExcelFile() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  console.log('--- 1. Loading Real Excel File:', filePath, '---');
  const stat = fs.statSync(filePath);
  console.log(`File size: ${(stat.size / (1024 * 1024)).toFixed(2)} MB`);

  const workbook = new ExcelJS.Workbook();
  const startTime = Date.now();
  await workbook.xlsx.readFile(filePath);
  console.log(`ExcelJS loaded file in ${Date.now() - startTime} ms.`);
  console.log(`Worksheet count: ${workbook.worksheets.length}`);

  const sheetSummaries = [];
  const formulaCellsSample = [];

  workbook.worksheets.forEach((ws, sIdx) => {
    let rowCount = 0;
    let cellCount = 0;
    let formulaCount = 0;
    const formulasInSheet = [];

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      rowCount++;
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        cellCount++;

        let isFormula = false;
        let rawFormula = undefined;

        if (cell.type === ExcelJS.ValueType.Formula) {
          isFormula = true;
          rawFormula = cell.formula || (cell.model && cell.model.formula) || (typeof cell.value === 'object' && cell.value !== null ? cell.value.formula : undefined);
        } else if (typeof cell.value === 'object' && cell.value !== null && cell.value.formula) {
          isFormula = true;
          rawFormula = cell.value.formula;
        } else if (cell.model && cell.model.formula) {
          isFormula = true;
          rawFormula = cell.model.formula;
        }

        if (isFormula) {
          formulaCount++;
          const cellAddr = cell.address;
          const formulaDetail = {
            sheetIndex: sIdx,
            sheetName: ws.name,
            cellAddress: cellAddr,
            rowNumber,
            colNumber,
            cellType: cell.type,
            rawFormula,
            cellResult: cell.result,
            cellValue: cell.value,
            cellModel: cell.model,
            sharedFormula: cell.model ? cell.model.sharedFormula : undefined,
            masterAddress: cell.model ? cell.model.master : undefined
          };
          formulasInSheet.push(formulaDetail);

          if (formulaCellsSample.length < 15) {
            formulaCellsSample.push(formulaDetail);
          }
        }
      });
    });

    sheetSummaries.push({
      index: sIdx,
      name: ws.name,
      rowCount,
      cellCount,
      formulaCount,
      sampleFormulas: formulasInSheet.slice(0, 5)
    });
  });

  console.log('\n--- 2. SHEET SUMMARIES ---');
  sheetSummaries.forEach(s => {
    console.log(`Sheet [${s.index}] "${s.name}": ${s.rowCount} rows, ${s.cellCount} cells, ${s.formulaCount} formulas`);
  });

  console.log('\n--- 3. SAMPLE FORMULA CELLS DETAILED AUDIT ---');
  console.log(JSON.stringify(formulaCellsSample.slice(0, 5), null, 2));

  fs.writeFileSync('real-file-audit.json', JSON.stringify({
    file: filePath,
    sizeBytes: stat.size,
    sheetsCount: workbook.worksheets.length,
    sheets: sheetSummaries,
    sampleFormulas: formulaCellsSample
  }, null, 2));

  console.log('\nSaved real-file-audit.json');
}

auditRealExcelFile().catch(console.error);
