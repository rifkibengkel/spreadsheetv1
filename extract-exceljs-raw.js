const fs = require('fs');
const ExcelJS = require('exceljs');

async function extractRawExcel() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const summary = workbook.getWorksheet('Summary');
  const summaryHadiah = workbook.getWorksheet('Summary Hadiah');
  
  const getCellDetails = (ws, cellAddr) => {
    const cell = ws.getCell(cellAddr);
    return {
      address: cellAddr,
      value: cell.value,
      result: cell.result,
      type: cell.type,
      formula: cell.formula || (cell.value && cell.value.formula)
    };
  };

  const report = {
    Summary_I5: getCellDetails(summary, 'I5'),
    Summary_I6: getCellDetails(summary, 'I6'),
    Summary_I7: getCellDetails(summary, 'I7'),
    SummaryHadiah_C4: getCellDetails(summaryHadiah, 'C4'),
    SummaryHadiah_B4: getCellDetails(summaryHadiah, 'B4')
  };

  // Mathematical Proof for Summary Hadiah!C4
  const c4Cell = summaryHadiah.getCell('C4');
  const b4Val = summaryHadiah.getCell('B4').value;

  let matchCount = 0;
  let sumTotal = 0;
  const matches = [];

  for (let r = 4; r <= 188; r++) {
    const codeVal = summary.getCell(`A${r}`).value;
    const numVal = summary.getCell(`C${r}`).value;

    const codeStr = typeof codeVal === 'object' && codeVal !== null ? (codeVal.result || codeVal.text) : codeVal;
    const num = typeof numVal === 'object' && numVal !== null ? (numVal.result || 0) : (Number(numVal) || 0);

    if (String(codeStr) === String(b4Val)) {
      matchCount++;
      sumTotal += num;
      matches.push({ row: r, code: codeStr, num: num });
    }
  }

  report.MathematicalProof_SummaryHadiah_C4 = {
    formula: c4Cell.formula || c4Cell.value?.formula,
    criteria_B4: b4Val,
    matchingRowsCount: matchCount,
    sumTotal: sumTotal,
    matches: matches,
    isMathematicallyZero: (sumTotal === 0)
  };

  fs.writeFileSync('raw-excel-details.json', JSON.stringify(report, null, 2));
  console.log('Saved raw-excel-details.json successfully.');
}

extractRawExcel().catch(console.error);
