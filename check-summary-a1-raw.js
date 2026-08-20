const fs = require('fs');
const ExcelJS = require('exceljs');

async function checkSummaryA1Raw() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  const buf = fs.readFileSync(filePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);

  const logs = [];

  logs.push('--- ALL WORKBOOK SHEETS ---');
  workbook.worksheets.forEach(ws => logs.push(`Sheet name: ${ws.name}`));

  const wsSum = workbook.getWorksheet('Summary');
  logs.push('--- Summary Sheet Row 1 & Row 2 Cells ---');
  wsSum.getRow(1).eachCell((c, col) => logs.push(`Summary R1C${col} (${c.address}): type=${c.type}, val=${JSON.stringify(c.value)}, result=${c.result}, formula=${c.formula}`));
  wsSum.getRow(2).eachCell((c, col) => logs.push(`Summary R2C${col} (${c.address}): type=${c.type}, val=${JSON.stringify(c.value)}, result=${c.result}, formula=${c.formula}`));

  const wsHad = workbook.getWorksheet('Summary Hadiah');
  logs.push('--- Summary Hadiah Sheet Row 1 & Row 2 Cells ---');
  wsHad.getRow(1).eachCell((c, col) => logs.push(`Summary Hadiah R1C${col} (${c.address}): type=${c.type}, val=${JSON.stringify(c.value)}, result=${c.result}, formula=${c.formula}`));
  wsHad.getRow(2).eachCell((c, col) => logs.push(`Summary Hadiah R2C${col} (${c.address}): type=${c.type}, val=${JSON.stringify(c.value)}, result=${c.result}, formula=${c.formula}`));

  fs.writeFileSync('summary-a1-raw.txt', logs.join('\n'));
  console.log('Saved summary-a1-raw.txt');
}

checkSummaryA1Raw().catch(console.error);
