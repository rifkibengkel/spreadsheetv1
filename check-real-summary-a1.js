const fs = require('fs');
const ExcelJS = require('exceljs');

async function checkRealSummaryA1() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  const buf = fs.readFileSync(filePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);

  const logs = [];

  const ws = workbook.getWorksheet('Summary');
  logs.push('--- Summary Sheet Row 1 Cells ---');
  for (let c = 1; c <= 10; c++) {
    const cell = ws.getCell(1, c);
    logs.push(`Summary Row 1 Col ${c} (${cell.address}): type=${cell.type}, val=${JSON.stringify(cell.value)}, result=${cell.result}, formula=${cell.formula}`);
  }

  const wsHad = workbook.getWorksheet('Summary Hadiah');
  logs.push('--- Summary Hadiah Sheet Row 1 Cells ---');
  for (let c = 1; c <= 10; c++) {
    const cell = wsHad.getCell(1, c);
    logs.push(`Summary Hadiah Row 1 Col ${c} (${cell.address}): type=${cell.type}, val=${JSON.stringify(cell.value)}, result=${cell.result}, formula=${cell.formula}`);
  }

  fs.writeFileSync('real-summary-a1.txt', logs.join('\n'));
  console.log('Saved real-summary-a1.txt');
}

checkRealSummaryA1().catch(console.error);
