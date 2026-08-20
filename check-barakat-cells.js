const fs = require('fs');
const ExcelJS = require('exceljs');

async function checkBarakatCells() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  const buf = fs.readFileSync(filePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);

  const logs = [];

  logs.push('--- Summary Sheet Cells (Row 1) ---');
  const wsSum = workbook.getWorksheet('Summary');
  wsSum.getRow(1).eachCell((cell, col) => {
    logs.push(`Summary col ${col} (${cell.address}): type=${cell.type}, val=${JSON.stringify(cell.value)}`);
  });

  logs.push('--- Entries data Sheet Cells (Row 1 & 2) ---');
  const wsEnt = workbook.getWorksheet('Entries data');
  wsEnt.getRow(1).eachCell((cell, col) => {
    logs.push(`Entries row 1 col ${col} (${cell.address}): type=${cell.type}, val=${JSON.stringify(cell.value)}`);
  });
  wsEnt.getRow(2).eachCell((cell, col) => {
    logs.push(`Entries row 2 col ${col} (${cell.address}): type=${cell.type}, val=${JSON.stringify(cell.value)}`);
  });

  fs.writeFileSync('check-barakat-cells.txt', logs.join('\n'));
  console.log('Saved check-barakat-cells.txt');
}

checkBarakatCells().catch(console.error);
