const fs = require('fs');
const ExcelJS = require('exceljs');

async function inspectValues() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  const buf = fs.readFileSync(filePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);

  const cellsToInspect = [
    { sheet: 'Summary', cell: 'A1' },
    { sheet: 'Summary Hadiah', cell: 'A1' },
    { sheet: 'Summary', cell: 'B4' },
    { sheet: 'Summary Hadiah', cell: 'B4' },
    { sheet: 'Entries data', cell: 'B2' },
    { sheet: 'Entries data', cell: 'M2' },
    { sheet: 'Summary', cell: 'C4' },
    { sheet: 'Summary Hadiah', cell: 'C4' }
  ];

  const out = cellsToInspect.map(item => {
    const ws = workbook.getWorksheet(item.sheet);
    if (!ws) return { ...item, error: 'SHEET NOT FOUND' };
    const cell = ws.getCell(item.cell);
    return {
      sheet: item.sheet,
      cell: item.cell,
      type: cell.type,
      value: cell.value,
      result: cell.result
    };
  });

  fs.writeFileSync('raw-excel-details.json', JSON.stringify(out, null, 2));
  console.log('Saved raw-excel-details.json');
}

inspectValues().catch(console.error);
