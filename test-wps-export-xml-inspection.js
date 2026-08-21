const fs = require('fs');
const ExcelJS = require('exceljs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Forensic Inspection of Exported XLSX XML Structure for WPS Recalculation ===');

  // 1. Create a test workbook with various formula types
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Univer Spreadsheets App';

  const ws1 = wb.addWorksheet('Data Sheet');
  ws1.getCell('A1').value = 'Item';
  ws1.getCell('B1').value = 'Amount';
  ws1.getCell('A2').value = 'Apple';
  ws1.getCell('B2').value = 10;
  ws1.getCell('A3').value = 'Banana';
  ws1.getCell('B3').value = 20;

  const ws2 = wb.addWorksheet('Summary');
  ws2.getCell('A1').value = 'Total';
  ws2.getCell('B1').value = { formula: "SUM('Data Sheet'!B2:B3)", result: 30 };
  ws2.getCell('A2').value = 'Count';
  ws2.getCell('B2').value = { formula: "COUNTIF('Data Sheet'!A2:A3, \"Apple\")", result: 1 };
  ws2.getCell('A3').value = 'Lookup';
  ws2.getCell('B3').value = { formula: "VLOOKUP(\"Banana\", 'Data Sheet'!A2:B3, 2, FALSE)", result: 20 };

  const buffer = await wb.xlsx.writeBuffer();
  fs.writeFileSync('test-exported.xlsx', buffer);

  console.log('Saved test-exported.xlsx. Unzipping XML to inspect structure...');

  const directory = await unzipper.Open.file('test-exported.xlsx');

  for (const file of directory.files) {
    if (file.path.startsWith('xl/') && (file.path.endsWith('.xml') || file.path.endsWith('.rels'))) {
      const content = await file.buffer();
      console.log(`\n--- FILE: ${file.path} ---`);
      console.log(content.toString('utf8').slice(0, 1500));
    }
  }
})();
