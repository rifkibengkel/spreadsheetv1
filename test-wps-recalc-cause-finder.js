const fs = require('fs');
const ExcelJS = require('exceljs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Forensic Cause Finder: WPS Formula Recalculation to 0 ===');

  // Test Case 1: Standard formula vs Array Formula vs Cross Sheet Formula in ExcelJS
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Univer Spreadsheets App';

  const wsData = wb.addWorksheet('Data All');
  wsData.getCell('A1').value = 'Category';
  wsData.getCell('B1').value = 'Value';

  wsData.getCell('A2').value = 'X';
  wsData.getCell('B2').value = 100;

  wsData.getCell('A3').value = 'X';
  wsData.getCell('B3').value = 200;

  wsData.getCell('A4').value = 'Y';
  wsData.getCell('B4').value = 300;

  const wsSummary = wb.addWorksheet('Summary');

  // 1. Standard cross-sheet SUM
  wsSummary.getCell('A1').value = 'Standard SUM';
  wsSummary.getCell('B1').value = {
    formula: "SUM('Data All'!B2:B4)",
    result: 600
  };

  // 2. Unquoted cross-sheet SUM (Potential Bug)
  wsSummary.getCell('A2').value = 'Unquoted SUM';
  wsSummary.getCell('B2').value = {
    formula: "SUM(Data All!B2:B4)",
    result: 600
  };

  // 3. Array Formula without shareType (Potential Bug)
  wsSummary.getCell('A3').value = 'Array Formula (No shareType)';
  wsSummary.getCell('B3').value = {
    formula: "SUM(IF('Data All'!A2:A4=\"X\", 'Data All'!B2:B4, 0))",
    result: 300
  };

  // 4. Array Formula WITH shareType: 'array'
  wsSummary.getCell('A4').value = 'Array Formula (WITH shareType)';
  wsSummary.getCell('B4').value = {
    formula: "SUM(IF('Data All'!A2:A4=\"X\", 'Data All'!B2:B4, 0))",
    result: 300,
    shareType: 'array',
    ref: 'B4'
  };

  const buffer = await wb.xlsx.writeBuffer();
  fs.writeFileSync('wps-recalc-test-cases.xlsx', buffer);
  console.log('Saved wps-recalc-test-cases.xlsx. Inspecting generated XML...');

  const directory = await unzipper.Open.file('wps-recalc-test-cases.xlsx');
  for (const file of directory.files) {
    if (file.path.startsWith('xl/worksheets/sheet2') || file.path === 'xl/workbook.xml') {
      const content = await file.buffer();
      console.log(`\n=== FILE: ${file.path} ===`);
      console.log(content.toString('utf8'));
    }
  }
})();
