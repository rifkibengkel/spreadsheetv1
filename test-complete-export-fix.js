const fs = require('fs');
const ExcelJS = require('exceljs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Testing Complete Shared Formula + Special Sheet Name + Array Export Fix ===');

  const sheetNames = [
    'Data All',
    'Data Valid',
    'Summary',
    'E-Voucher & E-Wallet',
    'Unik Konsumen'
  ];

  function sanitizeSheetNamesInFormula(formula, availableSheetNames) {
    if (!formula) return formula;
    let result = formula;
    for (const name of availableSheetNames) {
      if (/[\s&\-\.\/\\]/.test(name)) {
        const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const regex = new RegExp(`(?<!')${escapedName}(?!')!`, 'g');
        result = result.replace(regex, `'${name}'!`);
      }
    }
    return result;
  }

  const testFormulas = [
    "SUM(E-Voucher & E-Wallet!B2:B10)",
    "VLOOKUP(A1, Unik Konsumen!A:B, 2, 0)",
    "COUNTIF(Data Valid!O2:O100, B4)"
  ];

  console.log('\n--- Sheet Name Sanitization Test ---');
  testFormulas.forEach(f => {
    console.log('Original :', f);
    console.log('Sanitized:', sanitizeSheetNamesInFormula(f, sheetNames));
    console.log('---');
  });

  // Create ExcelJS workbook with master and slave shared formulas
  const exportWb = new ExcelJS.Workbook();
  exportWb.creator = 'Univer Spreadsheets App';
  exportWb.calcProperties.fullCalcOnLoad = true;

  const ws = exportWb.addWorksheet('E-Voucher & E-Wallet');

  const sharedMap = new Map();
  // Master cell at row 1 (A1)
  sharedMap.set('si_0', 'DATE(YEAR(B1),MONTH(B1),DAY(B1))');

  // Master cell
  ws.getCell('A1').value = {
    formula: sharedMap.get('si_0'),
    result: '2025-01-01'
  };

  // Slave cell at row 2 (A2) using sharedMap
  ws.getCell('A2').value = {
    formula: sharedMap.get('si_0'),
    result: '2025-01-02'
  };

  const buffer = await exportWb.xlsx.writeBuffer();
  fs.writeFileSync('complete-fix-test.xlsx', buffer);

  const directory = await unzipper.Open.file('complete-fix-test.xlsx');
  for (const file of directory.files) {
    if (file.path.startsWith('xl/worksheets/sheet')) {
      const content = await file.buffer();
      console.log(`\n=== FILE: ${file.path} ===`);
      console.log(content.toString('utf8'));
    }
  }
})();
