const fs = require('fs');
const ExcelJS = require('exceljs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Validating XLSX Export Fix for WPS Recalculation ===');

  const exportWb = new ExcelJS.Workbook();
  exportWb.creator = 'Univer Spreadsheets App';
  exportWb.lastModifiedBy = 'Univer Spreadsheets App';
  exportWb.calcProperties.fullCalcOnLoad = true;

  const wsData = exportWb.addWorksheet('Data All');
  wsData.getCell('A1').value = 'Category';
  wsData.getCell('B1').value = 'Value';

  wsData.getCell('A2').value = 'X';
  wsData.getCell('B2').value = 100;

  wsData.getCell('A3').value = 'X';
  wsData.getCell('B3').value = 200;

  const wsSummary = exportWb.addWorksheet('Summary');

  // Test cell 1: Unquoted sheet ref formula
  const unquotedFormula = "SUM(Data All!B2:B3)";
  const sanitized1 = unquotedFormula.replace(/(?<!')\b([A-Za-z0-9_]+(?:\s+[A-Za-z0-9_]+)+)\b(?!')!/g, "'$1'!");

  wsSummary.getCell('A1').value = 'Sanitized Cross-Sheet';
  wsSummary.getCell('B1').value = {
    formula: sanitized1,
    result: 300
  };

  // Test cell 2: Array formula with shareType: 'array'
  const arrayFormulaStr = "SUM(IF('Data All'!A2:A3=\"X\", 'Data All'!B2:B3, 0))";
  const sanitized2 = arrayFormulaStr.replace(/(?<!')\b([A-Za-z0-9_]+(?:\s+[A-Za-z0-9_]+)+)\b(?!')!/g, "'$1'!");

  const isArrayFormula = /SUM\s*\(\s*IF\b/i.test(sanitized2);

  wsSummary.getCell('A2').value = 'Array Formula Fixed';
  const cellObj = {
    formula: sanitized2,
    result: 300
  };

  if (isArrayFormula) {
    cellObj.shareType = 'array';
    cellObj.ref = 'B2';
  }
  wsSummary.getCell('B2').value = cellObj;

  const buffer = await exportWb.xlsx.writeBuffer();
  fs.writeFileSync('wps-export-fixed-test.xlsx', buffer);

  console.log('Saved wps-export-fixed-test.xlsx. Inspecting generated XML...');

  const directory = await unzipper.Open.file('wps-export-fixed-test.xlsx');
  let xmlReport = '';
  for (const file of directory.files) {
    if (file.path === 'xl/worksheets/sheet2.xml' || file.path === 'xl/workbook.xml') {
      const content = await file.buffer();
      xmlReport += `\n=== ${file.path} ===\n` + content.toString('utf8') + '\n';
    }
  }

  console.log(xmlReport);
  fs.writeFileSync('wps-fix-xml-report.txt', xmlReport, 'utf8');
})();
