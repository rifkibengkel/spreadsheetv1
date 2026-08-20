const ExcelJS = require('exceljs');
const fs = require('fs');

async function streamInspectBarakat() {
  console.log('=== STREAMING INSPECTION OF report barakat.xlsx ===');
  const filePath = 'c:\\Users\\BKKI-1\\.gemini\\antigravity\\scratch\\sheet-V-2\\barakat-forensic.xlsx';

  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, {
    entries: 'emit',
    sharedStrings: 'cache',
    styles: 'cache',
    hyperlinks: 'ignore'
  });

  const cellsOfInterest = {
    'Summary Hadiah!A1': null,
    'Summary Hadiah!C4': null,
    'Summary Hadiah!C5': null,
    'Summary!I9': null,
    'Summary!C4': null
  };

  for await (const worksheetReader of workbookReader) {
    const sheetName = worksheetReader.name;
    console.log(`Processing worksheet: ${sheetName}`);

    for await (const row of worksheetReader) {
      row.eachCell((cell, colNumber) => {
        const key = `${sheetName}!${cell.address}`;
        if (cellsOfInterest.hasOwnProperty(key)) {
          cellsOfInterest[key] = {
            sheet: sheetName,
            address: cell.address,
            type: cell.type,
            formula: cell.formula,
            result: cell.result,
            value: cell.value
          };
          console.log(`Found target cell ${key}:`, cellsOfInterest[key]);
        }
      });
    }
  }

  console.log('\nFinal Extracted Cells from report barakat.xlsx:');
  console.log(JSON.stringify(cellsOfInterest, null, 2));
}

streamInspectBarakat().catch(console.error);
