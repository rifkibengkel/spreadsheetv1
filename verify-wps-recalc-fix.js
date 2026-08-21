const fs = require('fs');
const ExcelJS = require('exceljs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Final Forensic Verification of WPS Recalculation Fix ===');

  const snapshot = {
    sheetOrder: ['sheet-data', 'sheet-summary'],
    sheets: {
      'sheet-data': {
        name: 'Data All',
        cellData: {
          0: { 0: { v: 'Category' }, 1: { v: 'Value' } },
          1: { 0: { v: 'X' }, 1: { v: 100 } },
          2: { 0: { v: 'X' }, 1: { v: 200 } }
        }
      },
      'sheet-summary': {
        name: 'Summary Sheet',
        cellData: {
          0: {
            0: { v: 'Cross Sheet Unquoted' },
            1: { f: '=SUM(Data All!B2:B3)', v: 300 }
          },
          1: {
            0: { v: 'Array Formula SUM IF' },
            1: { f: '=SUM(IF(Data All!A2:A3="X", Data All!B2:B3, 0))', v: 300 }
          }
        }
      }
    }
  };

  // Simulate export worker logic
  const exportWb = new ExcelJS.Workbook();
  exportWb.creator = 'Univer Spreadsheets App';
  exportWb.lastModifiedBy = 'Univer Spreadsheets App';
  exportWb.calcProperties.fullCalcOnLoad = true;

  const sortedSheetIds = snapshot.sheetOrder;

  for (let sIdx = 0; sIdx < sortedSheetIds.length; sIdx++) {
    const sheetId = sortedSheetIds[sIdx];
    const sheetData = snapshot.sheets[sheetId];
    if (!sheetData) continue;

    const ws = exportWb.addWorksheet(sheetData.name || sheetId);
    const cellDataRaw = sheetData.cellData || {};
    const rowKeys = Object.keys(cellDataRaw).map(Number).sort((a, b) => a - b);

    for (let rIdx = 0; rIdx < rowKeys.length; rIdx++) {
      const rowIndex = rowKeys[rIdx];
      const cols = cellDataRaw[rowIndex];
      if (cols) {
        const colKeys = Object.keys(cols).map(Number).sort((a, b) => a - b);
        for (let cIdx = 0; cIdx < colKeys.length; cIdx++) {
          const colIndex = colKeys[cIdx];
          const cellPay = cols[colIndex];
          if (!cellPay) continue;

          const cell = ws.getCell(rowIndex + 1, colIndex + 1);
          if (cellPay.f) {
            let parsedFormula = cellPay.f;
            if (parsedFormula.startsWith('=')) {
              parsedFormula = parsedFormula.substring(1);
            }
            parsedFormula = parsedFormula.replace(/(?<!')\b([A-Za-z0-9_]+(?:\s+[A-Za-z0-9_]+)+)\b(?!')!/g, "'$1'!");

            const isArrayFormula = cellPay.t === 1 || 
              /SUM\s*\(\s*IF\b|COUNT\s*\(\s*IF\b|AVERAGE\s*\(\s*IF\b|MODE\.MULT\b/i.test(parsedFormula) ||
              (parsedFormula.startsWith('{') && parsedFormula.endsWith('}'));

            if (parsedFormula.startsWith('{') && parsedFormula.endsWith('}')) {
              parsedFormula = parsedFormula.substring(1, parsedFormula.length - 1);
            }

            const formulaRes = cellPay.v !== undefined && cellPay.v !== null ? cellPay.v : undefined;

            const cellObj = {
              formula: parsedFormula,
              result: formulaRes,
            };

            if (isArrayFormula) {
              cellObj.shareType = 'array';
              cellObj.ref = cell.address;
            }

            cell.value = cellObj;
          } else if (cellPay.v !== undefined && cellPay.v !== null) {
            cell.value = cellPay.v;
          }
        }
      }
    }
  }

  const arrayBuffer = await exportWb.xlsx.writeBuffer();
  fs.writeFileSync('final-verified-wps-export.xlsx', arrayBuffer);
  console.log('Saved final-verified-wps-export.xlsx. Verifying XML content...');

  const directory = await unzipper.Open.file('final-verified-wps-export.xlsx');
  let workbookXml = '';
  let summarySheetXml = '';

  for (const file of directory.files) {
    if (file.path === 'xl/workbook.xml') {
      workbookXml = (await file.buffer()).toString('utf8');
    }
    if (file.path === 'xl/worksheets/sheet2.xml') {
      summarySheetXml = (await file.buffer()).toString('utf8');
    }
  }

  console.log('\n--- VERIFICATION CHECKLIST ---');
  console.log('1. calcPr fullCalcOnLoad="1":', workbookXml.includes('fullCalcOnLoad="1"') ? 'PASS [OK]' : 'FAIL');
  console.log('2. Quoted sheet names (\'Data All\'!):', summarySheetXml.includes("'Data All'!") ? 'PASS [OK]' : 'FAIL');
  console.log('3. Array formula tag (<f t="array" ref="B2">):', summarySheetXml.includes('t="array"') ? 'PASS [OK]' : 'FAIL');
  console.log('4. Cached value preserved (<v>300</v>):', summarySheetXml.includes('<v>300</v>') ? 'PASS [OK]' : 'FAIL');

  console.log('\n--- Summary Sheet XML Snippet ---');
  console.log(summarySheetXml);
})();
