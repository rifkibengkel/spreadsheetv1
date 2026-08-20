const fs = require('fs');
const ExcelJS = require('exceljs');

(async () => {
  console.log('=== Offline Forensic Inspection of updatehexos.xlsx Formulas ===');

  const workbook = new ExcelJS.Workbook();
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx';

  console.log('Loading workbook...');
  await workbook.xlsx.readFile(filePath);

  console.log(`Workbook loaded. Worksheets count: ${workbook.worksheets.length}`);

  const sheetStats = [];
  const formulaTypeCounts = {};
  let totalFormulas = 0;
  let fullColumnFormulas = 0;

  workbook.worksheets.forEach(ws => {
    let sheetFormulas = 0;
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (cell.formula || (cell.value && typeof cell.value === 'object' && cell.value.formula)) {
          sheetFormulas++;
          totalFormulas++;

          const fText = cell.formula || cell.value.formula || '';

          // Extract function names (e.g., SUM, VLOOKUP, SUMIFS)
          const funcMatches = fText.match(/[A-Z0-9\._]+(?=\()/gi);
          if (funcMatches) {
            funcMatches.forEach(fn => {
              const upperFn = fn.toUpperCase();
              formulaTypeCounts[upperFn] = (formulaTypeCounts[upperFn] || 0) + 1;
            });
          }

          // Check for full column / full row references (e.g. A:A, Summary!A:Z, 1:100000)
          if (/[A-Z]+:[A-Z]+/i.test(fText)) {
            fullColumnFormulas++;
          }
        }
      });
    });

    sheetStats.push({
      name: ws.name,
      rowCount: ws.rowCount,
      columnCount: ws.columnCount,
      formulaCount: sheetFormulas
    });
  });

  const topFunctions = Object.entries(formulaTypeCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 20);

  const report = {
    totalFormulas,
    fullColumnFormulas,
    sheetStats,
    topFunctions
  };

  console.log('Inspection Summary:', JSON.stringify(report, null, 2));
  fs.writeFileSync('updatehexos-formula-structure.json', JSON.stringify(report, null, 2));
})();
