const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== FORENSIC ANALYSIS OF 407 MISMATCHED CELLS IN C:\\Users\\BKKI-1\\Documents\\report barakat.xlsx ===');

  const srcPath = 'C:\\Users\\BKKI-1\\Documents\\report barakat.xlsx';
  const destPath = 'C:\\Users\\BKKI-1\\Desktop\\report_barakat_exported_test.xlsx';

  const srcZip = await unzipper.Open.file(srcPath);
  const expZip = await unzipper.Open.file(destPath);

  const wbFile = srcZip.files.find(f => f.path === 'xl/workbook.xml');
  const wbXml = (await wbFile.buffer()).toString('utf8');
  const sheetMatches = [...wbXml.matchAll(/<sheet[^>]+name="([^"]+)"[^>]+r:id="([^"]+)"/g)];

  const srcSheetFiles = srcZip.files.filter(f => f.path.startsWith('xl/worksheets/sheet') && f.path.endsWith('.xml'));
  const expSheetFiles = expZip.files.filter(f => f.path.startsWith('xl/worksheets/sheet') && f.path.endsWith('.xml'));

  const mismatches = [];
  const mismatchBySheet = {};

  for (let i = 0; i < srcSheetFiles.length; i++) {
    const srcFile = srcSheetFiles[i];
    const expFile = expSheetFiles[i];
    const sheetName = sheetMatches[i] ? sheetMatches[i][1] : `Sheet${i+1}`;

    mismatchBySheet[sheetName] = 0;

    const srcXml = (await srcFile.buffer()).toString('utf8');
    const expXml = (await expFile.buffer()).toString('utf8');

    const cellRegex = /<c r="([A-Z0-9]+)"[^>]*>(?:<f([^>]*)>([^<]*)<\/f>)?(?:<v>([^<]*)<\/v>)?/g;

    const srcMap = {};
    let m;
    while ((m = cellRegex.exec(srcXml)) !== null) {
      srcMap[m[1]] = { formula: m[3], value: m[4] };
    }

    const expMap = {};
    while ((m = cellRegex.exec(expXml)) !== null) {
      expMap[m[1]] = { formula: m[3], value: m[4] };
    }

    Object.keys(srcMap).forEach(cellRef => {
      const srcCell = srcMap[cellRef];
      const expCell = expMap[cellRef];

      if (!expCell) {
        mismatchBySheet[sheetName]++;
        mismatches.push({
          sheet: sheetName,
          cell: cellRef,
          type: 'MISSING_IN_EXPORT',
          srcFormula: srcCell.formula || 'None',
          srcValue: srcCell.value || 'None'
        });
      } else if (srcCell.formula && !expCell.formula) {
        mismatchBySheet[sheetName]++;
        mismatches.push({
          sheet: sheetName,
          cell: cellRef,
          type: 'FORMULA_LOST',
          srcFormula: srcCell.formula,
          expFormula: expCell.formula || 'None',
          srcValue: srcCell.value || 'None',
          expValue: expCell.value || 'None'
        });
      } else if (srcCell.value !== undefined && expCell.value === undefined) {
        mismatchBySheet[sheetName]++;
        mismatches.push({
          sheet: sheetName,
          cell: cellRef,
          type: 'VALUE_LOST',
          srcFormula: srcCell.formula || 'None',
          srcValue: srcCell.value,
          expValue: 'undefined'
        });
      } else if (srcCell.value !== undefined && expCell.value !== undefined && srcCell.value !== expCell.value) {
        mismatchBySheet[sheetName]++;
        mismatches.push({
          sheet: sheetName,
          cell: cellRef,
          type: 'VALUE_MISMATCH',
          srcFormula: srcCell.formula || 'None',
          srcValue: srcCell.value,
          expValue: expCell.value
        });
      }
    });
  }

  console.log('Mismatch Count per Sheet:', mismatchBySheet);
  console.log(`Total Mismatches Found: ${mismatches.length}`);

  const summaryObj = { mismatchBySheet, total: mismatches.length, top50Mismatches: mismatches.slice(0, 50) };
  fs.writeFileSync('barakat-mismatches-summary.json', JSON.stringify(summaryObj, null, 2), 'utf8');
  console.log('Saved barakat-mismatches-summary.json');
})();
