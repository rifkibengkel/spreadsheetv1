const fs = require('fs');
const unzipper = require('unzipper');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

(async () => {
  console.log('=== REAL CELL VALUE COMPATIBILITY TEST (RESOLVING SHARED STRINGS) ===');

  const srcPath = 'C:\\Users\\BKKI-1\\Documents\\report barakat.xlsx';
  const destPath = 'C:\\Users\\BKKI-1\\Desktop\\report_barakat_exported_test.xlsx';

  async function loadWorkbookCellMap(filePath) {
    const zip = await unzipper.Open.file(filePath);

    // Read shared strings
    const sstFile = zip.files.find(f => f.path === 'xl/sharedStrings.xml');
    let sharedStrings = [];
    if (sstFile) {
      const sstXml = (await sstFile.buffer()).toString('utf8');
      sharedStrings = [...sstXml.matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map(m => m[1]);
    }

    const wbFile = zip.files.find(f => f.path === 'xl/workbook.xml');
    const wbXml = (await wbFile.buffer()).toString('utf8');
    const sheetMatches = [...wbXml.matchAll(/<sheet[^>]+name="([^"]+)"[^>]+r:id="([^"]+)"/g)];

    const sheetFiles = zip.files.filter(f => f.path.startsWith('xl/worksheets/sheet') && f.path.endsWith('.xml'));

    const workbookMap = {};

    for (let i = 0; i < sheetFiles.length; i++) {
      const file = sheetFiles[i];
      const sheetName = sheetMatches[i] ? sheetMatches[i][1] : `Sheet${i+1}`;

      const content = await file.buffer();
      const xmlStr = content.toString('utf8');

      workbookMap[sheetName] = {};

      const cellRegex = /<c r="([A-Z0-9]+)"([^>]*)>(?:<f([^>]*)>([^<]*)<\/f>)?(?:<v>([^<]*)<\/v>)?/g;
      let m;

      while ((m = cellRegex.exec(xmlStr)) !== null) {
        const cellRef = m[1];
        const attrs = m[2];
        const formula = m[4];
        const rawVal = m[5];

        let actualVal = rawVal;
        if (attrs.includes('t="s"') && rawVal !== undefined) {
          const idx = Number(rawVal);
          actualVal = sharedStrings[idx] !== undefined ? sharedStrings[idx] : rawVal;
        } else if (attrs.includes('t="str"') && rawVal !== undefined) {
          actualVal = rawVal;
        }

        workbookMap[sheetName][cellRef] = {
          formula: formula || undefined,
          value: actualVal
        };
      }
    }

    return workbookMap;
  }

  console.log('\nLoading Original File (Resolving SharedStrings)...');
  const srcMap = await loadWorkbookCellMap(srcPath);

  console.log('Loading Exported File (Resolving SharedStrings)...');
  const expMap = await loadWorkbookCellMap(destPath);

  let totalTested = 0;
  let totalPass = 0;
  let totalFail = 0;

  const comparisonRows = [];
  const perSheetAudit = {};

  Object.keys(srcMap).forEach(sheetName => {
    const srcSheet = srcMap[sheetName] || {};
    const expSheet = expMap[sheetName] || {};

    perSheetAudit[sheetName] = { pass: 0, fail: 0 };

    Object.keys(srcSheet).forEach(cellRef => {
      const srcCell = srcSheet[cellRef];
      const expCell = expSheet[cellRef];

      totalTested++;

      let status = 'PASS';

      if (!expCell) {
        status = 'FAIL (Missing in Export)';
      } else if (srcCell.formula && !expCell.formula) {
        status = 'FAIL (Formula Lost)';
      } else if (srcCell.value !== undefined && expCell.value === undefined) {
        status = 'FAIL (Value Lost)';
      } else if (srcCell.value !== undefined && expCell.value !== undefined && String(srcCell.value) !== String(expCell.value)) {
        status = 'FAIL (Value Mismatch)';
      }

      if (status === 'PASS') {
        totalPass++;
        perSheetAudit[sheetName].pass++;
      } else {
        totalFail++;
        perSheetAudit[sheetName].fail++;
      }

      if (comparisonRows.length < 50 || status.startsWith('FAIL')) {
        comparisonRows.push({
          Sheet: sheetName,
          Cell: cellRef,
          OriginalWPS: srcCell.value !== undefined ? String(srcCell.value) : (srcCell.formula ? `[F: ${srcCell.formula}]` : 'Empty'),
          ExportedInWPS: expCell ? (expCell.value !== undefined ? String(expCell.value) : (expCell.formula ? `[F: ${expCell.formula}]` : 'Empty')) : 'MISSING',
          Status: status
        });
      }
    });
  });

  const finalSummary = {
    totalTested,
    totalPass,
    totalFail,
    passPercentage: ((totalPass / totalTested) * 100).toFixed(2) + '%',
    perSheetAudit
  };

  console.log('\n=== REAL CELL VALUE COMPATIBILITY RESULTS ===');
  console.log(JSON.stringify(finalSummary, null, 2));

  const sampleTable = comparisonRows.slice(0, 50);
  fs.writeFileSync('barakat-real-compatibility-report.json', JSON.stringify({ finalSummary, sampleTable }, null, 2), 'utf8');
})();
