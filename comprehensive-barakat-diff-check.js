const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Deep Cell-by-Cell Diff Audit for report barakat_07-03-07.xlsx ===');

  const exportedPath = 'C:\\Users\\BKKI-1\\Desktop\\barakat_exported.xlsx';
  const sourcePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat_07-03-07.xlsx';

  console.log('Unzipping source XLSX...');
  const srcZip = await unzipper.Open.file(sourcePath);
  console.log('Unzipping exported XLSX...');
  const expZip = await unzipper.Open.file(exportedPath);

  const srcSheets = srcZip.files.filter(f => f.path.startsWith('xl/worksheets/sheet') && f.path.endsWith('.xml'));
  const expSheets = expZip.files.filter(f => f.path.startsWith('xl/worksheets/sheet') && f.path.endsWith('.xml'));

  console.log(`Source Sheets Count: ${srcSheets.length}, Exported Sheets Count: ${expSheets.length}`);

  let totalSrcCells = 0;
  let totalExpCells = 0;
  let totalSrcFormulas = 0;
  let totalExpFormulas = 0;
  let totalSrcValues = 0;
  let totalExpValues = 0;

  const perSheetDiff = [];

  for (let i = 0; i < srcSheets.length; i++) {
    const srcFile = srcSheets[i];
    const expFile = expSheets[i];

    const srcXml = (await srcFile.buffer()).toString('utf8');
    const expXml = (await expFile.buffer()).toString('utf8');

    const srcFormulas = (srcXml.match(/<f[^>]*>/g) || []).length;
    const expFormulas = (expXml.match(/<f[^>]*>/g) || []).length;

    const srcValues = (srcXml.match(/<v>[^<]*<\/v>/g) || []).length;
    const expValues = (expXml.match(/<v>[^<]*<\/v>/g) || []).length;

    const srcCells = (srcXml.match(/<c r="[A-Z0-9]+"/g) || []).length;
    const expCells = (expXml.match(/<c r="[A-Z0-9]+"/g) || []).length;

    totalSrcFormulas += srcFormulas;
    totalExpFormulas += expFormulas;
    totalSrcValues += srcValues;
    totalExpValues += expValues;
    totalSrcCells += srcCells;
    totalExpCells += expCells;

    perSheetDiff.push({
      sheet: srcFile.path,
      srcCells,
      expCells,
      cellDiff: expCells - srcCells,
      srcFormulas,
      expFormulas,
      formulaDiff: expFormulas - srcFormulas,
      srcValues,
      expValues,
      valueDiff: expValues - srcValues
    });
  }

  const finalDiffReport = {
    totalSrcCells,
    totalExpCells,
    totalSrcFormulas,
    totalExpFormulas,
    totalSrcValues,
    totalExpValues,
    perSheetDiff
  };

  console.log('\n=== COMPLETE SOURCE vs EXPORTED METRICS REPORT ===');
  console.log(JSON.stringify(finalDiffReport, null, 2));

  fs.writeFileSync('barakat-complete-diff-report.json', JSON.stringify(finalDiffReport, null, 2));
})();
