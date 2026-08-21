const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Step 1: Inspecting Original File C:\\Users\\BKKI-1\\Documents\\report barakat.xlsx ===');

  const filePath = 'C:\\Users\\BKKI-1\\Documents\\report barakat.xlsx';
  if (!fs.existsSync(filePath)) {
    console.error(`File NOT found: ${filePath}`);
    return;
  }

  const stats = fs.statSync(filePath);
  console.log(`Source File Size: ${(stats.size / (1024 * 1024)).toFixed(2)} MB`);

  const directory = await unzipper.Open.file(filePath);

  const workbookFile = directory.files.find(f => f.path === 'xl/workbook.xml');
  const wbXml = (await workbookFile.buffer()).toString('utf8');

  const sheetMatches = [...wbXml.matchAll(/<sheet[^>]+name="([^"]+)"[^>]+r:id="([^"]+)"/g)];
  console.log(`Total Sheets: ${sheetMatches.length}`);
  sheetMatches.forEach(m => console.log(`  Sheet: "${m[1]}"`));

  let totalFormulas = 0;
  let totalValues = 0;
  let totalCells = 0;

  const sheetFiles = directory.files.filter(f => f.path.startsWith('xl/worksheets/sheet') && f.path.endsWith('.xml'));

  const perSheetStats = [];

  for (let i = 0; i < sheetFiles.length; i++) {
    const file = sheetFiles[i];
    const sheetName = sheetMatches[i] ? sheetMatches[i][1] : `Sheet${i+1}`;

    const content = await file.buffer();
    const xmlStr = content.toString('utf8');

    const formulas = (xmlStr.match(/<f[^>]*>/g) || []).length;
    const cached = (xmlStr.match(/<v>[^<]*<\/v>/g) || []).length;
    const cells = (xmlStr.match(/<c r="[A-Z0-9]+"/g) || []).length;

    totalFormulas += formulas;
    totalValues += cached;
    totalCells += cells;

    perSheetStats.push({
      sheet: sheetName,
      file: file.path,
      cells,
      formulas,
      cachedValues: cached
    });
  }

  const summaryReport = {
    filePath,
    fileSizeMB: (stats.size / (1024 * 1024)).toFixed(2),
    totalSheets: sheetMatches.length,
    totalCells,
    totalFormulas,
    totalValues,
    perSheetStats
  };

  console.log('\nOriginal File Summary:');
  console.log(JSON.stringify(summaryReport, null, 2));

  fs.writeFileSync('documents-barakat-original-summary.json', JSON.stringify(summaryReport, null, 2));
})();
