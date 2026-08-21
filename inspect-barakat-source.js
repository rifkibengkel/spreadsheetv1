const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Step 0: Inspecting Source File report barakat_07-03-07.xlsx ===');

  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat_07-03-07.xlsx';
  if (!fs.existsSync(filePath)) {
    console.error(`File NOT found: ${filePath}`);
    return;
  }

  const stats = fs.statSync(filePath);
  console.log(`Source File Size: ${(stats.size / (1024 * 1024)).toFixed(2)} MB`);

  const directory = await unzipper.Open.file(filePath);

  let workbookXml = '';
  const sheetFiles = [];

  for (const file of directory.files) {
    if (file.path === 'xl/workbook.xml') {
      workbookXml = (await file.buffer()).toString('utf8');
    }
    if (file.path.startsWith('xl/worksheets/sheet') && file.path.endsWith('.xml')) {
      sheetFiles.push(file);
    }
  }

  // Extract sheet names from workbook.xml
  const sheetNameMatches = [...workbookXml.matchAll(/<sheet[^>]*name="([^"]+)"[^>]*sheetId="([^"]+)"/g)];
  console.log(`Total Sheets in Workbook: ${sheetNameMatches.length}`);
  sheetNameMatches.forEach(m => console.log(`  Sheet [ID ${m[2]}]: "${m[1]}"`));

  let totalFormulas = 0;
  let totalValues = 0;

  const perSheetStats = [];

  for (const file of sheetFiles) {
    const content = await file.buffer();
    const xmlStr = content.toString('utf8');

    const formulas = (xmlStr.match(/<f[^>]*>/g) || []).length;
    const cached = (xmlStr.match(/<v>[^<]*<\/v>/g) || []).length;

    totalFormulas += formulas;
    totalValues += cached;

    perSheetStats.push({
      file: file.path,
      formulas,
      cachedValues: cached
    });
  }

  console.log('\nSource Workbook SAX Summary:');
  console.log(`- Total <f> Formulas: ${totalFormulas}`);
  console.log(`- Total <v> Cached Values: ${totalValues}`);
  console.log('- Per Sheet SAX Stats:', perSheetStats);
})();
