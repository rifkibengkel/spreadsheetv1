const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Forensic Inspection of Failing WPS Formulas in test-exported.xlsx ===');

  const directory = await unzipper.Open.file('test-exported.xlsx');

  const sheetFiles = directory.files.filter(f => f.path.startsWith('xl/worksheets/sheet') && f.path.endsWith('.xml'));

  for (const file of sheetFiles) {
    const content = await file.buffer();
    const xmlStr = content.toString('utf8');
    console.log(`\n=== Sheet: ${file.path} ===`);
    console.log(xmlStr);
  }
})();
