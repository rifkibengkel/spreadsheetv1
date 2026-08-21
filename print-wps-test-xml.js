const unzipper = require('unzipper');

(async () => {
  const directory = await unzipper.Open.file('wps-recalc-test-cases.xlsx');
  for (const file of directory.files) {
    if (file.path === 'xl/worksheets/sheet2.xml' || file.path === 'xl/workbook.xml') {
      const content = await file.buffer();
      console.log(`\n=== FILE: ${file.path} ===`);
      console.log(content.toString('utf8'));
    }
  }
})();
