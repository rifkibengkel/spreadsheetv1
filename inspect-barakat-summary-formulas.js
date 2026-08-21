const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Inspection of Summary Sheet Formulas in report barakat_07-03-07.xlsx ===');

  const srcPath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat_07-03-07.xlsx';
  const directory = await unzipper.Open.file(srcPath);

  const sheet1File = directory.files.find(f => f.path === 'xl/worksheets/sheet1.xml');
  const xmlStr = (await sheet1File.buffer()).toString('utf8');

  // Find 20 formula cells from sheet 1 (Summary sheet)
  const formulaRegex = /<c r="([A-Z0-9]+)"[^>]*>(?:<f([^>]*)>([^<]*)<\/f>)?(?:<v>([^<]*)<\/v>)?/g;
  let m;

  const samples = [];
  while ((m = formulaRegex.exec(xmlStr)) !== null) {
    if (m[3]) {
      samples.push({
        cell: m[1],
        attrs: m[2],
        formula: m[3],
        value: m[4]
      });
      if (samples.length >= 25) break;
    }
  }

  console.log('Summary Sheet Formula Samples from Source XLSX:');
  console.log(JSON.stringify(samples, null, 2));

  fs.writeFileSync('barakat-summary-formulas-samples.json', JSON.stringify(samples, null, 2));
})();
