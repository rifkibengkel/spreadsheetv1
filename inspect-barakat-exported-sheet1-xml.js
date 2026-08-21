const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Inspecting Exported Sheet 1 XML for barakat_exported.xlsx ===');

  const destPath = 'C:\\Users\\BKKI-1\\Desktop\\barakat_exported.xlsx';
  const directory = await unzipper.Open.file(destPath);

  const sheet1File = directory.files.find(f => f.path === 'xl/worksheets/sheet1.xml');
  const xmlStr = (await sheet1File.buffer()).toString('utf8');

  // Inspect cells O4, T4, U4, V4, W4, Z4, AF4
  const targetCells = ['O4', 'T4', 'U4', 'V4', 'W4', 'Z4', 'AF4', 'AG4', 'AK4', 'AL4'];

  const results = [];

  targetCells.forEach(cellRef => {
    const regex = new RegExp(`<c r="${cellRef}"[^>]*>(?:<f([^>]*)>([^<]*)<\\/f>)?(?:<v>([^<]*)<\\/v>)?`, 'g');
    const m = regex.exec(xmlStr);
    if (m) {
      results.push({
        cell: cellRef,
        fAttrs: m[1],
        formula: m[2],
        value: m[3]
      });
    } else {
      results.push({ cell: cellRef, status: 'NOT FOUND IN XML' });
    }
  });

  console.log('Exported XML Cell Data for Target Cells:');
  console.log(JSON.stringify(results, null, 2));

  fs.writeFileSync('barakat-exported-target-cells.json', JSON.stringify(results, null, 2));
})();
