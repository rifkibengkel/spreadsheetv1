const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  const directory = await unzipper.Open.file('wps-recalc-test-cases.xlsx');
  let text = '';
  for (const file of directory.files) {
    if (file.path === 'xl/worksheets/sheet2.xml' || file.path === 'xl/workbook.xml') {
      const content = await file.buffer();
      text += `\n=== FILE: ${file.path} ===\n` + content.toString('utf8') + '\n';
    }
  }
  fs.writeFileSync('wps_xml_utf8.txt', text, 'utf8');
})();
