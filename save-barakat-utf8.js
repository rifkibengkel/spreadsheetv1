const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat_07-03-07.xlsx';
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

  const sheetNameMatches = [...workbookXml.matchAll(/<sheet[^>]*name="([^"]+)"[^>]*sheetId="([^"]+)"/g)];

  let reportStr = `=== Source File: report barakat_07-03-07.xlsx ===\n`;
  reportStr += `Total Sheets: ${sheetNameMatches.length}\n`;
  sheetNameMatches.forEach(m => reportStr += `  Sheet [ID ${m[2]}]: "${m[1]}"\n`);

  let totalFormulas = 0;
  let totalValues = 0;

  for (const file of sheetFiles) {
    const content = await file.buffer();
    const xmlStr = content.toString('utf8');

    const formulas = (xmlStr.match(/<f[^>]*>/g) || []).length;
    const cached = (xmlStr.match(/<v>[^<]*<\/v>/g) || []).length;

    totalFormulas += formulas;
    totalValues += cached;

    reportStr += `\n${file.path}:\n  Formulas: ${formulas}\n  Cached Values: ${cached}\n`;
  }

  reportStr += `\nTOTAL FORMULAS: ${totalFormulas}\nTOTAL CACHED VALUES: ${totalValues}\n`;

  fs.writeFileSync('barakat_source_utf8.txt', reportStr, 'utf8');
})();
