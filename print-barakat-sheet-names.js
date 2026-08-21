const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat_07-03-07.xlsx';
  const directory = await unzipper.Open.file(filePath);

  const workbookFile = directory.files.find(f => f.path === 'xl/workbook.xml');
  if (workbookFile) {
    const wbXml = (await workbookFile.buffer()).toString('utf8');
    const sheetMatches = wbXml.match(/<sheet[^>]+>/g) || [];
    console.log('Sheet tags in workbook.xml:');
    sheetMatches.forEach((s, i) => console.log(`Sheet ${i+1}: ${s}`));
  }
})();
