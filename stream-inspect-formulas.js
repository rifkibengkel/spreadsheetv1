const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Fast SAX XML Formula Inspection of updatehexos.xlsx ===');

  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx';
  const directory = await unzipper.Open.file(filePath);

  const sheetFiles = directory.files.filter(f => f.path.startsWith('xl/worksheets/sheet') && f.path.endsWith('.xml'));

  console.log(`Found ${sheetFiles.length} worksheet XML files.`);

  let totalFormulasAllSheets = 0;
  const sheetCounts = [];

  for (const file of sheetFiles) {
    const content = await file.buffer();
    const xmlStr = content.toString('utf8');

    // Match <f> tags or f= attributes
    const fMatches = xmlStr.match(/<f[^>]*>([^<]*)<\/f>/g) || [];
    totalFormulasAllSheets += fMatches.length;

    // Check for full column references like A:A, Summary!A:Z, etc.
    let fullColRefs = 0;
    fMatches.forEach(fTag => {
      if (/[A-Z]+:[A-Z]+/i.test(fTag)) fullColRefs++;
    });

    sheetCounts.push({
      path: file.path,
      formulaCount: fMatches.length,
      fullColRefs,
      sampleFormula: fMatches[0] ? fMatches[0].slice(0, 80) : 'none'
    });
  }

  console.log(`TOTAL FORMULAS ACROSS ALL SHEETS: ${totalFormulasAllSheets}`);
  console.log('Sheet Breakdown:', JSON.stringify(sheetCounts, null, 2));

  fs.writeFileSync('sax-formula-inspection.json', JSON.stringify({
    totalFormulasAllSheets,
    sheetCounts
  }, null, 2));
})();
