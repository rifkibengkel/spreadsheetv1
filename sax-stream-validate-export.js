const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== SAX Stream Forensic Validation of Exported XLSX ===');

  const fileList = ['wps-export-fixed-test.xlsx', 'final-verified-half-results.xlsx', 'final-verified-wps-export.xlsx'];

  for (const filePath of fileList) {
    if (!fs.existsSync(filePath)) continue;

    console.log(`\n========================================`);
    console.log(`VALIDATING FILE: ${filePath}`);
    console.log(`========================================`);

    const directory = await unzipper.Open.file(filePath);

    let totalFormulas = 0;
    let totalArrayFormulas = 0;
    let totalCachedValues = 0;

    for (const file of directory.files) {
      if (file.path.startsWith('xl/worksheets/sheet') && file.path.endsWith('.xml')) {
        const content = await file.buffer();
        const xmlStr = content.toString('utf8');

        const formulas = (xmlStr.match(/<f[^>]*>/g) || []).length;
        const arrays = (xmlStr.match(/<f[^>]*t="array"[^>]*>/g) || []).length;
        const cached = (xmlStr.match(/<v>[^<]*<\/v>/g) || []).length;

        totalFormulas += formulas;
        totalArrayFormulas += arrays;
        totalCachedValues += cached;

        console.log(`  - ${file.path}: ${formulas} formulas (${arrays} array), ${cached} cached values`);
      }
    }

    let fullCalcOnLoad = false;
    const wbFile = directory.files.find(f => f.path === 'xl/workbook.xml');
    if (wbFile) {
      const wbXml = (await wbFile.buffer()).toString('utf8');
      fullCalcOnLoad = wbXml.includes('fullCalcOnLoad="1"');
    }

    console.log(`Summary Checklist for ${filePath}:`);
    console.log('  1. calcPr fullCalcOnLoad="1":', fullCalcOnLoad ? 'PASS [OK]' : 'FAIL');
    console.log('  2. Total Formulas:', totalFormulas);
    console.log('  3. Array Formulas:', totalArrayFormulas);
    console.log('  4. Cached Values:', totalCachedValues);
  }
})();
