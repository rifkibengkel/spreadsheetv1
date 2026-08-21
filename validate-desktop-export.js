const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== FORENSIC VALIDATION OF EXPORTED WORKBOOK XML ===');

  const exportedPath = 'C:\\Users\\BKKI-1\\Desktop\\updatehexos_exported_validated.xlsx';
  if (!fs.existsSync(exportedPath)) {
    console.log(`File not found: ${exportedPath}`);
    return;
  }

  console.log(`Analyzing exported file: ${exportedPath}`);

  const directory = await unzipper.Open.file(exportedPath);

  let totalFormulas = 0;
  let totalArrayFormulas = 0;
  let totalCachedValues = 0;
  let unquotedSpecialSheetRefs = 0;

  const sheetSummaries = [];

  for (const file of directory.files) {
    if (file.path.startsWith('xl/worksheets/sheet') && file.path.endsWith('.xml')) {
      const content = await file.buffer();
      const xmlStr = content.toString('utf8');

      const formulaCount = (xmlStr.match(/<f[^>]*>/g) || []).length;
      const arrayCount = (xmlStr.match(/<f[^>]*t="array"[^>]*>/g) || []).length;
      const cachedCount = (xmlStr.match(/<v>[^<]*<\/v>/g) || []).length;

      // Check for unquoted sheet names with special characters
      const unquotedSpecial = (xmlStr.match(/<f[^>]*>[^<]*(?:E-Voucher & E-Wallet|Unik Konsumen|Data All|Data Valid|Summary Registrasi|Summary Valid|Registrasi 2025|Entries 2025|Direct Prize|Hadiah Fisik)!/g) || []).length;

      unquotedSpecialSheetRefs += unquotedSpecial;
      totalFormulas += formulaCount;
      totalArrayFormulas += arrayCount;
      totalCachedValues += cachedCount;

      sheetSummaries.push({
        sheet: file.path,
        formulas: formulaCount,
        arrayFormulas: arrayCount,
        cachedValues: cachedCount,
        unquotedSpecial
      });
    }
  }

  let calcPrFullCalcOnLoad = false;
  const workbookFile = directory.files.find(f => f.path === 'xl/workbook.xml');
  if (workbookFile) {
    const wbXml = (await workbookFile.buffer()).toString('utf8');
    calcPrFullCalcOnLoad = wbXml.includes('fullCalcOnLoad="1"');
  }

  console.log('\n--- VERIFICATION REPORT ---');
  console.log('1. Workbook calcPr fullCalcOnLoad="1":', calcPrFullCalcOnLoad ? 'PASS [OK]' : 'FAIL');
  console.log('2. Total <f> formula tags in exported XLSX:', totalFormulas);
  console.log('3. Array formula tags (<f t="array">):', totalArrayFormulas);
  console.log('4. Unquoted sheet ref issues (spaces/&/-):', unquotedSpecialSheetRefs === 0 ? 'PASS [0 Issues]' : `FAIL [${unquotedSpecialSheetRefs} Issues]`);

  console.log('\nSheet-by-Sheet Summary:');
  console.log(JSON.stringify(sheetSummaries, null, 2));
})();
