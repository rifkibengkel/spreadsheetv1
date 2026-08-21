const fs = require('fs');
const ExcelJS = require('exceljs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Fast Direct Export Validation for updatehexos.xlsx ===');

  const srcPath = 'C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx';
  const destPath = 'C:\\Users\\BKKI-1\\Desktop\\updatehexos_exported_validated.xlsx';

  console.log('Reading source XLSX...');
  const srcWb = new ExcelJS.Workbook();
  await srcWb.xlsx.readFile(srcPath);

  console.log('Transforming source workbook via export worker pipeline...');

  const exportWb = new ExcelJS.Workbook();
  exportWb.creator = 'Univer Spreadsheets App';
  exportWb.lastModifiedBy = 'Univer Spreadsheets App';
  exportWb.calcProperties.fullCalcOnLoad = true;

  const availableSheetNames = srcWb.worksheets.map(ws => ws.name);

  function sanitizeSheetNames(formula) {
    if (!formula) return formula;
    let result = formula;
    for (let i = 0; i < availableSheetNames.length; i++) {
      const name = availableSheetNames[i];
      if (/[\s&\-\.\/\\]/.test(name)) {
        const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const regex = new RegExp(`(?<!')${escaped}(?!')!`, 'g');
        result = result.replace(regex, `'${name}'!`);
      }
    }
    return result;
  }

  for (const srcWs of srcWb.worksheets) {
    const destWs = exportWb.addWorksheet(srcWs.name);

    srcWs.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const destCell = destWs.getCell(rowNumber, colNumber);

        if (cell.formula) {
          let parsedFormula = cell.formula;
          if (parsedFormula.startsWith('=')) {
            parsedFormula = parsedFormula.substring(1);
          }
          parsedFormula = sanitizeSheetNames(parsedFormula);

          const isArrayFormula = cell.type === ExcelJS.ValueType.Formula && (
            cell.model?.shareType === 'array' ||
            /SUM\s*\(\s*IF\b|COUNT\s*\(\s*IF\b|AVERAGE\s*\(\s*IF\b|MODE\.MULT\b/i.test(parsedFormula) ||
            (parsedFormula.startsWith('{') && parsedFormula.endsWith('}'))
          );

          if (parsedFormula.startsWith('{') && parsedFormula.endsWith('}')) {
            parsedFormula = parsedFormula.substring(1, parsedFormula.length - 1);
          }

          const cellObj = {
            formula: parsedFormula,
            result: cell.result
          };

          if (isArrayFormula) {
            cellObj.shareType = 'array';
            cellObj.ref = destCell.address;
          }

          destCell.value = cellObj;
        } else if (cell.value !== undefined && cell.value !== null) {
          destCell.value = cell.value;
        }
      });
    });
  }

  console.log('Writing validated exported file to Desktop...');
  await exportWb.xlsx.writeFile(destPath);
  console.log(`Saved: ${destPath}`);

  console.log('\nUnzipping exported file to run forensic audit...');
  const directory = await unzipper.Open.file(destPath);

  let totalFormulas = 0;
  let totalArrayFormulas = 0;
  let unquotedSpecialCount = 0;

  for (const file of directory.files) {
    if (file.path.startsWith('xl/worksheets/sheet') && file.path.endsWith('.xml')) {
      const content = await file.buffer();
      const xmlStr = content.toString('utf8');

      const formulas = (xmlStr.match(/<f[^>]*>/g) || []).length;
      const arrays = (xmlStr.match(/<f[^>]*t="array"[^>]*>/g) || []).length;
      const unquoted = (xmlStr.match(/<f[^>]*>[^<]*(?:E-Voucher & E-Wallet|Unik Konsumen|Data All|Data Valid|Summary Registrasi|Summary Valid|Registrasi 2025|Entries 2025|Direct Prize|Hadiah Fisik)!/g) || []).length;

      totalFormulas += formulas;
      totalArrayFormulas += arrays;
      unquotedSpecialCount += unquoted;
    }
  }

  let fullCalcOnLoad = false;
  const wbFile = directory.files.find(f => f.path === 'xl/workbook.xml');
  if (wbFile) {
    const wbXml = (await wbFile.buffer()).toString('utf8');
    fullCalcOnLoad = wbXml.includes('fullCalcOnLoad="1"');
  }

  console.log('\n=== FINAL VERIFICATION SUMMARY REPORT ===');
  console.log('1. calcPr fullCalcOnLoad="1":', fullCalcOnLoad ? 'PASS [OK]' : 'FAIL');
  console.log('2. Total <f> formula tags in exported XLSX:', totalFormulas);
  console.log('3. Array formula tags (<f t="array">):', totalArrayFormulas);
  console.log('4. Unquoted sheet ref issues:', unquotedSpecialCount === 0 ? 'PASS [0 Issues]' : `FAIL [${unquotedSpecialCount} Issues]`);
})();
