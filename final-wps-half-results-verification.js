const fs = require('fs');
const ExcelJS = require('exceljs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Final Verification of WPS Half-Results Export Fix ===');

  const snapshot = {
    sheetOrder: ['sheet-1', 'sheet-2'],
    sheets: {
      'sheet-1': {
        name: 'E-Voucher & E-Wallet',
        cellData: {
          0: { 0: { v: 'Date' }, 1: { v: 'Value' } },
          1: { 0: { v: '2025-01-01' }, 1: { v: 100 } },
          2: { 0: { v: '2025-01-02' }, 1: { v: 200 } }
        }
      },
      'sheet-2': {
        name: 'Summary Sheet',
        cellData: {
          0: {
            0: { v: 'Shared Master' },
            1: { f: '=SUM(E-Voucher & E-Wallet!B2:B3)', si: '0', v: '300' }
          },
          1: {
            0: { v: 'Shared Slave' },
            1: { si: '0', v: '300' } // f is missing on slave cell!
          },
          2: {
            0: { v: 'Array Range' },
            1: { f: '=SUM(IF(E-Voucher & E-Wallet!A2:A3="2025-01-01", E-Voucher & E-Wallet!B2:B3, 0))', v: '100' }
          }
        }
      }
    }
  };

  // Run worker logic
  const exportWb = new ExcelJS.Workbook();
  exportWb.creator = 'Univer Spreadsheets App';
  exportWb.lastModifiedBy = 'Univer Spreadsheets App';
  exportWb.calcProperties.fullCalcOnLoad = true;

  const sortedSheetIds = snapshot.sheetOrder;
  const availableSheetNames = sortedSheetIds
    .map(id => snapshot.sheets[id]?.name)
    .filter(n => typeof n === 'string' && n.length > 0);

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

  for (let sIdx = 0; sIdx < sortedSheetIds.length; sIdx++) {
    const sheetId = sortedSheetIds[sIdx];
    const sheetData = snapshot.sheets[sheetId];
    if (!sheetData) continue;

    const ws = exportWb.addWorksheet(sheetData.name || sheetId);
    const cellDataRaw = sheetData.cellData || {};
    const rowKeys = Object.keys(cellDataRaw).map(Number).sort((a, b) => a - b);

    const sharedFormulaMap = new Map();
    for (let rIdx = 0; rIdx < rowKeys.length; rIdx++) {
      const rowIndex = rowKeys[rIdx];
      const cols = cellDataRaw[rowIndex];
      if (cols) {
        const colKeys = Object.keys(cols).map(Number);
        for (let cIdx = 0; cIdx < colKeys.length; cIdx++) {
          const cellPay = cols[colKeys[cIdx]];
          if (cellPay && cellPay.f && cellPay.si !== undefined && cellPay.si !== null) {
            sharedFormulaMap.set(String(cellPay.si), cellPay.f);
          }
        }
      }
    }

    for (let rIdx = 0; rIdx < rowKeys.length; rIdx++) {
      const rowIndex = rowKeys[rIdx];
      const cols = cellDataRaw[rowIndex];
      if (cols) {
        const colKeys = Object.keys(cols).map(Number).sort((a, b) => a - b);
        for (let cIdx = 0; cIdx < colKeys.length; cIdx++) {
          const colIndex = colKeys[cIdx];
          const cellPay = cols[colIndex];
          if (!cellPay) continue;

          const cell = ws.getCell(rowIndex + 1, colIndex + 1);

          let formulaText = cellPay.f;
          if (!formulaText && cellPay.si !== undefined && cellPay.si !== null) {
            formulaText = sharedFormulaMap.get(String(cellPay.si));
          }

          if (formulaText) {
            let parsedFormula = formulaText;
            if (parsedFormula.startsWith('=')) {
              parsedFormula = parsedFormula.substring(1);
            }
            parsedFormula = sanitizeSheetNames(parsedFormula);

            const isArrayFormula = cellPay.t === 1 || 
              /SUM\s*\(\s*IF\b|COUNT\s*\(\s*IF\b|AVERAGE\s*\(\s*IF\b|MODE\.MULT\b/i.test(parsedFormula) ||
              (parsedFormula.startsWith('{') && parsedFormula.endsWith('}'));

            let formulaRes = cellPay.v !== undefined && cellPay.v !== null ? cellPay.v : undefined;
            if (typeof formulaRes === 'string' && formulaRes.trim() !== '' && !isNaN(Number(formulaRes))) {
              formulaRes = Number(formulaRes);
            }

            const cellObj = {
              formula: parsedFormula,
              result: formulaRes,
            };

            if (isArrayFormula) {
              cellObj.shareType = 'array';
              cellObj.ref = cell.address;
            }

            cell.value = cellObj;
          } else if (cellPay.v !== undefined && cellPay.v !== null) {
            cell.value = cellPay.v;
          }
        }
      }
    }
  }

  const arrayBuffer = await exportWb.xlsx.writeBuffer();
  fs.writeFileSync('final-verified-half-results.xlsx', arrayBuffer);
  console.log('Saved final-verified-half-results.xlsx. Unzipping XML to verify...');

  const directory = await unzipper.Open.file('final-verified-half-results.xlsx');
  let summaryXml = '';
  for (const file of directory.files) {
    if (file.path === 'xl/worksheets/sheet2.xml') {
      summaryXml = (await file.buffer()).toString('utf8');
    }
  }

  console.log('\n--- VERIFICATION CHECKS ---');
  const masterF = summaryXml.includes('<c r="B1"><f>SUM(&apos;E-Voucher &amp; E-Wallet&apos;!B2:B3)</f>');
  const slaveF = summaryXml.includes('<c r="B2"><f>SUM(&apos;E-Voucher &amp; E-Wallet&apos;!B2:B3)</f>');
  const arrayF = summaryXml.includes('t="array" ref="B3"');

  console.log('1. Master Shared Formula export:', masterF ? 'PASS [OK]' : 'FAIL');
  console.log('2. Slave Shared Formula export (formula restored):', slaveF ? 'PASS [OK]' : 'FAIL');
  console.log('3. Special Sheet Name (& symbol) quoting:', summaryXml.includes('&apos;E-Voucher &amp; E-Wallet&apos;!') ? 'PASS [OK]' : 'FAIL');
  console.log('4. Array Formula tag present:', arrayF ? 'PASS [OK]' : 'FAIL');

  console.log('\nSummary Sheet XML Output:\n', summaryXml);
})();
