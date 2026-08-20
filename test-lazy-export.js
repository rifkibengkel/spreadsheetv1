const fs = require('fs');
const ExcelJS = require('exceljs');

// Test lazy chunking logic standalone
async function testLazyExportLogic() {
  console.log('Testing lazy chunking logic...');

  // Create mock snapshot with 50,000 cells
  const mockCellData = {};
  for (let r = 0; r < 5000; r++) {
    mockCellData[r] = {
      0: { v: `Name_${r}`, t: 1 },
      1: { v: r * 10, t: 2 },
      2: { f: `=B${r + 1}*2`, v: r * 20 }
    };
  }

  const snapshot = {
    sheetOrder: ['Sheet1'],
    sheets: {
      Sheet1: {
        name: 'Sheet1',
        cellData: mockCellData
      }
    }
  };

  const workbook = new ExcelJS.Workbook();
  const sortedSheetIds = snapshot.sheetOrder;

  let totalRowsOverall = 0;
  sortedSheetIds.forEach((sheetId) => {
    const sData = snapshot.sheets[sheetId];
    if (sData && sData.cellData) {
      totalRowsOverall += Object.keys(sData.cellData).length;
    }
  });

  let processedRowsOverall = 0;
  const yieldToMain = () => new Promise((resolve) => setTimeout(resolve, 0));
  const CHUNK_SIZE = 500;
  const progressLogs = [];

  for (let sIdx = 0; sIdx < sortedSheetIds.length; sIdx++) {
    const sheetId = sortedSheetIds[sIdx];
    const sheetData = snapshot.sheets[sheetId];
    const ws = workbook.addWorksheet(sheetData.name || sheetId);
    const cellDataRaw = sheetData.cellData || {};
    const rowKeys = Object.keys(cellDataRaw);

    for (let rIdx = 0; rIdx < rowKeys.length; rIdx++) {
      const rowKey = rowKeys[rIdx];
      const rowIndex = parseInt(rowKey, 10);
      const cols = cellDataRaw[rowIndex];

      if (cols) {
        Object.keys(cols).forEach((colKey) => {
          const colIndex = parseInt(colKey, 10);
          const cellPay = cols[colIndex];
          const cell = ws.getCell(rowIndex + 1, colIndex + 1);
          if (cellPay.f) {
            let parsedFormula = cellPay.f.startsWith('=') ? cellPay.f.substring(1) : cellPay.f;
            cell.value = { formula: parsedFormula, result: cellPay.v };
          } else if (cellPay.v !== undefined && cellPay.v !== null) {
            cell.value = cellPay.v;
          }
        });
      }

      processedRowsOverall++;

      if (rIdx % CHUNK_SIZE === 0 || rIdx === rowKeys.length - 1) {
        const percent = Math.min(85, Math.round(5 + (processedRowsOverall / totalRowsOverall) * 80));
        progressLogs.push(`Progress: ${percent}% - Row ${rIdx + 1}/${rowKeys.length}`);
        await yieldToMain();
      }
    }
  }

  const buf = await workbook.xlsx.writeBuffer();
  console.log(`Export completed! Buffer size: ${buf.byteLength} bytes.`);
  console.log('Sample progress logs:');
  console.log(progressLogs.slice(0, 5));
  console.log(progressLogs.slice(-3));
}

testLazyExportLogic().catch(console.error);
