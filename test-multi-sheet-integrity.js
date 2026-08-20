const ExcelJS = require('exceljs');

async function testMultiSheetParsing() {
  console.log('🧪 Creating test Excel workbook with 5 sheets and cross-sheet formulas...');
  const wb = new ExcelJS.Workbook();

  // Sheet 1: Raw numbers
  const ws1 = wb.addWorksheet('Sheet1');
  ws1.getCell('A1').value = 100;
  ws1.getCell('A2').value = 200;
  ws1.getCell('A3').value = 300;

  // Sheet 2: Formulas referencing Sheet1, result NOT set (undefined result)
  const ws2 = wb.addWorksheet('Sheet2');
  ws2.getCell('A1').value = { formula: 'Sheet1!A1' };
  ws2.getCell('A2').value = { formula: 'Sheet1!A2' };
  ws2.getCell('A3').value = { formula: 'Sheet1!A3' };

  // Sheet 3: Cross-sheet formula adding Sheet1 and Sheet2
  const ws3 = wb.addWorksheet('Sheet3');
  ws3.getCell('A1').value = { formula: 'Sheet1!A1+Sheet2!A1' };

  // Sheet 4: Range formula
  const ws4 = wb.addWorksheet('Sheet4');
  ws4.getCell('A1').value = { formula: 'SUM(Sheet1!A1:A3)' };

  // Sheet 5: Formula chain
  const ws5 = wb.addWorksheet('Sheet5');
  ws5.getCell('A1').value = { formula: 'Sheet3!A1*2' };

  const buffer = await wb.xlsx.writeBuffer();
  console.log('📦 Workbook generated, buffer size:', buffer.byteLength, 'bytes');

  // Now simulate excel.worker.ts parsing logic on buffer
  const loadedWb = new ExcelJS.Workbook();
  await loadedWb.xlsx.load(buffer);

  const sheetsData = {};
  const sheetOrder = [];

  for (let sIdx = 0; sIdx < loadedWb.worksheets.length; sIdx++) {
    const ws = loadedWb.worksheets[sIdx];
    const sheetId = `sheet-${sIdx + 1}`;
    sheetOrder.push(sheetId);

    const cellData = {};
    let maxRow = 0;
    let maxCol = 0;

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const rowIndex = rowNumber - 1;
      if (rowIndex > maxRow) maxRow = rowIndex;

      const rowCellMap = {};
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const colIndex = colNumber - 1;
        if (colIndex > maxCol) maxCol = colIndex;

        let v = undefined;
        let f = undefined;

        let rawFormula = undefined;
        if (cell.type === ExcelJS.ValueType.Formula) {
          rawFormula = cell.formula || cell.model?.formula || (typeof cell.value === 'object' && cell.value !== null ? cell.value.formula : undefined);
        } else if (typeof cell.value === 'object' && cell.value !== null && cell.value.formula) {
          rawFormula = cell.value.formula;
        }

        if (rawFormula && typeof rawFormula === 'string' && rawFormula.trim().length > 0) {
          const trimmed = rawFormula.trim();
          f = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
        }

        if (cell.type === ExcelJS.ValueType.Formula) {
          const res = cell.result;
          if (res !== undefined && res !== null) {
            if (typeof res === 'object') {
              if (res.result !== undefined) v = res.result;
              else if (res.error !== undefined) v = res.error;
              else v = String(res);
            } else {
              v = res;
            }
          }
        } else if (cell.value !== undefined && cell.value !== null) {
          const val = cell.value;
          if (typeof val === 'object') {
            if (val instanceof Date) {
              v = val.toISOString();
            } else if (Array.isArray(val.richText)) {
              v = val.richText.map(t => t.text || '').join('');
            } else if (val.text !== undefined) {
              v = val.text;
            } else if (val.result !== undefined) {
              v = val.result;
            } else {
              v = String(val);
            }
          } else {
            v = val;
          }
        }

        const cellInfo = {};
        if (v !== undefined && v !== null && v !== '') cellInfo.v = v;
        if (f !== undefined && f !== null && f !== '') cellInfo.f = f;

        if (Object.keys(cellInfo).length > 0) {
          rowCellMap[colIndex] = cellInfo;
        }
      });

      if (Object.keys(rowCellMap).length > 0) {
        cellData[rowIndex] = rowCellMap;
      }
    });

    sheetsData[sheetId] = {
      id: sheetId,
      name: ws.name,
      cellData
    };
  }

  console.log('✅ Parsed Sheets Summary:');
  for (const sId of sheetOrder) {
    const s = sheetsData[sId];
    console.log(` - Sheet "${s.name}" (${s.id}): ${Object.keys(s.cellData).length} rows with cellData:`, JSON.stringify(s.cellData));
  }

  // Assertions
  if (Object.keys(sheetsData).length !== 5) {
    throw new Error(`FAIL: Expected 5 sheets, got ${Object.keys(sheetsData).length}`);
  }
  if (!sheetsData['sheet-2'].cellData[0][0].f || sheetsData['sheet-2'].cellData[0][0].f !== '=Sheet1!A1') {
    throw new Error(`FAIL: Sheet2 A1 formula missing or incorrect! Got: ${JSON.stringify(sheetsData['sheet-2'].cellData[0][0])}`);
  }
  if (!sheetsData['sheet-3'].cellData[0][0].f || sheetsData['sheet-3'].cellData[0][0].f !== '=Sheet1!A1+Sheet2!A1') {
    throw new Error(`FAIL: Sheet3 A1 formula missing or incorrect! Got: ${JSON.stringify(sheetsData['sheet-3'].cellData[0][0])}`);
  }
  if (!sheetsData['sheet-4'].cellData[0][0].f || sheetsData['sheet-4'].cellData[0][0].f !== '=SUM(Sheet1!A1:A3)') {
    throw new Error(`FAIL: Sheet4 A1 formula missing or incorrect! Got: ${JSON.stringify(sheetsData['sheet-4'].cellData[0][0])}`);
  }

  console.log('\n🎉 ALL MULTI-SHEET PARSING INTEGRITY CHECKS PASSED PERFECTLY!');
}

testMultiSheetParsing().catch(err => {
  console.error('❌ Test failed:', err);
  process.exit(1);
});
