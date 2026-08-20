const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');

async function runForensicAudit() {
  console.log('=== STEP 1: PARSING TEST-COMPARE.XLSX WITH CUSTOM PARSER LOGIC ===');
  const filePath = path.join(__dirname, 'test-compare.xlsx');
  const fileBuffer = fs.readFileSync(filePath);
  
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(fileBuffer);

  const rawExcelJSCells = {};
  const customParsedSheetsData = {};
  const sheetOrder = [];

  const totalSheets = workbook.worksheets.length;
  workbook.worksheets.forEach((ws, sIdx) => {
    const sheetId = ws.name.toLowerCase().replace(/[^a-z0-9]/g, '-') || `sheet-${sIdx + 1}`;
    sheetOrder.push(sheetId);
    
    rawExcelJSCells[ws.name] = {};
    const cellData = {};

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const rowIndex = rowNumber - 1;
      const rowCellMap = {};
      
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const colIndex = colNumber - 1;

        // Raw ExcelJS inspection
        const cellAddress = cell.address;
        rawExcelJSCells[ws.name][cellAddress] = {
          address: cell.address,
          type: cell.type,
          value: cell.value,
          result: cell.result,
          formula: cell.formula,
          model: cell.model,
          style: cell.style,
          numFmt: cell.numFmt
        };

        // Custom Parser extraction logic (from excel.worker.ts)
        let v = undefined;
        let f = undefined;

        let rawFormula = undefined;
        if (cell.type === ExcelJS.ValueType.Formula) {
          rawFormula = cell.formula || (cell.model && cell.model.formula) || (typeof cell.value === 'object' && cell.value !== null ? cell.value.formula : undefined);
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
            if (val instanceof Date) v = val.toISOString();
            else if (Array.isArray(val.richText)) v = val.richText.map(t => t.text || '').join('');
            else if (val.text !== undefined) v = val.text;
            else if (val.result !== undefined) v = val.result;
            else v = String(val);
          } else {
            v = val;
          }
        }

        const cellInfo = {};
        if (f !== undefined && f !== null && f !== '') {
          const trimmed = f.trim();
          cellInfo.f = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
        }
        if (v !== undefined && v !== null && v !== '') {
          if (typeof v === 'number') {
            cellInfo.v = v;
            cellInfo.t = 2; // CellValueType.NUMBER
          } else if (typeof v === 'boolean') {
            cellInfo.v = v;
            cellInfo.t = 3; // CellValueType.BOOLEAN
          } else if (typeof v === 'string') {
            const trimmedV = v.trim();
            const numV = Number(trimmedV);
            if (!isNaN(numV) && trimmedV !== '' && !cellInfo.f) {
              cellInfo.v = numV;
              cellInfo.t = 2;
            } else {
              cellInfo.v = v;
              cellInfo.t = 1;
            }
          } else {
            cellInfo.v = String(v);
            cellInfo.t = 1;
          }
        }

        if (Object.keys(cellInfo).length > 0) {
          rowCellMap[colIndex] = cellInfo;
        }
      });

      if (Object.keys(rowCellMap).length > 0) {
        cellData[rowIndex] = rowCellMap;
      }
    });

    customParsedSheetsData[sheetId] = {
      id: sheetId,
      name: ws.name,
      type: 2,
      status: sIdx === 0 ? 1 : 0,
      rowCount: 100,
      columnCount: 30,
      cellData
    };
  });

  const customWorkbookData = {
    id: `workbook-imported-${Date.now()}`,
    name: 'test-compare.xlsx',
    sheetOrder,
    appVersion: '3.0.0-alpha',
    sheets: customParsedSheetsData
  };

  console.log('\n--- 1A. RAW EXCELJS PARSE RESULT ---');
  console.log(JSON.stringify(rawExcelJSCells, null, 2));

  console.log('\n--- 1B. OUR CUSTOM PARSED WORKBOOK DATA ---');
  console.log(JSON.stringify(customWorkbookData, null, 2));

  console.log('\n=== STEP 2: PUPPETEER BROWSER FORENSIC AUDIT (MANUAL VS IMPORTED IN UNIVER) ===');
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  await new Promise(r => setTimeout(r, 2000));

  const forensicResult = await page.evaluate(async (customWbData) => {
    const api = window.univerAPI;
    const auditLog = {};

    // --- MANUAL TEST IN UNIVER ---
    // Clear and create a clean manual workbook
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
    oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });

    const manualInitData = {
      id: 'manual-wb',
      name: 'ManualWorkbook',
      sheetOrder: ['sheet1', 'sheet2'],
      appVersion: '3.0.0-alpha',
      sheets: {
        'sheet1': { id: 'sheet1', name: 'Sheet1', type: 2, status: 1, rowCount: 50, columnCount: 20, cellData: {} },
        'sheet2': { id: 'sheet2', name: 'Sheet2', type: 2, status: 0, rowCount: 50, columnCount: 20, cellData: {} }
      }
    };

    api.createUniverSheet(manualInitData);
    await new Promise(r => setTimeout(r, 500));

    const manualWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    const unitId = manualWb.getId();

    // Execute commands to simulate user typing manual formulas
    // Sheet1: A1=100, A2=200, A3==A1+A2
    await api.executeCommand('sheet.command.set-range-values', {
      unitId, subUnitId: 'sheet1', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, value: { v: 100, t: 2 }
    });
    await api.executeCommand('sheet.command.set-range-values', {
      unitId, subUnitId: 'sheet1', range: { startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 }, value: { v: 200, t: 2 }
    });
    await api.executeCommand('sheet.command.set-range-values', {
      unitId, subUnitId: 'sheet1', range: { startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 }, value: { f: '=A1+A2' }
    });

    // Sheet2: A1==Sheet1!A1, A2==Sheet1!A2, A3==Sheet1!A1+Sheet1!A2
    await api.executeCommand('sheet.command.set-range-values', {
      unitId, subUnitId: 'sheet2', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, value: { f: '=Sheet1!A1' }
    });
    await api.executeCommand('sheet.command.set-range-values', {
      unitId, subUnitId: 'sheet2', range: { startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 }, value: { f: '=Sheet1!A2' }
    });
    await api.executeCommand('sheet.command.set-range-values', {
      unitId, subUnitId: 'sheet2', range: { startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 }, value: { f: '=Sheet1!A1+Sheet1!A2' }
    });

    // Wait for calculation engine to calculate
    await new Promise(r => setTimeout(r, 2000));

    // Read manual cell models
    const manualSnapshot = manualWb.save();
    auditLog.manual_workbook_snapshot_cellData = {
      sheet1: {
        A1: manualSnapshot.sheets['sheet1']?.cellData?.[0]?.[0],
        A2: manualSnapshot.sheets['sheet1']?.cellData?.[1]?.[0],
        A3: manualSnapshot.sheets['sheet1']?.cellData?.[2]?.[0],
      },
      sheet2: {
        A1: manualSnapshot.sheets['sheet2']?.cellData?.[0]?.[0],
        A2: manualSnapshot.sheets['sheet2']?.cellData?.[1]?.[0],
        A3: manualSnapshot.sheets['sheet2']?.cellData?.[2]?.[0],
      }
    };

    auditLog.manual_calculated_values = {
      Sheet1_A3_value: manualWb.getSheetByName('Sheet1')?.getRange(2, 0, 1, 1).getValue(),
      Sheet2_A1_value: manualWb.getSheetByName('Sheet2')?.getRange(0, 0, 1, 1).getValue(),
      Sheet2_A2_value: manualWb.getSheetByName('Sheet2')?.getRange(1, 0, 1, 1).getValue(),
      Sheet2_A3_value: manualWb.getSheetByName('Sheet2')?.getRange(2, 0, 1, 1).getValue(),
    };

    // --- IMPORTED TEST IN UNIVER ---
    const allWbs2 = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
    allWbs2.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });

    // Load custom imported workbook
    api.createUniverSheet(customWbData);

    const importedWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];

    // Immediately inspect cell data after creation
    const importedImmediateSnapshot = importedWb.save();
    auditLog.imported_immediate_snapshot_cellData = {
      sheet1: {
        A1: importedImmediateSnapshot.sheets[customWbData.sheetOrder[0]]?.cellData?.[0]?.[0],
        A2: importedImmediateSnapshot.sheets[customWbData.sheetOrder[0]]?.cellData?.[1]?.[0],
        A3: importedImmediateSnapshot.sheets[customWbData.sheetOrder[0]]?.cellData?.[2]?.[0],
      },
      sheet2: {
        A1: importedImmediateSnapshot.sheets[customWbData.sheetOrder[1]]?.cellData?.[0]?.[0],
        A2: importedImmediateSnapshot.sheets[customWbData.sheetOrder[1]]?.cellData?.[1]?.[0],
        A3: importedImmediateSnapshot.sheets[customWbData.sheetOrder[1]]?.cellData?.[2]?.[0],
      }
    };

    auditLog.imported_immediate_values = {
      Sheet1_A3_value: importedWb.getSheetByName('Sheet1')?.getRange(2, 0, 1, 1).getValue(),
      Sheet2_A1_value: importedWb.getSheetByName('Sheet2')?.getRange(0, 0, 1, 1).getValue(),
      Sheet2_A2_value: importedWb.getSheetByName('Sheet2')?.getRange(1, 0, 1, 1).getValue(),
      Sheet2_A3_value: importedWb.getSheetByName('Sheet2')?.getRange(2, 0, 1, 1).getValue(),
    };

    // Wait 2 seconds for any reactive formula calculation
    await new Promise(r => setTimeout(r, 2000));

    const importedSettledSnapshot = importedWb.save();
    auditLog.imported_settled_snapshot_cellData = {
      sheet1: {
        A1: importedSettledSnapshot.sheets[customWbData.sheetOrder[0]]?.cellData?.[0]?.[0],
        A2: importedSettledSnapshot.sheets[customWbData.sheetOrder[0]]?.cellData?.[1]?.[0],
        A3: importedSettledSnapshot.sheets[customWbData.sheetOrder[0]]?.cellData?.[2]?.[0],
      },
      sheet2: {
        A1: importedSettledSnapshot.sheets[customWbData.sheetOrder[1]]?.cellData?.[0]?.[0],
        A2: importedSettledSnapshot.sheets[customWbData.sheetOrder[1]]?.cellData?.[1]?.[0],
        A3: importedSettledSnapshot.sheets[customWbData.sheetOrder[1]]?.cellData?.[2]?.[0],
      }
    };

    auditLog.imported_settled_values = {
      Sheet1_A3_value: importedWb.getSheetByName('Sheet1')?.getRange(2, 0, 1, 1).getValue(),
      Sheet2_A1_value: importedWb.getSheetByName('Sheet2')?.getRange(0, 0, 1, 1).getValue(),
      Sheet2_A2_value: importedWb.getSheetByName('Sheet2')?.getRange(1, 0, 1, 1).getValue(),
      Sheet2_A3_value: importedWb.getSheetByName('Sheet2')?.getRange(2, 0, 1, 1).getValue(),
    };

    return auditLog;
  }, customWorkbookData);

  const finalReport = {
    rawExcelJSCells,
    customWorkbookData,
    forensicResult
  };

  fs.writeFileSync(path.join(__dirname, 'forensic-output.json'), JSON.stringify(finalReport, null, 2));
  console.log('✅ Forensic audit complete! Output saved to forensic-output.json');

  await browser.close();
}

runForensicAudit().catch(console.error);
