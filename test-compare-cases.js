const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

async function runComparisonTest() {
  try {
    console.log('=== STEP 1: Creating test-compare.xlsx ===');
    const wb = new ExcelJS.Workbook();
    const s1 = wb.addWorksheet('Sheet1');
    s1.getCell('A1').value = 100;

    const s2 = wb.addWorksheet('Sheet2');
    s2.getCell('A1').value = { formula: 'Sheet1!A1' };

    const xlsxPath = path.join(__dirname, 'test-compare.xlsx');
    await wb.xlsx.writeFile(xlsxPath);
    console.log('Saved test-compare.xlsx');

    console.log('\n=== STEP 2: Parsing XLSX via excel.worker logic ===');
    const fileBuffer = fs.readFileSync(xlsxPath);
    const wbRead = new ExcelJS.Workbook();
    await wbRead.xlsx.load(fileBuffer);

    const sheetsData = {};
    const sheetOrder = [];

    wbRead.worksheets.forEach((ws, sIdx) => {
      const sheetId = ws.name; // Preserving exact name as sheet ID
      sheetOrder.push(sheetId);
      const cellData = {};

      ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        const rowIndex = rowNumber - 1;
        const rowCellMap = {};
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          const colIndex = colNumber - 1;
          let f = undefined;
          let v = undefined;

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
            if (res !== undefined && res !== null) v = res;
          } else if (cell.value !== undefined && cell.value !== null) {
            v = cell.value;
          }

          const cellInfo = {};
          if (f !== undefined && f !== null && f !== '') {
            cellInfo.f = f;
          }
          if (v !== undefined && v !== null && v !== '') {
            if (typeof v === 'number') { cellInfo.v = v; cellInfo.t = 2; }
            else if (typeof v === 'boolean') { cellInfo.v = v; cellInfo.t = 3; }
            else { cellInfo.v = String(v); cellInfo.t = 1; }
          }

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
        type: 2, // SheetType.GRID
        status: sIdx === 0 ? 1 : 0,
        hidden: 0,
        rowCount: 50,
        columnCount: 20,
        zoomRatio: 1,
        scrollTop: 0,
        scrollLeft: 0,
        defaultColumnWidth: 88,
        defaultRowHeight: 24,
        showGridlines: 1,
        rightToLeft: 0,
        rowHeader: { width: 46 },
        columnHeader: { height: 20 },
        cellData,
        columnData: {},
        rowData: {},
      };
    });

    const importedWorkbookData = {
      id: `workbook-imported-${Date.now()}`,
      name: 'test-compare.xlsx',
      sheetOrder,
      appVersion: '3.0.0-alpha',
      sheets: sheetsData,
    };

    console.log('\n=== STEP 3: Launching Puppeteer to compare CASE A vs CASE B ===');
    const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
    const page = await browser.newPage();

    page.on('console', msg => {
      console.log('[BROWSER LOG]', msg.text());
    });
    page.on('pageerror', err => {
      console.log('[BROWSER PAGE ERROR]', err);
    });

    await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction('window.univerAPI !== undefined');
    await new Promise(r => setTimeout(r, 1000));

    console.log('Evaluating comparison in browser...');
    const result = await page.evaluate(async (impWbData) => {
      const api = window.univerAPI;
      const commandLog = [];

      try {
        if (api.onCommandExecuted) {
          api.onCommandExecuted((command) => {
            commandLog.push({ id: command.id, type: command.type, params: command.params });
          });
        }
      } catch (e) {
        console.log('onCommandExecuted not supported:', e.message);
      }

      function inspectWorkbookCell(wb, sheetName, r, c) {
        try {
          const sheet = wb.getSheetByName(sheetName);
          if (!sheet) return { error: `Sheet ${sheetName} not found` };
          
          const sheetId = sheet.getSheetId();
          const range = sheet.getRange(r, c, 1, 1);
          const snapshot = wb.save();
          const sheetSnap = snapshot.sheets[sheetId] || snapshot.sheets[sheetName];
          const rawCellInSnapshot = sheetSnap?.cellData?.[r]?.[c];

          return {
            unitId: wb.getId(),
            sheetId: sheetId,
            sheetName: sheet.getName(),
            sheetType: sheetSnap?.type,
            rangeValue: range.getValue(),
            rangeFormula: range.getFormula ? range.getFormula() : undefined,
            rawCellInSnapshot: rawCellInSnapshot,
            sheetConfig: {
              id: sheetSnap?.id,
              name: sheetSnap?.name,
              type: sheetSnap?.type,
              status: sheetSnap?.status
            }
          };
        } catch (e) {
          return { error: e.message, stack: e.stack };
        }
      }

      // -------------------------------------------------------------
      // CASE A: Create sheets & formula MANUALLY inside Univer
      // -------------------------------------------------------------
      console.log('[CASE A] Starting manual creation...');
      const oldWbs = api.getAllWorkbooks();
      oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

      commandLog.length = 0;
      const manualWb = api.createUniverSheet({
        id: 'wb-manual',
        name: 'Manual Workbook',
        sheetOrder: ['s1', 's2'],
        sheets: {
          's1': { id: 's1', name: 'Sheet1', type: 2, status: 1, cellData: {} },
          's2': { id: 's2', name: 'Sheet2', type: 2, status: 0, cellData: {} },
        }
      });

      const s1 = manualWb.getSheetByName('Sheet1');
      s1.getRange(0, 0, 1, 1).setValue(100);

      commandLog.length = 0;
      const s2 = manualWb.getSheetByName('Sheet2');
      s2.getRange(0, 0, 1, 1).setValue('=Sheet1!A1');

      await new Promise(r => setTimeout(r, 1500));

      const caseA_State = inspectWorkbookCell(manualWb, 'Sheet2', 0, 0);
      const caseA_Commands = [...commandLog];

      // -------------------------------------------------------------
      // CASE B: Import external XLSX via createUniverSheet
      // -------------------------------------------------------------
      console.log('[CASE B] Starting external XLSX import...');
      api.disposeUnit(manualWb.getId());

      commandLog.length = 0;
      const importedWb = api.createUniverSheet(impWbData);

      await new Promise(r => setTimeout(r, 1500));

      const caseB_ImmediateState = inspectWorkbookCell(importedWb, 'Sheet2', 0, 0);
      const caseB_ImportCommands = [...commandLog];

      // -------------------------------------------------------------
      // CASE B + MANUAL EDIT: Manually edit Sheet2!A1 on imported WB
      // -------------------------------------------------------------
      console.log('[CASE B Edit] Manually editing imported Sheet2!A1...');
      commandLog.length = 0;
      const importedS2 = importedWb.getSheetByName('Sheet2');
      importedS2.getRange(0, 0, 1, 1).setValue('=Sheet1!A1');

      await new Promise(r => setTimeout(r, 1500));

      const caseB_AfterEditState = inspectWorkbookCell(importedWb, 'Sheet2', 0, 0);
      const caseB_ManualEditCommands = [...commandLog];

      return {
        caseA: {
          state: caseA_State,
          commands: caseA_Commands
        },
        caseB_Immediate: {
          state: caseB_ImmediateState,
          importCommands: caseB_ImportCommands
        },
        caseB_AfterEdit: {
          state: caseB_AfterEditState,
          editCommands: caseB_ManualEditCommands
        }
      };
    }, importedWorkbookData);

    fs.writeFileSync('comparison-results.json', JSON.stringify(result, null, 2));
    console.log('\n=== RESULTS WRITTEN TO comparison-results.json ===');
    console.log(JSON.stringify(result, null, 2));

    await browser.close();
  } catch (err) {
    console.error('SCRIPT ERROR:', err);
  }
}

runComparisonTest();
