const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');

async function runInvestigate() {
  console.log('--- STARTING BUG INVESTIGATION ---');

  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();

  page.on('console', msg => console.log('[BROWSER LOG]', msg.text()));
  page.on('pageerror', err => console.log('[BROWSER PAGE ERROR]', err.message));

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Page loaded and window.univerAPI ready.');

  const results = await page.evaluate(async () => {
    const api = window.univerAPI;
    const commandLog = [];

    // Helper to log executed commands
    let listenerRef = null;
    try {
      if (api.onCommandExecuted) {
        listenerRef = api.onCommandExecuted((cmd) => {
          commandLog.push({ id: cmd.id, type: cmd.type, params: cmd.params });
        });
      }
    } catch (e) {}

    function getWorkbookSnapshot(wb) {
      if (!wb) return null;
      return wb.save ? wb.save() : null;
    }

    // -------------------------------------------------------------
    // CASE A: MANUAL FORMULA INSERTION INSIDE UNIVER
    // -------------------------------------------------------------
    console.log('[CASE A] Creating Workbook Manually with Sheet1 and Sheet2...');
    
    // Clear active workbooks
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
    oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

    const manualWbData = {
      id: `wb-manual-${Date.now()}`,
      name: 'Manual Workbook',
      sheetOrder: ['s1', 's2'],
      appVersion: '3.0.0-alpha',
      sheets: {
        's1': { id: 's1', name: 'Sheet1', type: 2, status: 1, cellData: {} },
        's2': { id: 's2', name: 'Sheet2', type: 2, status: 0, cellData: {} },
      }
    };

    const caseA_wb = api.createUniverSheet(manualWbData);
    
    // 1. Set Sheet1!A1 = 100
    const caseA_s1 = caseA_wb.getSheetByName('Sheet1');
    caseA_s1.getRange(0, 0, 1, 1).setValue(100);

    // 2. Set Sheet2!A1 = =Sheet1!A1 manually (capturing command)
    commandLog.length = 0;
    const caseA_s2 = caseA_wb.getSheetByName('Sheet2');
    caseA_s2.getRange(0, 0, 1, 1).setValue('=Sheet1!A1');

    await new Promise(r => setTimeout(r, 1500)); // wait for formula engine

    const caseA_cellValue = caseA_s2.getRange(0, 0, 1, 1).getValue();
    const caseA_snapshot = getWorkbookSnapshot(caseA_wb);
    const caseA_commandsFired = [...commandLog];

    // -------------------------------------------------------------
    // CASE B: IMPORTED FORMULA (No cached 'v', exact output of excel.worker)
    // -------------------------------------------------------------
    console.log('[CASE B] Importing Workbook Data (No cached formula result)...');

    // Clear previous unit
    try { api.disposeUnit(caseA_wb.getId()); } catch(e){}

    const importedWbData = {
      id: `wb-imported-${Date.now()}`,
      name: 'Imported Workbook',
      sheetOrder: ['Sheet1', 'Sheet2'],
      appVersion: '3.0.0-alpha',
      sheets: {
        'Sheet1': {
          id: 'Sheet1',
          name: 'Sheet1',
          type: 2,
          status: 1,
          cellData: {
            0: { 0: { v: 100, t: 2 } } // Sheet1!A1 = 100
          }
        },
        'Sheet2': {
          id: 'Sheet2',
          name: 'Sheet2',
          type: 2,
          status: 0,
          cellData: {
            0: { 0: { f: '=Sheet1!A1' } } // Sheet2!A1 = =Sheet1!A1 (NO v, NO t)
          }
        }
      }
    };

    commandLog.length = 0;
    const caseB_wb = api.createUniverSheet(importedWbData);

    await new Promise(r => setTimeout(r, 1500)); // wait for formula engine

    const caseB_s2 = caseB_wb.getSheetByName('Sheet2');
    const caseB_immediateValue = caseB_s2.getRange(0, 0, 1, 1).getValue();
    const caseB_immediateSnapshot = getWorkbookSnapshot(caseB_wb);
    const caseB_importCommands = [...commandLog];

    // -------------------------------------------------------------
    // CASE B + MANUAL EDIT: Manually Edit Sheet2!A1 on imported workbook
    // -------------------------------------------------------------
    console.log('[CASE B + MANUAL EDIT] Manually editing Sheet2!A1 on imported workbook...');

    commandLog.length = 0;
    caseB_s2.getRange(0, 0, 1, 1).setValue('=Sheet1!A1');

    await new Promise(r => setTimeout(r, 1500)); // wait for formula engine

    const caseB_afterEditValue = caseB_s2.getRange(0, 0, 1, 1).getValue();
    const caseB_afterEditSnapshot = getWorkbookSnapshot(caseB_wb);
    const caseB_manualEditCommands = [...commandLog];

    return {
      caseA: {
        cellValue: caseA_cellValue,
        cellDataInSnapshot: caseA_snapshot?.sheets['s2']?.cellData?.[0]?.[0],
        sheetConfig: {
          id: caseA_snapshot?.sheets['s2']?.id,
          name: caseA_snapshot?.sheets['s2']?.name,
          type: caseA_snapshot?.sheets['s2']?.type,
        },
        commands: caseA_commandsFired
      },
      caseB_Immediate: {
        cellValue: caseB_immediateValue,
        cellDataInSnapshot: caseB_immediateSnapshot?.sheets['Sheet2']?.cellData?.[0]?.[0],
        sheetConfig: {
          id: caseB_immediateSnapshot?.sheets['Sheet2']?.id,
          name: caseB_immediateSnapshot?.sheets['Sheet2']?.name,
          type: caseB_immediateSnapshot?.sheets['Sheet2']?.type,
        },
        commands: caseB_importCommands
      },
      caseB_AfterManualEdit: {
        cellValue: caseB_afterEditValue,
        cellDataInSnapshot: caseB_afterEditSnapshot?.sheets['Sheet2']?.cellData?.[0]?.[0],
        sheetConfig: {
          id: caseB_afterEditSnapshot?.sheets['Sheet2']?.id,
          name: caseB_afterEditSnapshot?.sheets['Sheet2']?.name,
          type: caseB_afterEditSnapshot?.sheets['Sheet2']?.type,
        },
        commands: caseB_manualEditCommands
      }
    };
  });

  fs.writeFileSync('investigation-report.json', JSON.stringify(results, null, 2));
  console.log('--- INVESTIGATION REPORT SAVED ---');
  console.log(JSON.stringify(results, null, 2));

  await browser.close();
}

runInvestigate().catch(console.error);
