const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');

async function testFormulaNoCacheComparison() {
  console.log('--- Testing Formula Without Cached Value ---');

  const impWbDataWithoutV = {
    id: `workbook-imported-${Date.now()}`,
    name: 'test-no-cache-import.xlsx',
    sheetOrder: ['Sheet1', 'Sheet2'],
    appVersion: '3.0.0-alpha',
    sheets: {
      'Sheet1': {
        id: 'Sheet1',
        name: 'Sheet1',
        type: 2,
        status: 1,
        rowCount: 50,
        columnCount: 20,
        cellData: {
          0: { 0: { v: 100, t: 2 } }, // Sheet1!A1 = 100
          1: { 0: { v: 200, t: 2 } }  // Sheet1!A2 = 200
        }
      },
      'Sheet2': {
        id: 'Sheet2',
        name: 'Sheet2',
        type: 2,
        status: 0,
        rowCount: 50,
        columnCount: 20,
        cellData: {
          0: { 0: { f: '=Sheet1!A1' } },         // Sheet2!A1 = =Sheet1!A1
          1: { 0: { f: '=Sheet1!A2' } },         // Sheet2!A2 = =Sheet1!A2
          2: { 0: { f: '=A1+A2' } }              // Sheet2!A3 = =A1+A2
        }
      }
    }
  };

  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  page.on('console', msg => console.log('[BROWSER LOG]', msg.text()));
  page.on('pageerror', err => console.log('[BROWSER PAGE ERROR]', err.message));

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Evaluating in browser...');
  const res = await page.evaluate(async (data) => {
    const api = window.univerAPI;

    try {
      const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
      oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

      console.log('Step 1: Calling createUniverSheet...');
      const impWb = api.createUniverSheet(data);

      await new Promise(r => setTimeout(r, 2500));

      const is2 = impWb.getSheetByName('Sheet2');
      const valA1_imported = is2 ? is2.getRange(0, 0, 1, 1).getValue() : 'NO_SHEET2';
      const valA2_imported = is2 ? is2.getRange(1, 0, 1, 1).getValue() : 'NO_SHEET2';
      const valA3_imported = is2 ? is2.getRange(2, 0, 1, 1).getValue() : 'NO_SHEET2';

      const cellA1_imported_snapshot = impWb.save().sheets['Sheet2']?.cellData?.[0]?.[0];
      const cellA3_imported_snapshot = impWb.save().sheets['Sheet2']?.cellData?.[2]?.[0];

      console.log('Step 2: Manually editing Sheet2!A1...');
      const editCommands = [];
      if (api.onCommandExecuted) {
        api.onCommandExecuted((cmd) => {
          editCommands.push({ id: cmd.id, type: cmd.type, params: cmd.params });
        });
      }

      if (is2) {
        is2.getRange(0, 0, 1, 1).setValue('=Sheet1!A1');
      }

      await new Promise(r => setTimeout(r, 2500));

      const valA1_afterManualEdit = is2 ? is2.getRange(0, 0, 1, 1).getValue() : 'NO_SHEET2';
      const valA2_afterManualEdit = is2 ? is2.getRange(1, 0, 1, 1).getValue() : 'NO_SHEET2';
      const valA3_afterManualEdit = is2 ? is2.getRange(2, 0, 1, 1).getValue() : 'NO_SHEET2';

      const cellA1_afterEdit_snapshot = impWb.save().sheets['Sheet2']?.cellData?.[0]?.[0];

      console.log('Step 3: Creating Manual Workbook...');
      api.disposeUnit(impWb.getId());

      const manualWb = api.createUniverSheet({
        id: 'manual-wb',
        name: 'Manual',
        sheetOrder: ['s1', 's2'],
        sheets: {
          s1: { id: 's1', name: 'Sheet1', type: 2, status: 1, cellData: {} },
          s2: { id: 's2', name: 'Sheet2', type: 2, status: 0, cellData: {} }
        }
      });

      const ms1 = manualWb.getSheetByName('Sheet1');
      if (ms1) {
        ms1.getRange(0, 0, 1, 1).setValue(100);
        ms1.getRange(1, 0, 1, 1).setValue(200);
      }

      const ms2 = manualWb.getSheetByName('Sheet2');
      if (ms2) {
        ms2.getRange(0, 0, 1, 1).setValue('=Sheet1!A1');
      }

      await new Promise(r => setTimeout(r, 2500));

      const valA1_manualCaseA = ms2 ? ms2.getRange(0, 0, 1, 1).getValue() : 'NO_SHEET2';
      const cellA1_manualCaseA_snapshot = manualWb.save().sheets['s2']?.cellData?.[0]?.[0];

      return {
        importedState_Immediate: {
          Sheet2_A1_value: valA1_imported,
          Sheet2_A2_value: valA2_imported,
          Sheet2_A3_value: valA3_imported,
          Sheet2_A1_snapshotCell: cellA1_imported_snapshot,
          Sheet2_A3_snapshotCell: cellA3_imported_snapshot,
        },
        importedState_AfterManualEditOfA1: {
          Sheet2_A1_value: valA1_afterManualEdit,
          Sheet2_A2_value: valA2_afterManualEdit,
          Sheet2_A3_value: valA3_afterManualEdit,
          Sheet2_A1_snapshotCell: cellA1_afterEdit_snapshot,
          commandsFiredOnManualEdit: editCommands
        },
        manualCaseA_State: {
          Sheet2_A1_value: valA1_manualCaseA,
          Sheet2_A1_snapshotCell: cellA1_manualCaseA_snapshot,
        }
      };
    } catch (err) {
      console.log('EVAL ERR:', err.stack || err.message);
      return { error: err.message, stack: err.stack };
    }
  }, impWbDataWithoutV);

  fs.writeFileSync('no-cache-comparison-out.json', JSON.stringify(res, null, 2));
  console.log('=== RESULTS WRITTEN TO no-cache-comparison-out.json ===');
  console.log(JSON.stringify(res, null, 2));

  await browser.close();
}

testFormulaNoCacheComparison().catch(console.error);
