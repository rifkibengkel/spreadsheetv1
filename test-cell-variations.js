const puppeteer = require('puppeteer');

async function runCellVariationsTest() {
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  await new Promise(r => setTimeout(r, 2000));

  const results = await page.evaluate(async () => {
    const api = window.univerAPI;

    // Test cases for A3 cell definition in imported workbookData
    const testCases = {
      'has_f_v_t': { f: '=A1+A2', v: 300, t: 2 },
      'has_f_only': { f: '=A1+A2' },
      'has_f_v_null': { f: '=A1+A2', v: null },
      'has_f_v_empty': { f: '=A1+A2', v: '', t: 1 },
      'has_f_t_only': { f: '=A1+A2', t: 2 },
      'has_f_no_equals': { f: 'A1+A2' },
      'has_f_lowercase': { f: '=a1+a2' }
    };

    const auditOutput = {};

    for (const [caseName, a3CellData] of Object.entries(testCases)) {
      // Dispose existing workbooks
      const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
      oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });

      const wbData = {
        id: `wb-${caseName}`,
        name: `Test-${caseName}`,
        sheetOrder: ['sheet1'],
        appVersion: '3.0.0-alpha',
        sheets: {
          'sheet1': {
            id: 'sheet1',
            name: 'Sheet1',
            type: 2,
            status: 1,
            rowCount: 20,
            columnCount: 10,
            cellData: {
              '0': { '0': { v: 100, t: 2 } }, // A1
              '1': { '0': { v: 200, t: 2 } }, // A2
              '2': { '0': a3CellData }        // A3
            }
          }
        }
      };

      api.createUniverSheet(wbData);
      const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : (api.getAllWorkbooks && api.getAllWorkbooks()[0]);
      const sheet = activeWb.getActiveSheet ? activeWb.getActiveSheet() : (activeWb.getSheetByName && activeWb.getSheetByName('Sheet1'));

      // Immediate check
      const immVal = sheet.getRange(2, 0, 1, 1).getValue();
      const immCell = activeWb.save().sheets['sheet1']?.cellData?.[2]?.[0];

      // Wait 2 seconds for formula calculation
      await new Promise(r => setTimeout(r, 2000));

      const setVal = sheet.getRange(2, 0, 1, 1).getValue();
      const setCell = activeWb.save().sheets['sheet1']?.cellData?.[2]?.[0];

      auditOutput[caseName] = {
        inputCell: a3CellData,
        immediate: { value: immVal, cell: immCell },
        settled: { value: setVal, cell: setCell }
      };
    }

    return auditOutput;
  });

  fs.writeFileSync('cell-variations-output.json', JSON.stringify(results, null, 2));
  console.log('✅ Variations audit saved to cell-variations-output.json');
  await browser.close();
}

runCellVariationsTest().catch(console.error);
