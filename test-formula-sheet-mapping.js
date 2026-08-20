const fs = require('fs');
const puppeteer = require('puppeteer');

async function testFormulaSheetMapping() {
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  page.on('console', msg => console.log('[PAGE]', msg.text()));

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  const result = await page.evaluate(async () => {
    const api = window.univerAPI;
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
    oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

    // TEST SCENARIO 1: sheet.id != sheet.name
    // Sheet1: id="s1", name="Sheet1"
    // Sheet2: id="s2", name="Sheet2"
    console.log('Testing Scenario 1: sheet.id != sheet.name...');
    const wb1 = api.createUniverSheet({
      id: 'wb1',
      name: 'Workbook 1',
      sheetOrder: ['s1', 's2'],
      sheets: {
        s1: { id: 's1', name: 'Sheet1', type: 2, status: 1, cellData: { 0: { 0: { v: 100, t: 2 } } } },
        s2: { id: 's2', name: 'Sheet2', type: 2, status: 0, cellData: { 0: { 0: { f: '=Sheet1!A1' } } } }
      }
    });

    await new Promise(r => setTimeout(r, 1500));
    const s2_wb1 = wb1.getSheetByName('Sheet2');
    const val1_unquoted = s2_wb1.getRange(0, 0, 1, 1).getValue();

    s2_wb1.getRange(0, 1, 1, 1).setValue("='Sheet1'!A1");
    s2_wb1.getRange(0, 2, 1, 1).setValue("=s1!A1");

    await new Promise(r => setTimeout(r, 1500));
    const val1_quoted = s2_wb1.getRange(0, 1, 1, 1).getValue();
    const val1_by_id = s2_wb1.getRange(0, 2, 1, 1).getValue();

    // TEST SCENARIO 2: sheet.id == sheet.name
    // Sheet1: id="Sheet1", name="Sheet1"
    // Sheet2: id="Sheet2", name="Sheet2"
    console.log('Testing Scenario 2: sheet.id == sheet.name...');
    api.disposeUnit(wb1.getId());

    const wb2 = api.createUniverSheet({
      id: 'wb2',
      name: 'Workbook 2',
      sheetOrder: ['Sheet1', 'Sheet2'],
      sheets: {
        Sheet1: { id: 'Sheet1', name: 'Sheet1', type: 2, status: 1, cellData: { 0: { 0: { v: 100, t: 2 } } } },
        Sheet2: { id: 'Sheet2', name: 'Sheet2', type: 2, status: 0, cellData: { 0: { 0: { f: '=Sheet1!A1' } } } }
      }
    });

    await new Promise(r => setTimeout(r, 1500));
    const s2_wb2 = wb2.getSheetByName('Sheet2');
    const val2_unquoted = s2_wb2.getRange(0, 0, 1, 1).getValue();

    // TEST SCENARIO 3: Sheet name with spaces: "Summary Hadiah"
    // Sheet1: id="Summary Hadiah", name="Summary Hadiah"
    // Sheet2: id="Sheet2", name="Sheet2"
    console.log('Testing Scenario 3: Sheet name with spaces...');
    api.disposeUnit(wb2.getId());

    const wb3 = api.createUniverSheet({
      id: 'wb3',
      name: 'Workbook 3',
      sheetOrder: ['Summary Hadiah', 'Sheet2'],
      sheets: {
        'Summary Hadiah': { id: 'Summary Hadiah', name: 'Summary Hadiah', type: 2, status: 1, cellData: { 0: { 0: { v: 500, t: 2 } } } },
        'Sheet2': { 
          id: 'Sheet2', 
          name: 'Sheet2', 
          type: 2, 
          status: 0, 
          cellData: { 
            0: { 0: { f: "='Summary Hadiah'!A1" } },
            1: { 0: { f: "=Summary Hadiah!A1" } }
          } 
        }
      }
    });

    await new Promise(r => setTimeout(r, 1500));
    const s2_wb3 = wb3.getSheetByName('Sheet2');
    const val3_quoted = s2_wb3.getRange(0, 0, 1, 1).getValue();
    const val3_unquoted = s2_wb3.getRange(1, 0, 1, 1).getValue();

    return {
      scenario1_id_ne_name: {
        formula_ref_by_name_unquoted: val1_unquoted,
        formula_ref_by_name_quoted: val1_quoted,
        formula_ref_by_id: val1_by_id
      },
      scenario2_id_eq_name: {
        formula_ref_by_name: val2_unquoted
      },
      scenario3_with_spaces: {
        formula_ref_quoted: val3_quoted,
        formula_ref_unquoted: val3_unquoted
      }
    };
  });

  fs.writeFileSync('mapping-results.json', JSON.stringify(result, null, 2));
  console.log('MAPPING TEST RESULTS WRITTEN TO mapping-results.json');
  await browser.close();
}

testFormulaSheetMapping().catch(console.error);
