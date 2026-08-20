const puppeteer = require('puppeteer');

(async () => {
  console.log('Starting POC Formula Test...');
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox', '--disable-setuid-sandbox'] });
  const page = await browser.newPage();

  page.on('console', msg => {
    const text = msg.text();
    if (text.includes('[POC]') || text.includes('ERROR') || text.includes('error')) {
      console.log(`[Browser] ${text}`);
    }
  });

  await page.goto('http://localhost:3000', { waitUntil: 'networkidle2' });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });

  // Wait for auto-test to complete (it runs at 2s + ~2s internal delays)
  await new Promise(r => setTimeout(r, 6000));

  // Now run our own comprehensive test
  const results = await page.evaluate(async () => {
    const api = window.univerAPI;
    if (!api) return { error: 'No API' };

    const wb = api.getActiveWorkbook();
    if (!wb) return { error: 'No Workbook' };

    const unitId = wb.getId();
    const subUnitId = wb.getActiveSheet().getSheetId();
    const cmd = (range, value) =>
      api.executeCommand('sheet.command.set-range-values', { unitId, subUnitId, range, value });
    const r = (sr, sc) => ({ startRow: sr, endRow: sr, startColumn: sc, endColumn: sc });
    const getVal = (row, col) => wb.getActiveSheet().getRange(row, col, 1, 1).getValue();
    const getFormula = (row, col) => wb.getActiveSheet().getRange(row, col, 1, 1).getFormula();

    const out = {};

    // === TEST 1: Constant formula ===
    await cmd(r(0, 0), { f: '=1+1' });
    await new Promise(r => setTimeout(r, 1500));
    out.test1_constant = {
      value: getVal(0, 0),
      formula: getFormula(0, 0),
      expected_value: 2,
      expected_formula: '=1+1'
    };

    // === TEST 2: Normal value vs formula ===
    await cmd(r(0, 0), { v: 100, t: 2 });
    await cmd(r(1, 0), { v: 200, t: 2 });
    await cmd(r(2, 0), { f: '=A1+A2' });
    await new Promise(r => setTimeout(r, 1500));
    out.test2_formula = {
      A1_value: getVal(0, 0),
      A2_value: getVal(1, 0),
      A3_value: getVal(2, 0),
      A3_formula: getFormula(2, 0),
      expected_A3_value: 300,
      expected_A3_formula: '=A1+A2'
    };

    // === TEST 3: Dependency update - change A1 to 500 ===
    await cmd(r(0, 0), { v: 500, t: 2 });
    await new Promise(r => setTimeout(r, 1500));
    out.test3_dep_update_1 = {
      A1_value: getVal(0, 0),
      A3_value: getVal(2, 0),
      A3_formula: getFormula(2, 0),
      expected_A3_value: 700,
      expected_A3_formula: '=A1+A2'
    };

    // === TEST 4: Dependency update - change A2 to 300 ===
    await cmd(r(1, 0), { v: 300, t: 2 });
    await new Promise(r => setTimeout(r, 1500));
    out.test4_dep_update_2 = {
      A2_value: getVal(1, 0),
      A3_value: getVal(2, 0),
      A3_formula: getFormula(2, 0),
      expected_A3_value: 800,
      expected_A3_formula: '=A1+A2'
    };

    // === TEST 5: Cross Sheet ===
    const sheet1Id = wb.getActiveSheet().getSheetId();
    // Simulate creating a new sheet (or just use existing if available, Univer starts with 1 sheet usually)
    // Wait, let's just create a new sheet natively
    await api.executeCommand('sheet.command.add-worksheet');
    await new Promise(r => setTimeout(r, 1000));
    const sheet2Id = wb.getActiveSheet().getSheetId();

    // In Sheet1, A1 is 500. Let's make sure.
    await api.executeCommand('sheet.command.set-range-values', { unitId, subUnitId: sheet1Id, range: r(0, 0), value: { v: 500, t: 2 } });
    
    // In Sheet2, A1 = =Sheet1!A1
    await api.executeCommand('sheet.command.set-range-values', { unitId, subUnitId: sheet2Id, range: r(0, 0), value: { f: '=Sheet1!A1' } });
    await new Promise(r => setTimeout(r, 2000));

    const sheet2A1Val = wb.getSheetBySheetId(sheet2Id).getRange(0, 0, 1, 1).getValue();
    const sheet2A1Formula = wb.getSheetBySheetId(sheet2Id).getRange(0, 0, 1, 1).getFormula();

    out.test5_cross_sheet = {
      Sheet1_A1_value: wb.getSheetBySheetId(sheet1Id).getRange(0, 0, 1, 1).getValue(),
      Sheet2_A1_value: sheet2A1Val,
      Sheet2_A1_formula: sheet2A1Formula,
      expected_Sheet2_A1_value: 500,
      expected_Sheet2_A1_formula: '=Sheet1!A1'
    };

    // Update Sheet1!A1 to 999
    await api.executeCommand('sheet.command.set-range-values', { unitId, subUnitId: sheet1Id, range: r(0, 0), value: { v: 999, t: 2 } });
    await new Promise(r => setTimeout(r, 2000));
    
    out.test5_cross_sheet_update = {
      Sheet1_A1_value: wb.getSheetBySheetId(sheet1Id).getRange(0, 0, 1, 1).getValue(),
      Sheet2_A1_value: wb.getSheetBySheetId(sheet2Id).getRange(0, 0, 1, 1).getValue(),
      Sheet2_A1_formula: wb.getSheetBySheetId(sheet2Id).getRange(0, 0, 1, 1).getFormula(),
      expected_Sheet2_A1_value: 999,
      expected_Sheet2_A1_formula: '=Sheet1!A1'
    };

    return out;
  });

  require('fs').writeFileSync('poc-results.json', JSON.stringify(results, null, 2));
  console.log('\n=== POC RESULTS ===');
  console.log(JSON.stringify(results, null, 2));

  await browser.close();
  process.exit(0);
})();
