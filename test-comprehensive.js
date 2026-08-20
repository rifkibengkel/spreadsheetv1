const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox', '--disable-setuid-sandbox'] });
  const page = await browser.newPage();
  
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  
  await new Promise(r => setTimeout(r, 4000));

  const results = await page.evaluate(async () => {
    const api = window.univerAPI;
    const out = {};
    const univerInst = window.univer || api._univer;
    const injector = api._injector || (window.univer && window.univer._injector);

    // Question 3 & 4 (Before)
    let applyCtrl = null;
    let formulaModel = null;
    try {
      applyCtrl = injector.get('CalculateResultApplyController') || injector.get('calculate-result-apply-controller');
      formulaModel = injector.get('FormulaDataModel') || injector.get('formula-data-model');
    } catch(e) {}
    out.Q3_applyCtrl_before = !!applyCtrl;
    out.Q4_formulaModel_before = !!formulaModel;

    // Mimic Excel Import (Only static values)
    const importedData = {
      id: 'imported-wb-q',
      sheetOrder: ['sheet1'],
      name: 'QTest',
      sheets: {
        'sheet1': {
          id: 'sheet1', name: 'Sheet1', type: 2, rowCount: 10, columnCount: 10,
          cellData: {
            '0': { '0': { v: 100, t: 2 } }, // A1
            '1': { '0': { v: 200, t: 2 } }  // A2
          }
        }
      }
    };
    
    // Dispose old, create new
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
    oldWbs.forEach(w => { if(w) api.disposeUnit(w.getId()); });
    
    api.createUniverSheet(importedData);
    await new Promise(r => setTimeout(r, 2000));
    
    // Question 6: Does createUniverSheet replace injector?
    const injectorAfter = api._injector || (window.univer && window.univer._injector);
    out.Q6_same_injector = (injector === injectorAfter);
    
    // Question 3 & 4 (After)
    let applyCtrlAfter = null;
    let formulaModelAfter = null;
    try {
      applyCtrlAfter = injectorAfter.get('CalculateResultApplyController') || injectorAfter.get('calculate-result-apply-controller');
      formulaModelAfter = injectorAfter.get('FormulaDataModel') || injectorAfter.get('formula-data-model');
    } catch(e) {}
    out.Q3_applyCtrl_after = !!applyCtrlAfter;
    out.Q4_formulaModel_after = !!formulaModelAfter;

    const wb = api.getActiveWorkbook();
    const sheet = wb.getActiveSheet();

    // Question 2: Does an imported sheet with ONLY static values calculate manually entered formula?
    await api.executeCommand('sheet.command.set-range-values', {
      unitId: wb.getId(), subUnitId: sheet.getSheetId(),
      range: { startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 },
      value: { f: '=A1+A2' } // A3
    });
    await new Promise(r => setTimeout(r, 2000));
    out.Q2_formula_result_on_static_import = sheet.getRange(2, 0, 1, 1).getValue();


    
    return out;
  });

  console.log(JSON.stringify(results, null, 2));
  await browser.close();
  process.exit(0);
})();
