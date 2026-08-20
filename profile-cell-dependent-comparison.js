const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Forensic Comparison: Simple Dependent Cell vs Massive 496,610-Row Array Formula Dependent Cell ===');

  const fileBuffer = fs.readFileSync('C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx');
  const base64Data = fileBuffer.toString('base64');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=8192']
  });

  const page = await browser.newPage();
  page.on('console', msg => {
    const text = msg.text();
    if (text.includes('SetFormula') || text.includes('Formula Worker') || text.includes('Profile')) {
      console.log(`[PAGE LOG ${performance.now().toFixed(2)}ms]`, text);
    }
  });

  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Importing updatehexos.xlsx...');

  await page.evaluate(async (b64) => {
    const binaryString = atob(b64);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const file = new File([blob], 'updatehexos.xlsx');

    const api = window.univerAPI;
    const result = await window.importExcelToWorkbookDataAsync(file);
    window.replaceUniverWorkbook(api, result.workbookData);
  }, base64Data);

  console.log('Import finished. Executing Test 1 (Isolated cell edit) vs Test 2 (496k-row array formula range cell edit)...');

  const comparisonResult = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook();

    // Test 1: Isolated cell (e.g. Sheet 12 cell A1 or Sheet 9 cell A1)
    const sheets = wb.getSheets();
    const sheet9 = sheets[8] || wb.getActiveSheet();

    const commandEvents = [];
    const commandService = api.getCommandService ? api.getCommandService() : null;
    if (commandService) {
      commandService.onCommandExecuted((command, options) => {
        if (command.id === 'sheet.mutation.set-formula-calculation-result') {
          commandEvents.push({
            time: performance.now(),
            id: command.id,
            params: command.params
          });
        }
      });
    }

    // --- TEST 1: ISOLATED CELL EDIT ---
    const t0_1 = performance.now();
    sheet9.getRange(0, 0, 1, 1).setValue(999);
    const t1_1 = performance.now();

    await new Promise(r => setTimeout(r, 1500));
    const t2_1 = performance.now();

    // --- TEST 2: DATA SHEET COLUMN B CELL EDIT (Dependent on 496,610-row SUM(IF) Array Formula) ---
    const dataSheet = sheets.find(s => (s.getName ? s.getName() : s.getSheetName()).includes('Data')) || sheets[0];

    const t0_2 = performance.now();
    dataSheet.getRange(1, 1, 1, 1).setValue('TARGET_VAL');
    const t1_2 = performance.now();

    await new Promise(r => setTimeout(r, 18000));
    const t2_2 = performance.now();

    return {
      test1_IsolatedCell: {
        sheetName: sheet9.getName ? sheet9.getName() : 'Sheet9',
        mutationTimeMs: t1_1 - t0_1,
        totalTimeMs: t2_1 - t0_1
      },
      test2_DataCell_496kArrayFormula: {
        sheetName: dataSheet.getName ? dataSheet.getName() : 'DataSheet',
        mutationTimeMs: t1_2 - t0_2,
        totalTimeMs: t2_2 - t0_2
      },
      resultMutationCount: commandEvents.length
    };
  });

  console.log('\n--- FORENSIC COMPARISON RESULTS ---');
  console.log(JSON.stringify(comparisonResult, null, 2));

  fs.writeFileSync('isolated-vs-array-formula-profiling.json', JSON.stringify(comparisonResult, null, 2));

  await browser.close();
})();
