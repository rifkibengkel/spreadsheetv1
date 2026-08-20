const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Forensic Profiling: User Cell Edit Pipeline in updatehexos.xlsx ===');

  const fileBuffer = fs.readFileSync('C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx');
  const base64Data = fileBuffer.toString('base64');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=8192']
  });

  const page = await browser.newPage();

  // Capture all page logs & timestamps
  page.on('console', msg => {
    const text = msg.text();
    if (text.includes('[Formula') || text.includes('[Trigger') || text.includes('[Profile]')) {
      console.log(`[PAGE LOG ${performance.now().toFixed(2)}ms]`, text);
    }
  });

  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Importing updatehexos.xlsx...');
  const tImportStart = performance.now();

  const sheetSummary = await page.evaluate(async (b64) => {
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

    const wb = api.getActiveWorkbook();
    const sheets = wb.getSheets().map(s => ({
      id: s.getSheetId(),
      name: s.getName ? s.getName() : (s.getSheetName ? s.getSheetName() : 'Sheet')
    }));

    return { sheets };
  }, base64Data);

  const tImportEnd = performance.now();
  console.log(`Workbook imported in ${((tImportEnd - tImportStart)/1000).toFixed(2)}s. Sheets:`, sheetSummary.sheets);

  console.log('\n--- Starting Forensic Profiling of Cell Edit ---');

  const profileData = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook();

    // Find active sheet
    const activeSheet = wb.getActiveSheet();
    const sheetName = activeSheet.getName ? activeSheet.getName() : (activeSheet.getSheetName ? activeSheet.getSheetName() : 'Sheet');

    const commandEvents = [];
    const commandService = api.getCommandService ? api.getCommandService() : null;

    if (commandService) {
      commandService.onCommandExecuted((command, options) => {
        const now = performance.now();
        let detail = null;

        if (command.id === 'sheet.mutation.set-formula-calculation-start') {
          detail = { dirtyRanges: command.params?.dirtyRanges };
        } else if (command.id === 'sheet.mutation.set-formula-calculation-notification') {
          detail = {
            stageInfo: command.params?.stageInfo,
            functionsExecutedState: command.params?.functionsExecutedState
          };
        } else if (command.id === 'sheet.mutation.set-formula-calculation-result') {
          const unitData = command.params?.unitData || {};
          let resultCellCount = 0;
          Object.keys(unitData).forEach(u => {
            Object.keys(unitData[u] || {}).forEach(s => {
              const matrix = unitData[u][s];
              if (matrix) {
                Object.keys(matrix).forEach(r => {
                  resultCellCount += Object.keys(matrix[r] || {}).length;
                });
              }
            });
          });
          const payloadSizeEst = JSON.stringify(unitData).length;
          detail = { resultCellCount, payloadSizeEst };
        } else if (command.id === 'sheet.mutation.set-range-values') {
          detail = {
            trigger: command.params?.trigger,
            cellValue: command.params?.cellValue,
            applyFormulaCalculationResult: options?.applyFormulaCalculationResult
          };
        }

        commandEvents.push({
          timestamp: now,
          id: command.id,
          detail
        });
      });
    }

    // Measure stage 1: User Edit Trigger
    const tUserEditStart = performance.now();

    // Change cell (0,0) in active sheet
    const targetRange = activeSheet.getRange(0, 0, 1, 1);
    targetRange.setValue(999999);

    const tMutationDone = performance.now();

    // Wait for calculation completion or timeout up to 15 seconds
    let tCalculationComplete = 0;
    for (let i = 0; i < 300; i++) {
      await new Promise(r => setTimeout(r, 50));
      const lastNotification = commandEvents.slice().reverse().find(e => e.id === 'sheet.mutation.set-formula-calculation-notification');
      if (lastNotification && lastNotification.detail?.functionsExecutedState === 2) {
        tCalculationComplete = performance.now();
        break;
      }
    }
    if (tCalculationComplete === 0) tCalculationComplete = performance.now();

    await new Promise(r => requestAnimationFrame(r));
    const tRenderDone = performance.now();

    return {
      sheetName,
      tUserEditStart,
      tMutationDone,
      tCalculationComplete,
      tRenderDone,
      commandEvents
    };
  });

  console.log('\n--- PROFILING RESULTS SUMMARY ---');
  console.log(`Sheet Edited: ${profileData.sheetName}`);
  console.log(`Stage 1 (SetRangeValues synchronous execution): ${(profileData.tMutationDone - profileData.tUserEditStart).toFixed(2)} ms`);
  console.log(`Stage 2 (Total edit to calculation finish): ${(profileData.tCalculationComplete - profileData.tUserEditStart).toFixed(2)} ms`);
  console.log(`Stage 3 (Total edit to UI render complete): ${(profileData.tRenderDone - profileData.tUserEditStart).toFixed(2)} ms`);

  console.log('\n--- COMMAND EXECUTION TIMELINE ---');
  const t0 = profileData.tUserEditStart;
  profileData.commandEvents.forEach(e => {
    console.log(`+${(e.timestamp - t0).toFixed(2)}ms | Command: ${e.id} | Detail:`, JSON.stringify(e.detail));
  });

  fs.writeFileSync('edit-profile-results.json', JSON.stringify(profileData, null, 2));

  await browser.close();
})();
