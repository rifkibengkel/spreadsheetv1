const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Deep Forensic Profiling: Tracing 16.3 Seconds Formula Calculation Bottleneck ===');

  const fileBuffer = fs.readFileSync('C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx');
  const base64Data = fileBuffer.toString('base64');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=8192']
  });

  const page = await browser.newPage();

  // Capture Worker & console logs
  page.on('console', msg => {
    const text = msg.text();
    if (text.includes('SetFormula') || text.includes('Formula') || text.includes('Trigger') || text.includes('execute')) {
      console.log(`[CONSOLE LOG ${performance.now().toFixed(2)}ms]`, text);
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

  console.log('Import finished. Setting up command listener on core Univer engine...');

  const deepLog = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook();
    const activeSheet = wb.getActiveSheet();

    // Get core commandService from Univer instance
    const univerObj = window.univer;
    const injector = univerObj.__getInjector ? univerObj.__getInjector() : null;
    let commandService = null;
    if (injector) {
      try {
        commandService = injector.get(window.univer.ICommandService || injector.get('ICommandService'));
      } catch(e) {}
    }

    const commandLogs = [];
    if (api.getCommandService) {
      commandService = api.getCommandService();
    }

    if (commandService && commandService.onCommandExecuted) {
      commandService.onCommandExecuted((command, options) => {
        const now = performance.now();
        commandLogs.push({
          time: now,
          id: command.id,
          paramsStr: JSON.stringify(command.params || {}).slice(0, 300)
        });
      });
    }

    // Step 1: User Edit Trigger
    const t0 = performance.now();
    activeSheet.getRange(0, 0, 1, 1).setValue(999999);
    const t1 = performance.now();

    // Wait 18 seconds while logging all commands
    await new Promise(r => setTimeout(r, 18000));
    const t2 = performance.now();

    return {
      t0,
      t1,
      t2,
      mutationMs: t1 - t0,
      totalMs: t2 - t0,
      commandLogs
    };
  });

  console.log('\n--- DEEP PROFILING COMMAND LOGS ---');
  console.log(`Mutation time: ${deepLog.mutationMs.toFixed(2)} ms`);
  console.log(`Total duration: ${deepLog.totalMs.toFixed(2)} ms`);
  console.log(`Total commands captured: ${deepLog.commandLogs.length}`);

  const t0 = deepLog.t0;
  deepLog.commandLogs.forEach(c => {
    console.log(`+${(c.time - t0).toFixed(2)}ms | ${c.id} | ${c.paramsStr}`);
  });

  fs.writeFileSync('deep-profile-results.json', JSON.stringify(deepLog, null, 2));

  await browser.close();
})();
