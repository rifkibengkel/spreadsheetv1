const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({ headless: 'new' });
  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');
  await new Promise(r => setTimeout(r, 1000));
  
  const res = await page.evaluate(async () => {
    const api = window.univerAPI;
    
    const commandsToTest = [
      'sheet.command.set-range-values',
      'formula.command.set-formula-calculation-result',
      'formula.command.execute-formula',
      'formula.command.calculate',
      'sheet.command.recalculate',
      'formula.command.recalculate',
      'formula.mutation.set-formula-calculation-result',
      'formula.command.set-calculation-result',
      'sheet.mutation.set-range-values'
    ];

    const commandStatus = {};
    for (const cmd of commandsToTest) {
      try {
        await api.executeCommand(cmd, {});
        commandStatus[cmd] = 'EXECUTED_OR_EMPTY_PARAMS';
      } catch (err) {
        const msg = err?.message || String(err);
        if (msg.includes('not registered') || msg.includes('is not defined')) {
          commandStatus[cmd] = 'NOT_REGISTERED';
        } else {
          commandStatus[cmd] = `REGISTERED: ${msg.split('\n')[0]}`;
        }
      }
    }

    return commandStatus;
  });

  fs.writeFileSync('commands-audit.json', JSON.stringify(res, null, 2));
  console.log('Saved commands-audit.json');
  await browser.close();
})();
