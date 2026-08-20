const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({ headless: 'new' });
  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  const inspection = await page.evaluate(() => {
    const api = window.univerAPI;
    const univer = window.univer;

    const apiMethods = Object.keys(api);
    const apiProtoMethods = Object.getOwnPropertyNames(Object.getPrototypeOf(api));

    return {
      apiMethods,
      apiProtoMethods,
      hasExecuteCommand: typeof api.executeCommand === 'function',
      hasExecuteCommandOnUniver: typeof univer?.executeCommand === 'function',
      hasCommandServiceOnFormula: !!api.getFormula()?._commandService
    };
  });

  console.log('Univer Inspection Results:');
  console.log(JSON.stringify(inspection, null, 2));

  await browser.close();
})();
