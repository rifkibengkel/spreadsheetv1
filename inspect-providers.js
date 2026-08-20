const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({ headless: 'new' });
  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');
  await new Promise(r => setTimeout(r, 1000));
  
  const info = await page.evaluate(() => {
    const api = window.univerAPI;
    const activeWb = api.getActiveWorkbook();
    
    // Inspect all services available in injector
    const injector = api._injector || (window.univer && window.univer._injector);
    
    // Search injector providers for command/formula services
    const providers = [];
    if (injector && injector._providers) {
      injector._providers.forEach((val, key) => {
        const name = typeof key === 'symbol' ? key.toString() : String(key);
        providers.push(name);
      });
    }

    return { providers };
  });

  console.log('INJECTOR PROVIDERS:', JSON.stringify(info.providers, null, 2));
  await browser.close();
})();
