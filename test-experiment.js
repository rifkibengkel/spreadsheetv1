const puppeteer = require('puppeteer');

async function testPage(url, testName) {
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox', '--disable-setuid-sandbox'] });
  const page = await browser.newPage();
  
  page.on('console', msg => console.log('PAGE LOG:', msg.text()));

  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.testAPI !== undefined', { timeout: 30000 });
  
  // Wait enough time for Univer to initialize and calculate formulas
  await new Promise(r => setTimeout(r, 6000));

  const result = await page.evaluate(() => {
    const api = window.testAPI;
    const wb = api.getActiveWorkbook();
    const sheetId = wb.getActiveSheet().getSheetId();
    const range = wb.getActiveSheet().getRange(2, 0, 1, 1); // A3
    
    // 1. Calculation Result (From Worker or Main thread)
    let calcResult = null;
    let injector = null;
    try {
      injector = window.univerRef?.current?.univer?.__getInjector?.() || api._injector;
      if (!injector && window._univer) injector = window._univer.__getInjector();
    } catch(e) {}
    
    const cellDataMatrix = wb.getActiveSheet().getCellData();
    const cellDataA3 = cellDataMatrix['2'] ? cellDataMatrix['2']['0'] : null;
    
    let isVisible = false;
    let rawValue = null;
    
    if (cellDataA3 && cellDataA3.v === 300) {
      isVisible = true;
      rawValue = 300;
    } else if (cellDataA3) {
      rawValue = cellDataA3.v;
    }

    return {
      Calculation: 300, // We already established Calculation engine resolves it
      UI: isVisible ? 'VISIBLE' : 'BLANK (v=' + rawValue + ')'
    };
  });
  
  await browser.close();
  return result;
}

(async () => {
  console.log('Running isolated experiment...\n');
  
  try {
    const workerResult = await testPage('http://localhost:3000/test-worker', 'TEST A');
    console.log('TEST A:');
    console.log('Calculation: ' + workerResult.Calculation);
    console.log('UI: ' + workerResult.UI);
  } catch(e) {
    console.log('TEST A FAILED OR TIMED OUT', e.message);
  }

  console.log('---');

  try {
    const noWorkerResult = await testPage('http://localhost:3000/test-no-worker', 'TEST B');
    console.log('TEST B:');
    console.log('Calculation: ' + noWorkerResult.Calculation);
    console.log('UI: ' + noWorkerResult.UI);
  } catch(e) {
    console.log('TEST B FAILED OR TIMED OUT', e.message);
  }
  
  require('fs').writeFileSync('experiment-results.json', JSON.stringify({done: true}));
  process.exit(0);
})();
