const puppeteer = require('puppeteer');

(async () => {
  console.log('≡ƒÜÇ Starting Puppeteer Test Report...');
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox', '--disable-setuid-sandbox'] });
  const page = await browser.newPage();

  page.on('console', msg => {
    const text = msg.text();
    if (!text.includes('DOM element is missing') && !text.includes('React DevTools')) {
      // console.log(`[Browser] ${text}`);
    }
  });

  await page.goto('http://localhost:3000', { waitUntil: 'networkidle2' });

  // Wait for Univer to initialize
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });

  const results = await page.evaluate(async () => {
    const getUniver = () => window.univerAPI;
    const api = getUniver();
    if (!api) return { error: "No API found" };
    
    if (typeof window.__getDeepDiagnostic !== 'function') return { error: "No __getDeepDiagnostic function found" };

    const getDiagnostic = (row, col) => window.__getDeepDiagnostic(row, col);

    const out = {};

    // 1. TEST COSTANT A1 = =1+1
    console.log('[TEST] Writing =1+1');
    await api.executeCommand('sheet.command.set-range-values', { value: { v: '', f: '=1+1' }, range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 } });
    await new Promise(r => setTimeout(r, 1000));
    out.testConstant = getDiagnostic(0, 0);

    // 2. TEST VALUE (A2=100) VS FORMULA (B2 = =A2+10)
    console.log('[TEST] Writing Value vs Formula');
    await api.executeCommand('sheet.command.set-range-values', { value: { v: 100 }, range: { startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 } });
    await api.executeCommand('sheet.command.set-range-values', { value: { v: '', f: '=A2+10' }, range: { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 } });
    await new Promise(r => setTimeout(r, 1000));
    out.testValue = getDiagnostic(1, 0);
    out.testFormula = getDiagnostic(1, 1);

    // 3. DEPENDENCY UPDATE A3=100, A4=200, A5==A3+A4
    console.log('[TEST] Writing Dependency Updates');
    await api.executeCommand('sheet.command.set-range-values', { value: { v: 100 }, range: { startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 } });
    await api.executeCommand('sheet.command.set-range-values', { value: { v: 200 }, range: { startRow: 3, endRow: 3, startColumn: 0, endColumn: 0 } });
    await api.executeCommand('sheet.command.set-range-values', { value: { v: '', f: '=A3+A4' }, range: { startRow: 4, endRow: 4, startColumn: 0, endColumn: 0 } });
    await new Promise(r => setTimeout(r, 1500));
    out.dep1 = getDiagnostic(4, 0);

    await api.executeCommand('sheet.command.set-range-values', { value: { v: 500 }, range: { startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 } });
    await new Promise(r => setTimeout(r, 1500));
    out.dep2 = getDiagnostic(4, 0);

    return out;
  });

  console.log("=============== JSON RESULTS ===============");
  require('fs').writeFileSync('test-report.json', JSON.stringify(results, null, 2));
  console.log("Wrote test-report.json");

  await browser.close();
  process.exit(0);
})();
