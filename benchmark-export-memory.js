const fs = require('fs');
const puppeteer = require('puppeteer');

const EXCEL_PATH = 'C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx';

(async () => {
  console.log('=== Starting Safe Export Pipeline Memory Benchmark ===');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=4096', '--js-flags=--expose-gc']
  });

  const page = await browser.newPage();

  page.on('console', msg => console.log('[PAGE LOG]', msg.text()));

  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  // Open modal & upload updatehexos.xlsx
  const buttons = await page.$$('button');
  let importBtn = null;
  for (const b of buttons) {
    const text = await page.evaluate(el => el.textContent, b);
    if (text && text.includes('Import File')) {
      importBtn = b;
      break;
    }
  }
  await importBtn.click();
  await page.waitForSelector('input[type="file"]', { visible: false });

  const fileInput = await page.$('input[type="file"]');
  await fileInput.uploadFile(EXCEL_PATH);

  const modalButtons = await page.$$('button');
  let startBtn = null;
  for (const b of modalButtons) {
    const text = await page.evaluate(el => el.textContent, b);
    if (text && text.includes('Mulai Impor')) {
      startBtn = b;
      break;
    }
  }
  await startBtn.click();

  console.log('Waiting for updatehexos.xlsx to mount...');

  // Wait for workbook to mount with NO_CALCULATION
  await page.evaluate(async () => {
    const api = window.univerAPI;
    let attempts = 0;
    while (attempts < 120) {
      const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : null;
      if (wb && wb.getSheets && wb.getSheets().length > 1) break;
      await new Promise(r => setTimeout(r, 1000));
      attempts++;
    }
  });

  console.log('Workbook mounted. Running Dry-Run Export Memory Benchmark...');

  const benchmarkResult = await page.evaluate(async () => {
    const getMem = () => {
      if (window.performance && window.performance.memory) {
        return {
          usedMB: (window.performance.memory.usedJSHeapSize / 1024 / 1024).toFixed(2),
          totalMB: (window.performance.memory.totalJSHeapSize / 1024 / 1024).toFixed(2),
          limitMB: (window.performance.memory.jsHeapSizeLimit / 1024 / 1024).toFixed(2),
        };
      }
      return { usedMB: 'N/A', totalMB: 'N/A', limitMB: 'N/A' };
    };

    const timeline = [];
    timeline.push({ stage: '0_initial_loaded', mem: getMem() });

    const api = window.univerAPI;
    const activeWb = api.getActiveWorkbook();
    if (!activeWb) return { error: 'No active workbook' };

    // Stage 1: Measure memory after workbook.save() snapshot
    const t1 = Date.now();
    const snapshot = activeWb.save();
    const snapshotTimeMs = Date.now() - t1;
    timeline.push({ stage: '1_after_univer_save_snapshot', timeMs: snapshotTimeMs, mem: getMem() });

    // Estimate size of snapshot object structure
    let sheetCount = 0;
    let cellCount = 0;
    if (snapshot.sheets) {
      sheetCount = Object.keys(snapshot.sheets).length;
      Object.keys(snapshot.sheets).forEach(s => {
        if (snapshot.sheets[s].cellData) {
          Object.keys(snapshot.sheets[s].cellData).forEach(r => {
            cellCount += Object.keys(snapshot.sheets[s].cellData[r] || {}).length;
          });
        }
      });
    }

    // Stage 2: Measure memory overhead of stringifying snapshot (JSON.stringify dry run)
    const t2 = Date.now();
    let jsonLength = 0;
    try {
      const jsonStr = JSON.stringify(snapshot);
      jsonLength = jsonStr.length;
      timeline.push({ stage: '2_during_json_stringify_snapshot', jsonSizeMB: (jsonLength / 1024 / 1024).toFixed(2), timeMs: Date.now() - t2, mem: getMem() });
    } catch (e) {
      timeline.push({ stage: '2_during_json_stringify_snapshot_FAILED', error: e.message, mem: getMem() });
    }

    return {
      sheetCount,
      cellCount,
      jsonSizeMB: (jsonLength / 1024 / 1024).toFixed(2),
      timeline
    };
  });

  console.log('Benchmark Timeline:', JSON.stringify(benchmarkResult, null, 2));

  fs.writeFileSync('export-memory-benchmark-summary.json', JSON.stringify(benchmarkResult, null, 2));
  await browser.close();
})();
