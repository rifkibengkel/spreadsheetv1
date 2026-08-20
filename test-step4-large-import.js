const fs = require('fs');
const puppeteer = require('puppeteer');

const EXCEL_PATH = 'C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx';

(async () => {
  console.log('Starting Step 4 Test: Inspect Large Workbook Without Calculating...');

  if (!fs.existsSync(EXCEL_PATH)) {
    console.error('File not found:', EXCEL_PATH);
    process.exit(1);
  }

  const fileStats = fs.statSync(EXCEL_PATH);
  console.log(`Excel file size: ${(fileStats.size / 1024 / 1024).toFixed(2)} MB`);

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=4096']
  });

  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Page loaded, window.univerAPI ready.');

  // Set CalculationMode.NO_CALCULATION (2)
  await page.evaluate(() => {
    const api = window.univerAPI;
    const formulaEngine = api.getFormula ? api.getFormula() : null;
    if (formulaEngine && formulaEngine.setInitialFormulaComputing) {
      formulaEngine.setInitialFormulaComputing(2); // NO_CALCULATION
    }
  });

  // Click "Import File" button to open ImportModal
  console.log('Opening Import Modal...');
  const buttons = await page.$$('button');
  let importBtn = null;
  for (const b of buttons) {
    const text = await page.evaluate(el => el.textContent, b);
    if (text && text.includes('Import File')) {
      importBtn = b;
      break;
    }
  }

  if (!importBtn) {
    console.error('Import button not found');
    await browser.close();
    process.exit(1);
  }

  await importBtn.click();
  await page.waitForSelector('input[type="file"]', { visible: false });

  // Upload file
  console.log('Uploading updatehexos.xlsx to file input...');
  const fileInput = await page.$('input[type="file"]');
  await fileInput.uploadFile(EXCEL_PATH);

  // Click "Mulai Impor" button
  console.log('Clicking Mulai Impor button...');
  const modalButtons = await page.$$('button');
  let startBtn = null;
  for (const b of modalButtons) {
    const text = await page.evaluate(el => el.textContent, b);
    if (text && text.includes('Mulai Impor')) {
      startBtn = b;
      break;
    }
  }

  if (!startBtn) {
    console.error('Mulai Impor button not found');
    await browser.close();
    process.exit(1);
  }

  const startImportTime = Date.now();
  await startBtn.click();

  console.log('Waiting for import and workbook mounting...');

  const importMetrics = await page.evaluate(async () => {
    const api = window.univerAPI;

    let attempts = 0;
    while (attempts < 120) {
      const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : null;
      if (wb) {
        const sheets = wb.getSheets ? wb.getSheets() : [];
        if (sheets.length > 1) {
          break;
        }
      }
      await new Promise(r => setTimeout(r, 1000));
      attempts++;
    }

    const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : null;
    if (!activeWb) return { error: 'Workbook not mounted' };

    const wbSave = activeWb.save();
    let formulaCount = 0;
    let cachedValCount = 0;
    let totalCellCount = 0;
    const sampleFormulas = [];

    if (wbSave.sheets) {
      Object.keys(wbSave.sheets).forEach(sheetId => {
        const sheet = wbSave.sheets[sheetId];
        if (sheet.cellData) {
          Object.keys(sheet.cellData).forEach(r => {
            const row = sheet.cellData[r];
            if (row) {
              Object.keys(row).forEach(c => {
                const cell = row[c];
                if (cell) {
                  totalCellCount++;
                  if (cell.f) {
                    formulaCount++;
                    if (sampleFormulas.length < 5) {
                      sampleFormulas.push({
                        sheetName: sheet.name,
                        row: r,
                        col: c,
                        f: cell.f,
                        v: cell.v
                      });
                    }
                  }
                  if (cell.v !== undefined && cell.v !== null) {
                    cachedValCount++;
                  }
                }
              });
            }
          });
        }
      });
    }

    let memoryInfo = null;
    if (window.performance && window.performance.memory) {
      const mem = window.performance.memory;
      memoryInfo = {
        totalJSHeapSize: (mem.totalJSHeapSize / 1024 / 1024).toFixed(2) + ' MB',
        usedJSHeapSize: (mem.usedJSHeapSize / 1024 / 1024).toFixed(2) + ' MB',
        jsHeapSizeLimit: (mem.jsHeapSizeLimit / 1024 / 1024).toFixed(2) + ' MB'
      };
    }

    const sheets = activeWb.getSheets ? activeWb.getSheets() : [];
    const sheetNames = sheets.map(s => {
      if (typeof s.getName === 'function') return s.getName();
      if (typeof s.getSheetName === 'function') return s.getSheetName();
      return s.name || s.id || 'Sheet';
    });

    return {
      sheetCount: sheets.length,
      sheetNames,
      formulaCount,
      cachedValCount,
      totalCellCount,
      sampleFormulas,
      memoryInfo
    };
  });

  const totalDurationMs = Date.now() - startImportTime;
  console.log(`Step 4 completed in ${totalDurationMs} ms.`);
  console.log('Metrics:', JSON.stringify(importMetrics, null, 2));

  fs.writeFileSync('step4-metrics.json', JSON.stringify({ totalDurationMs, importMetrics }, null, 2));
  await browser.close();
})();
