const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Testing Cell Editor & Export Synchronization ===');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=4096']
  });

  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  const result = await page.evaluate(async () => {
    const api = window.univerAPI;

    // 1. Set a cell value with rich text (p) vs scalar value (v)
    const activeWb = api.getActiveWorkbook();
    const sheet = activeWb.getActiveSheet();
    const sheetId = sheet.getSheetId();

    // Create a cell with rich text (p) but no scalar (v)
    const univerData = activeWb.save();
    univerData.sheets[sheetId].cellData[5] = {
      '5': {
        p: {
          body: {
            dataStream: 'Rich Text User Edit\r\n'
          }
        }
        // Notice: NO 'v' field!
      }
    };

    // Re-create or check what export.worker.ts extracts
    const cellPay = univerData.sheets[sheetId].cellData[5]['5'];

    // Current exporter logic:
    let currentExporterResult = undefined;
    if (cellPay.f) {
      currentExporterResult = cellPay.f;
    } else if (cellPay.v !== undefined && cellPay.v !== null) {
      currentExporterResult = cellPay.v;
    }

    // Improved exporter logic (extracting text from p if v is missing):
    let improvedExporterResult = undefined;
    if (cellPay.f) {
      improvedExporterResult = cellPay.f;
    } else if (cellPay.v !== undefined && cellPay.v !== null) {
      improvedExporterResult = cellPay.v;
    } else if (cellPay.p && cellPay.p.body && typeof cellPay.p.body.dataStream === 'string') {
      improvedExporterResult = cellPay.p.body.dataStream.replace(/[\r\n]+$/, '');
    }

    return {
      cellPay,
      currentExporterResult,
      improvedExporterResult
    };
  });

  console.log('Test Result:', JSON.stringify(result, null, 2));

  await browser.close();
})();
