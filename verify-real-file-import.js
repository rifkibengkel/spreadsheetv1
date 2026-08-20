const fs = require('fs');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

async function verifyRealFileImport() {
  console.log('=====================================================================');
  console.log('VERIFYING REAL FILE IMPORT (C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx)');
  console.log('=====================================================================');

  const barakatPath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  const fileBuf = fs.readFileSync(barakatPath);

  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();

  const workerLogs = [];
  page.on('console', msg => {
    const txt = msg.text();
    if (txt.includes('[Formula Worker]') || txt.includes('[PROGRESS]')) {
      workerLogs.push(txt);
      console.log('[BROWSER]', txt);
    }
  });

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Feeding report barakat.xlsx arrayBuffer to Web Worker...');
  const startImportTime = Date.now();

  const testResults = await page.evaluate(async (arrayBuf) => {
    const api = window.univerAPI;
    const worker = new Worker('/_next/static/chunks/src_lib_workers_excel_worker_ts.js', { type: 'module' });

    return new Promise((resolve) => {
      worker.onmessage = async (e) => {
        if (e.data.type === 'COMPLETE') {
          const wbData = e.data.workbookData;

          // Dispose existing workbooks
          const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
          oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(err){} });

          // Create Univer Sheet
          const wb = api.createUniverSheet(wbData);

          // Allow Univer Formula Engine 3 seconds to process formula queue
          await new Promise(r => setTimeout(r, 3000));

          // Check key test cells
          const sSummary = wb.getSheetByName('Summary');
          const sHadiah = wb.getSheetByName('Summary Hadiah');

          const sameSheetCell = sSummary ? sSummary.getRange(8, 8, 1, 1) : null;
          const crossSheetCell = sHadiah ? sHadiah.getRange(0, 0, 1, 1) : null;
          const largeFormulaCell = sHadiah ? sHadiah.getRange(3, 2, 1, 1) : null;
          const sharedFollowerCell = sHadiah ? sHadiah.getRange(4, 2, 1, 1) : null;

          const snap = wb.save();

          worker.terminate();
          resolve({
            success: true,
            totalSheets: wbData.sheetOrder.length,
            sameSheet: {
              name: 'Summary!I9',
              formula: sameSheetCell ? sameSheetCell.getFormula() : null,
              value: sameSheetCell ? sameSheetCell.getValue() : null
            },
            crossSheet: {
              name: 'Summary Hadiah!A1',
              formula: crossSheetCell ? crossSheetCell.getFormula() : null,
              value: crossSheetCell ? crossSheetCell.getValue() : null
            },
            largeFormula: {
              name: 'Summary Hadiah!C4',
              formula: largeFormulaCell ? largeFormulaCell.getFormula() : null,
              value: largeFormulaCell ? largeFormulaCell.getValue() : null
            },
            sharedFollower: {
              name: 'Summary Hadiah!C5',
              formula: sharedFollowerCell ? sharedFollowerCell.getFormula() : null,
              value: sharedFollowerCell ? sharedFollowerCell.getValue() : null
            }
          });
        }
      };

      worker.postMessage({ fileArrayBuffer: arrayBuf, fileName: 'report barakat.xlsx' });
    });
  }, fileBuf.buffer);

  const importDuration = ((Date.now() - startImportTime) / 1000).toFixed(2);
  console.log(`\nImport Completed in ${importDuration} seconds!`);
  console.log('Test Results:', JSON.stringify(testResults, null, 2));

  fs.writeFileSync('c:\\Users\\BKKI-1\\.gemini\\antigravity\\scratch\\sheet-V-2\\real-file-verification-output.json', JSON.stringify({
    importDurationSeconds: parseFloat(importDuration),
    testResults,
    workerLogs
  }, null, 2));

  await browser.close();
}

verifyRealFileImport().catch(console.error);
