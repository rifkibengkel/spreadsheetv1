const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  page.on('console', msg => {
    const txt = msg.text();
    if (txt.includes('Replace workbook') || txt.includes('[Formula Worker]') || txt.includes('Calculation')) {
      console.log('PAGE LOG:', txt);
    }
  });

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  await new Promise(r => setTimeout(r, 1000));

  const testPayload = {
    id: `wb-test-${Date.now()}`,
    name: 'test-real-import.xlsx',
    sheetOrder: ['sheet1'],
    appVersion: '3.0.0-alpha',
    sheets: {
      'sheet1': {
        id: 'sheet1',
        name: 'Sheet1',
        type: 2,
        status: 1,
        rowCount: 30,
        columnCount: 20,
        cellData: {
          '0': { '0': { v: 100, t: 2 } }, // A1
          '1': { '0': { v: 200, t: 2 } }, // A2
          '2': { '0': { f: '=A1+A2' } }   // A3 = =A1+A2 (NO v, NO t)
        }
      }
    }
  };

  const res = await page.evaluate(async (payload) => {
    const api = window.univerAPI;
    
    // Dispose old units
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
    oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });

    // Call createUniverSheet
    api.createUniverSheet(payload);

    const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
    const sheet = activeWb.getActiveSheet();

    // Check immediate value
    const immVal = sheet.getRange(2, 0, 1, 1).getValue();

    // Wait 2.5 seconds
    await new Promise(r => setTimeout(r, 2500));

    const setVal = sheet.getRange(2, 0, 1, 1).getValue();
    const snap = activeWb.save();
    const cellModel = snap.sheets['sheet1']?.cellData?.[2]?.[0];

    return {
      immVal,
      setVal,
      cellModel
    };
  }, testPayload);

  console.log('REAL BROWSER IMPORT RESULT:', JSON.stringify(res, null, 2));
  await browser.close();
})();
