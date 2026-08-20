const fs = require('fs');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

function getMemoryUsageMB() {
  const mem = process.memoryUsage();
  return {
    rss: (mem.rss / 1024 / 1024).toFixed(2),
    heapTotal: (mem.heapTotal / 1024 / 1024).toFixed(2),
    heapUsed: (mem.heapUsed / 1024 / 1024).toFixed(2),
  };
}

async function benchmarkMemoryStages() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  console.log('=== MEMORY BENCHMARK FOR REAL FILE report barakat.xlsx ===');
  console.log('1. Memory before import (Node process start):', getMemoryUsageMB());

  // Stage 1: Read File Buffer
  const fileBuf = fs.readFileSync(filePath);
  console.log(`File read to buffer (${(fileBuf.length / 1024 / 1024).toFixed(2)} MB). Memory:`, getMemoryUsageMB());

  // Stage 2: Parse ExcelJS
  let workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(fileBuf);
  console.log('2. Memory after ExcelJS parsing:', getMemoryUsageMB());

  // Stage 3: Convert to workbookData payload
  const sheetOrder = [];
  const sheetsData = {};

  workbook.worksheets.forEach((ws, sIdx) => {
    sheetOrder.push(ws.name);
    const cellData = {};
    let maxRow = 0;
    let maxCol = 0;

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const rowIndex = rowNumber - 1;
      if (rowIndex > maxRow) maxRow = rowIndex;

      const rowCellMap = {};
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const colIndex = colNumber - 1;
        if (colIndex > maxCol) maxCol = colIndex;

        let v = undefined;
        let f = undefined;

        if (cell.type === ExcelJS.ValueType.Formula) {
          f = cell.formula ? (cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`) : undefined;
          if (cell.result !== undefined && cell.result !== null && cell.result !== '') {
            v = typeof cell.result === 'object' ? cell.result.result : cell.result;
          }
        } else if (cell.value !== undefined && cell.value !== null) {
          v = typeof cell.value === 'object' ? cell.value.text || cell.value.result : cell.value;
        }

        const cellInfo = {};
        if (f) cellInfo.f = f;
        if (v !== undefined && v !== null && v !== '') {
          cellInfo.v = typeof v === 'number' ? v : String(v);
          cellInfo.t = typeof v === 'number' ? 2 : 1;
        }
        if (Object.keys(cellInfo).length > 0) rowCellMap[colIndex] = cellInfo;
      });
      if (Object.keys(rowCellMap).length > 0) cellData[rowIndex] = rowCellMap;
    });

    sheetsData[ws.name] = {
      id: ws.name,
      name: ws.name,
      type: 2,
      cellData,
      rowCount: Math.max(maxRow + 50, 100),
      columnCount: Math.max(maxCol + 10, 30)
    };
  });

  console.log('3. Memory after workbookData conversion:', getMemoryUsageMB());

  // Clean ExcelJS workbook from RAM
  workbook = null;
  if (global.gc) global.gc();
  console.log('3b. Memory after garbage collecting ExcelJS workbook:', getMemoryUsageMB());

  // Launch Puppeteer to measure Browser Heap at Stages 4-7
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  const initialMetrics = await page.metrics();
  console.log(`4. Browser JS Heap BEFORE createUniverSheet: ${(initialMetrics.JSHeapUsedSize / 1024 / 1024).toFixed(2)} MB`);

  const workbookData = {
    id: `barakat-bench-${Date.now()}`,
    name: 'report barakat.xlsx',
    sheetOrder,
    sheets: sheetsData
  };

  const createMetrics = await page.evaluate(async (wbData) => {
    const api = window.univerAPI;
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
    oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

    const startMem = window.performance && window.performance.memory ? window.performance.memory.usedJSHeapSize : 0;
    const wb = api.createUniverSheet(wbData);

    const postCreateMem = window.performance && window.performance.memory ? window.performance.memory.usedJSHeapSize : 0;

    await new Promise(r => setTimeout(r, 3000));

    const postCalcMem = window.performance && window.performance.memory ? window.performance.memory.usedJSHeapSize : 0;

    return {
      startMemMB: (startMem / 1024 / 1024).toFixed(2),
      postCreateMemMB: (postCreateMem / 1024 / 1024).toFixed(2),
      postCalcMemMB: (postCalcMem / 1024 / 1024).toFixed(2),
    };
  }, workbookData);

  console.log('5. Browser JS Heap DURING createUniverSheet & Formula Init:', createMetrics);

  const postMetrics = await page.metrics();
  console.log(`6. Browser JS Heap AFTER 3s Formula Engine run: ${(postMetrics.JSHeapUsedSize / 1024 / 1024).toFixed(2)} MB`);

  await browser.close();
}

benchmarkMemoryStages().catch(console.error);
