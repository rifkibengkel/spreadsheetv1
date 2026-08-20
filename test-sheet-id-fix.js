const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

async function testSheetIdFix() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  console.log('--- Testing Sheet ID Name Matching for Cross-Sheet Formulas ---');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const sheetsDataSlugified = {};
  const sheetsDataExactName = {};
  const sheetOrderSlugified = [];
  const sheetOrderExactName = [];

  workbook.worksheets.forEach((ws, sIdx) => {
    const slugId = ws.name.toLowerCase().replace(/[^a-z0-9]/g, '-') || `sheet-${sIdx + 1}`;
    const exactId = ws.name; // Keep exact sheet name as sheet ID!

    sheetOrderSlugified.push(slugId);
    sheetOrderExactName.push(exactId);

    const cellDataSlugified = {};
    const cellDataExactName = {};

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const rowIndex = rowNumber - 1;
      const rowCellMapSlug = {};
      const rowCellMapExact = {};

      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const colIndex = colNumber - 1;
        let f = undefined;
        let v = undefined;

        let rawFormula = undefined;
        if (cell.type === ExcelJS.ValueType.Formula) {
          rawFormula = cell.formula || (cell.model && cell.model.formula) || (typeof cell.value === 'object' && cell.value !== null ? cell.value.formula : undefined);
        } else if (typeof cell.value === 'object' && cell.value !== null && cell.value.formula) {
          rawFormula = cell.value.formula;
        }

        if (rawFormula && typeof rawFormula === 'string' && rawFormula.trim().length > 0) {
          const trimmed = rawFormula.trim();
          f = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
        }

        if (cell.type === ExcelJS.ValueType.Formula) {
          const res = cell.result;
          if (res !== undefined && res !== null && res !== '') {
            if (typeof res === 'object') {
              if ((res).result !== undefined) v = (res).result;
              else if ((res).error !== undefined) v = (res).error;
            } else {
              v = res;
            }
          }
        } else if (cell.value !== undefined && cell.value !== null) {
          const val = cell.value;
          if (typeof val === 'object') {
            if (val instanceof Date) v = val.toISOString();
            else if (Array.isArray(val.richText)) v = val.richText.map(t => t.text || '').join('');
            else if (val.text !== undefined) v = val.text;
            else if (val.result !== undefined) v = val.result;
            else v = String(val);
          } else {
            v = val;
          }
        }

        const cellInfoSlug = {};
        const cellInfoExact = {};

        if (f !== undefined) {
          cellInfoSlug.f = f;
          cellInfoExact.f = f;
        }

        if (v !== undefined && v !== null && v !== '') {
          if (typeof v === 'number') { cellInfoSlug.v = v; cellInfoSlug.t = 2; cellInfoExact.v = v; cellInfoExact.t = 2; }
          else if (typeof v === 'boolean') { cellInfoSlug.v = v; cellInfoSlug.t = 3; cellInfoExact.v = v; cellInfoExact.t = 3; }
          else { cellInfoSlug.v = String(v); cellInfoSlug.t = 1; cellInfoExact.v = String(v); cellInfoExact.t = 1; }
        }

        if (Object.keys(cellInfoSlug).length > 0) rowCellMapSlug[colIndex] = cellInfoSlug;
        if (Object.keys(cellInfoExact).length > 0) rowCellMapExact[colIndex] = cellInfoExact;
      });

      if (Object.keys(rowCellMapSlug).length > 0) cellDataSlugified[rowIndex] = rowCellMapSlug;
      if (Object.keys(rowCellMapExact).length > 0) cellDataExactName[rowIndex] = rowCellMapExact;
    });

    sheetsDataSlugified[slugId] = {
      id: slugId,
      name: ws.name,
      type: 2,
      status: sIdx === 0 ? 1 : 0,
      rowCount: 100,
      columnCount: 20,
      cellData: cellDataSlugified
    };

    sheetsDataExactName[exactId] = {
      id: exactId,
      name: ws.name,
      type: 2,
      status: sIdx === 0 ? 1 : 0,
      rowCount: 100,
      columnCount: 20,
      cellData: cellDataExactName
    };
  });

  const wbSlugified = { id: `wb-slug-${Date.now()}`, name: 'slug', sheetOrder: sheetOrderSlugified, appVersion: '3.0.0-alpha', sheets: sheetsDataSlugified };
  const wbExact = { id: `wb-exact-${Date.now()}`, name: 'exact', sheetOrder: sheetOrderExactName, appVersion: '3.0.0-alpha', sheets: sheetsDataExactName };

  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction('window.univerAPI !== undefined', { timeout: 30000 });
  await new Promise(r => setTimeout(r, 1000));

  const compareRes = await page.evaluate(async (wSlug, wExact) => {
    const api = window.univerAPI;

    function testWb(wData) {
      const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [api.getActiveWorkbook()];
      oldWbs.forEach(w => { if (w) try { api.disposeUnit(w.getId()); } catch(e){} });
      api.createUniverSheet(wData);
      const activeWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
      return activeWb;
    }

    const wb1 = testWb(wSlug);
    const s1_slug = wb1.getSheetByName('Summary Hadiah');
    const val1 = s1_slug ? s1_slug.getRange(3, 2, 1, 1).getValue() : null;

    const wb2 = testWb(wExact);
    const s2_exact = wb2.getSheetByName('Summary Hadiah');
    const val2 = s2_exact ? s2_exact.getRange(3, 2, 1, 1).getValue() : null;

    return {
      slugifiedSheetIdVal: val1,
      exactSheetIdVal: val2
    };
  }, wbSlugified, wbExact);

  console.log('COMPARISON RESULT:', JSON.stringify(compareRes, null, 2));
  await browser.close();
}

testSheetIdFix().catch(console.error);
