const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Forensic Inspection of Univer Save Snapshot vs Export Worker Payload ===');

  // Create a small test workbook in Univer via page evaluate
  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=4096']
  });

  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Creating test workbook with formulas in Univer...');

  const snapshotReport = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook();
    
    // Get active sheet
    const sheet1 = wb.getActiveSheet();
    sheet1.getRange(0, 0, 1, 1).setValue('Category');
    sheet1.getRange(0, 1, 1, 1).setValue('Value');

    sheet1.getRange(1, 0, 1, 1).setValue('A');
    sheet1.getRange(1, 1, 1, 1).setValue(100);

    sheet1.getRange(2, 0, 1, 1).setValue('B');
    sheet1.getRange(2, 1, 1, 1).setValue(200);

    // Add cross-sheet or formulas
    const sheet2 = wb.create('Summary Sheet');
    sheet2.getRange(0, 0, 1, 1).setValue('Total');
    sheet2.getRange(0, 1, 1, 1).setValue({ f: "=SUM('Sheet1'!B2:B3)" });

    sheet2.getRange(1, 0, 1, 1).setValue('Cross Sheet Ref');
    sheet2.getRange(1, 1, 1, 1).setValue({ f: "='Sheet1'!B2" });

    sheet2.getRange(2, 0, 1, 1).setValue('VLOOKUP');
    sheet2.getRange(2, 1, 1, 1).setValue({ f: "=VLOOKUP(\"A\", 'Sheet1'!A2:B3, 2, FALSE)" });

    // Save snapshot
    const snapshot = wb.save();

    // Inspect how formulas are stored in snapshot
    const formulaCells = [];
    Object.keys(snapshot.sheets).forEach(sId => {
      const s = snapshot.sheets[sId];
      const cellData = s.cellData || {};
      Object.keys(cellData).forEach(r => {
        Object.keys(cellData[r]).forEach(c => {
          const cell = cellData[r][c];
          if (cell && (cell.f || cell.si)) {
            formulaCells.push({
              sheetId: sId,
              sheetName: s.name,
              row: r,
              col: c,
              f: cell.f,
              v: cell.v,
              si: cell.si,
              t: cell.t
            });
          }
        });
      });
    });

    return {
      sheetOrder: snapshot.sheetOrder,
      sheetsMeta: Object.keys(snapshot.sheets).map(id => ({ id, name: snapshot.sheets[id].name })),
      formulaCells
    };
  });

  console.log('Snapshot Formula Inspection Report:');
  console.log(JSON.stringify(snapshotReport, null, 2));

  fs.writeFileSync('snapshot-formula-report.json', JSON.stringify(snapshotReport, null, 2));

  await browser.close();
})();
