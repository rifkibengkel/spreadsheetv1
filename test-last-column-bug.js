const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Forensic Audit: Last Used Column Export Bug ===');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=4096']
  });

  const page = await browser.newPage();
  page.on('console', msg => console.log('[PAGE LOG]', msg.text()));

  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Page loaded. Testing typing into last used column...');

  const result = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook();
    const sheet = wb.getActiveSheet();
    const sheetId = sheet.getSheetId();

    // End active edit
    try {
      if (api.endEdit) api.endEdit();
      else if (api.getCommandService) {
        api.getCommandService().executeCommand('sheet.command.set-activate-cell-edit', { active: false });
      }
    } catch(e){}

    // Get current cellData
    const initialSnapshot = wb.save();
    const initialCellData = initialSnapshot.sheets[sheetId].cellData || {};

    // Find max row and max col in initialCellData
    let maxRow = 0;
    let maxCol = 0;
    Object.keys(initialCellData).forEach(rStr => {
      const r = parseInt(rStr, 10);
      if (r > maxRow) maxRow = r;
      Object.keys(initialCellData[r]).forEach(cStr => {
        const c = parseInt(cStr, 10);
        if (c > maxCol) maxCol = c;
      });
    });

    console.log(`Initial Max Row: ${maxRow}, Initial Max Col: ${maxCol}`);

    // Scenario A: Type value into existing maxCol (Last used column)
    const valLastCol = `LastColVal_${Date.now()}`;
    sheet.getRange(0, maxCol, 1, 1).setValue(valLastCol);

    // Scenario B: Type value into maxCol + 1 (Newly extended last column)
    const newLastCol = maxCol + 1;
    const valExtendedCol = `ExtendedColVal_${Date.now()}`;
    sheet.getRange(0, newLastCol, 1, 1).setValue(valExtendedCol);

    await new Promise(r => setTimeout(r, 500));

    // Get snapshot after edits
    const afterSnapshot = wb.save();
    const afterCellData = afterSnapshot.sheets[sheetId].cellData || {};
    const afterSheetConfig = afterSnapshot.sheets[sheetId];

    // Check if cell at (0, maxCol) and (0, newLastCol) exist in cellData
    const cellInMaxCol = afterCellData[0]?.[maxCol];
    const cellInExtendedCol = afterCellData[0]?.[newLastCol];

    return {
      initialMaxCol: maxCol,
      newLastCol,
      afterSheetColumnCount: afterSheetConfig.columnCount,
      cellInMaxCol,
      cellInExtendedCol
    };
  });

  console.log('Audit Result:', JSON.stringify(result, null, 2));
  fs.writeFileSync('last-col-audit.json', JSON.stringify(result, null, 2));

  await browser.close();
})();
