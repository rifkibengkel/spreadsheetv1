const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Verifying Last Column Export Fix ===');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=4096']
  });

  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  console.log('Page loaded. Testing typing into last column and newly extended last column...');

  const result = await page.evaluate(async () => {
    const api = window.univerAPI;
    const wb = api.getActiveWorkbook();
    const sheet = wb.getActiveSheet();

    // Step 1: Type into current last used column (index 2)
    const valCol2 = 'LastCol_Index2_Value';
    sheet.getRange(0, 2, 1, 1).setValue(valCol2);

    // Step 2: Set value in newly extended last column (index 3)
    const valCol3 = 'ExtendedLastCol_Index3_Value';
    sheet.getRange(0, 3, 1, 1).setValue(valCol3);

    await new Promise(r => setTimeout(r, 500));

    // Get snapshot
    const snapshot = wb.save();
    const sheetId = sheet.getSheetId();
    const cellData = snapshot.sheets[sheetId].cellData;

    return {
      cellCol2: cellData[0]?.[2]?.v,
      cellCol3: cellData[0]?.[3]?.v
    };
  });

  console.log('Snapshot Cell Values:', result);

  const success = (
    result.cellCol2 === 'LastCol_Index2_Value' &&
    result.cellCol3 === 'ExtendedLastCol_Index3_Value'
  );

  console.log('Verification Result:', success ? 'PASSED (LAST COLUMN PRESERVED)' : 'FAILED');

  fs.writeFileSync('last-column-verification.json', JSON.stringify({
    result,
    success
  }, null, 2));

  await browser.close();
})();
