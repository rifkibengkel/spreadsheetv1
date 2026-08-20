const fs = require('fs');
const puppeteer = require('puppeteer');

async function runFinalAll10UITests() {
  try {
    const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
    console.log('==================================================');
    console.log('FINAL REAL-WORLD VALIDATION (ALL 10 UI TESTS)');
    console.log('FILE:', filePath);
    console.log('==================================================\n');

    const browser = await puppeteer.launch({
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(120000);
    page.setDefaultTimeout(120000);

    await page.setViewport({ width: 1400, height: 900 });

    page.on('console', msg => console.log('BROWSER LOG:', msg.text()));

    console.log('Navigating to http://localhost:3000...');
    await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 120000 });
    await page.waitForFunction('window.univerAPI !== undefined', { timeout: 120000 });
    await new Promise(r => setTimeout(r, 1000));

    console.log('Step 1: Opening Import Modal...');
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const importBtn = btns.find(b => b.textContent.includes('Import'));
      if (importBtn) importBtn.click();
    });
    await new Promise(r => setTimeout(r, 800));

    console.log('Step 2: Uploading real Excel file report barakat.xlsx via File Picker...');
    const fileInput = await page.waitForSelector('input[type="file"]', { timeout: 20000 });
    await fileInput.uploadFile(filePath);
    await new Promise(r => setTimeout(r, 800));

    console.log('Step 3: Triggering Web Worker Import flow...');
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const startBtn = btns.find(b => b.textContent.includes('Mulai Impor') || b.textContent.includes('🚀'));
      if (startBtn) startBtn.click();
    });

    console.log('Waiting for Web Worker import and multi-sheet injection into active workbook (up to 90s)...');
    await page.waitForFunction(() => {
      const api = window.univerAPI;
      if (!api) return false;
      const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : null;
      if (!wb) return false;
      const sheets = wb.getSheets ? wb.getSheets() : [];
      return sheets.length > 1;
    }, { timeout: 90000 });

    console.log('Import successful! Active workbook now contains all imported sheets.');
    console.log('Waiting 5 seconds for initial formula calculation rendering...');
    await new Promise(r => setTimeout(r, 5000));

    await page.screenshot({ path: 'final_real_import_ui.png' });

    console.log('\n--- EXECUTING ALL 10 USER-FACING UI TESTS IN ACTIVE SHEET ENGINE ---');

    const fullReport = await page.evaluate(async () => {
      const api = window.univerAPI;
      const activeWb = api.getActiveWorkbook();
      const results = {};

      const sheets = activeWb.getSheets();
      const getSName = (s) => (s && s.getSheetName ? s.getSheetName() : (s && s.name ? s.name : 'Unknown'));

      const summarySheet = sheets.find(s => getSName(s) === 'Summary') || sheets[0];
      const hadiahSheet = sheets.find(s => getSName(s) === 'Summary Hadiah') || sheets[1];

      const safeGetCellVal = (sheet, r, c) => {
        try {
          if (!sheet) return null;
          const range = sheet.getRange(r, c, 1, 1);
          return range ? range.getValue() : null;
        } catch(e) {
          return String(e.message || e);
        }
      };

      // ----------------------------------------------------
      // TEST 1 — Imported shared formula
      // ----------------------------------------------------
      try {
        const i4 = safeGetCellVal(summarySheet, 3, 8);   // I4
        const i5 = safeGetCellVal(summarySheet, 4, 8);   // I5
        const i6 = safeGetCellVal(summarySheet, 5, 8);   // I6
        const i7 = safeGetCellVal(summarySheet, 6, 8);   // I7
        const i30 = safeGetCellVal(summarySheet, 29, 8); // I30

        results.TEST1_SharedFormulas = {
          I4_Master: i4,
          I5_Follower: i5,
          I6_Follower: i6,
          I7_Follower: i7,
          I30_Follower: i30,
          PASSED: (i5 === 5 && i6 === 105)
        };
      } catch(e) {
        results.TEST1_SharedFormulas = { error: String(e.message || e), PASSED: false };
      }

      // ----------------------------------------------------
      // TEST 2 — Imported cross-sheet formula
      // ----------------------------------------------------
      try {
        const c4Hadiah = safeGetCellVal(hadiahSheet, 3, 2); // C4
        results.TEST2_CrossSheetFormula = {
          SummaryHadiah_C4: c4Hadiah,
          PASSED: (c4Hadiah !== null && c4Hadiah !== undefined)
        };
      } catch(e) {
        results.TEST2_CrossSheetFormula = { error: String(e.message || e), PASSED: false };
      }

      // ----------------------------------------------------
      // TEST 3 — MANUAL FORMULA AFTER IMPORT (=1+1, =10+20)
      // ----------------------------------------------------
      try {
        summarySheet.getRange(0, 25, 1, 1).setValue('=1+1');
        await new Promise(r => setTimeout(r, 600));
        const valOnePlusOne = summarySheet.getRange(0, 25, 1, 1).getValue();

        summarySheet.getRange(0, 25, 1, 1).setValue('=10+20');
        await new Promise(r => setTimeout(r, 600));
        const valTenPlusTwenty = summarySheet.getRange(0, 25, 1, 1).getValue();

        results.TEST3_ManualFormula = {
          onePlusOne: valOnePlusOne,
          tenPlusTwenty: valTenPlusTwenty,
          PASSED: (valOnePlusOne === 2 && valTenPlusTwenty === 30)
        };
      } catch(e) {
        results.TEST3_ManualFormula = { error: String(e.message || e), PASSED: false };
      }

      // ----------------------------------------------------
      // TEST 4 — MANUAL FORMULA USING IMPORTED DATA (=A1+A2)
      // ----------------------------------------------------
      try {
        summarySheet.getRange(0, 0, 1, 1).setValue(100); // A1
        summarySheet.getRange(1, 0, 1, 1).setValue(200); // A2
        summarySheet.getRange(1, 25, 1, 1).setValue('=A1+A2'); // Z2
        await new Promise(r => setTimeout(r, 600));
        const initialSum = summarySheet.getRange(1, 25, 1, 1).getValue();

        summarySheet.getRange(0, 0, 1, 1).setValue(500); // A1
        await new Promise(r => setTimeout(r, 600));
        const updatedSum = summarySheet.getRange(1, 25, 1, 1).getValue();

        results.TEST4_ManualFormulaImportedData = {
          initialSum,
          updatedSum,
          PASSED: (initialSum === 300 && updatedSum === 700)
        };
      } catch(e) {
        results.TEST4_ManualFormulaImportedData = { error: String(e.message || e), PASSED: false };
      }

      // ----------------------------------------------------
      // TEST 5 — MANUAL CROSS-SHEET FORMULA
      // ----------------------------------------------------
      try {
        summarySheet.getRange(2, 25, 1, 1).setValue("='Entries data'!A1"); // Z3
        summarySheet.getRange(3, 25, 1, 1).setValue("='Summary Hadiah'!B4"); // Z4
        await new Promise(r => setTimeout(r, 600));

        const valCross1 = summarySheet.getRange(2, 25, 1, 1).getValue();
        const valCross2 = summarySheet.getRange(3, 25, 1, 1).getValue();

        results.TEST5_ManualCrossSheet = {
          refEntriesDataA1: valCross1,
          refSummaryHadiahB4: valCross2,
          PASSED: (valCross1 !== undefined && valCross2 !== undefined)
        };
      } catch(e) {
        results.TEST5_ManualCrossSheet = { error: String(e.message || e), PASSED: false };
      }

      // ----------------------------------------------------
      // TEST 6 — EDIT AN EXISTING IMPORTED FORMULA
      // ----------------------------------------------------
      try {
        const originalFormula = '=SUMIF($A$4:$A$188,$G5,C$4:C$188)';
        summarySheet.getRange(4, 8, 1, 1).setValue('=1+1'); // I5
        await new Promise(r => setTimeout(r, 600));
        const valTemp = summarySheet.getRange(4, 8, 1, 1).getValue();

        summarySheet.getRange(4, 8, 1, 1).setValue(originalFormula);
        await new Promise(r => setTimeout(r, 600));
        const valRestored = summarySheet.getRange(4, 8, 1, 1).getValue();

        results.TEST6_EditExistingFormula = {
          valTemp,
          valRestored,
          PASSED: (valTemp === 2 && valRestored === 5)
        };
      } catch(e) {
        results.TEST6_EditExistingFormula = { error: String(e.message || e), PASSED: false };
      }

      // ----------------------------------------------------
      // TEST 7 — EDIT DEPENDENCY
      // ----------------------------------------------------
      try {
        const oldG5Val = summarySheet.getRange(4, 6, 1, 1).getValue(); // G5
        const oldI5Val = summarySheet.getRange(4, 8, 1, 1).getValue(); // I5

        summarySheet.getRange(4, 6, 1, 1).setValue('NON_EXISTENT_CODE');
        await new Promise(r => setTimeout(r, 600));
        const newI5Val = summarySheet.getRange(4, 8, 1, 1).getValue();

        summarySheet.getRange(4, 6, 1, 1).setValue(oldG5Val);
        await new Promise(r => setTimeout(r, 600));
        const restoredI5Val = summarySheet.getRange(4, 8, 1, 1).getValue();

        results.TEST7_EditDependency = {
          oldI5Val,
          newI5Val,
          restoredI5Val,
          PASSED: (oldI5Val === 5 && newI5Val === 0 && restoredI5Val === 5)
        };
      } catch(e) {
        results.TEST7_EditDependency = { error: String(e.message || e), PASSED: false };
      }

      // ----------------------------------------------------
      // TEST 8 — FORMULA WITH NO CACHED RESULT
      // ----------------------------------------------------
      try {
        summarySheet.getRange(4, 25, 1, 1).setValue('=SUM(10, 20, 30)'); // Z5
        await new Promise(r => setTimeout(r, 600));
        const uncachedVal = summarySheet.getRange(4, 25, 1, 1).getValue();

        results.TEST8_FormulaNoCachedResult = {
          uncachedVal,
          PASSED: (uncachedVal === 60)
        };
      } catch(e) {
        results.TEST8_FormulaNoCachedResult = { error: String(e.message || e), PASSED: false };
      }

      // ----------------------------------------------------
      // TEST 9 — PAGE RESPONSIVENESS (Switching all 5 sheets)
      // ----------------------------------------------------
      try {
        const sheetNamesLog = [];
        for (const s of sheets) {
          activeWb.setActiveSheet(s);
          await new Promise(r => setTimeout(r, 150));
          sheetNamesLog.push(getSName(s));
        }
        activeWb.setActiveSheet(summarySheet);

        results.TEST9_PageResponsiveness = {
          switchedSheets: sheetNamesLog,
          PASSED: (sheetNamesLog.length === 5)
        };
      } catch(e) {
        results.TEST9_PageResponsiveness = { error: String(e.message || e), PASSED: false };
      }

      // ----------------------------------------------------
      // TEST 10 — EXPORT / REIMPORT
      // ----------------------------------------------------
      try {
        const snapshot = activeWb.save();
        api.disposeUnit(activeWb.getId());
        api.createUniverSheet(snapshot);
        await new Promise(r => setTimeout(r, 1000));

        const reimportedWb = api.getActiveWorkbook ? api.getActiveWorkbook() : api.getAllWorkbooks()[0];
        const reimportedSheets = reimportedWb.getSheets();
        const reimportedSummary = reimportedSheets.find(s => getSName(s) === 'Summary') || reimportedSheets[0];
        const reimportedI5 = reimportedSummary ? reimportedSummary.getRange(4, 8, 1, 1).getValue() : null;

        results.TEST10_ExportReimport = {
          reimportedI5Val: reimportedI5,
          PASSED: (reimportedI5 === 5)
        };
      } catch(e) {
        results.TEST10_ExportReimport = { error: String(e.message || e), PASSED: false };
      }

      return results;
    });

    console.log('\n==================================================');
    console.log('FINAL 10 USER-FACING UI TEST RESULTS:');
    console.log('==================================================');
    console.log(JSON.stringify(fullReport, null, 2));

    fs.writeFileSync('final-10-ui-report.json', JSON.stringify(fullReport, null, 2));
    await browser.close();
  } catch(err) {
    console.error('TOP LEVEL ERROR:', err);
    fs.writeFileSync('fatal-error.log', err.stack || err.message || String(err));
  }
}

runFinalAll10UITests();
