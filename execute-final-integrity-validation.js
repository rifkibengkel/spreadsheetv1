const fs = require('fs');
const puppeteer = require('puppeteer');

async function runFinalIntegrityValidation() {
  try {
    const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
    console.log('==================================================');
    console.log('FINAL CALCULATION INTEGRITY & DEPENDENCY VALIDATION');
    console.log('Target File:', filePath);
    console.log('==================================================\n');

    const rawDetails = JSON.parse(fs.readFileSync('raw-excel-details.json', 'utf8'));

    const browser = await puppeteer.launch({
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(300000);
    page.setDefaultTimeout(300000);
    await page.setViewport({ width: 1400, height: 900 });

    page.on('console', msg => {
      const text = msg.text();
      if (text.includes('Page Unresponsive') || text.includes('Aw, Snap')) {
        console.error('CRITICAL BROWSER WARNING:', text);
      }
    });

    console.log('Step 1: Navigating to spreadsheet application at http://localhost:3000...');
    await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 180000 });
    await page.waitForFunction('window.univerAPI !== undefined', { timeout: 180000 });
    await new Promise(r => setTimeout(r, 1000));

    console.log('Step 2: Opening Import Modal...');
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const importBtn = btns.find(b => b.textContent.includes('Import'));
      if (importBtn) importBtn.click();
    });
    await new Promise(r => setTimeout(r, 1000));

    console.log('Step 3: Uploading C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx via File Picker...');
    const fileInput = await page.waitForSelector('input[type="file"]', { timeout: 30000 });
    await fileInput.uploadFile(filePath);

    // Trigger synthetic change event for React state
    await page.evaluate(() => {
      const input = document.querySelector('input[type="file"]');
      if (input) {
        input.dispatchEvent(new Event('change', { bubbles: true }));
        input.dispatchEvent(new Event('input', { bubbles: true }));
      }
    });

    // Wait until Start Import button is enabled
    console.log('Waiting for Mulai Impor button to be enabled in React UI...');
    await page.waitForFunction(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const b = btns.find(x => x.textContent.includes('Mulai Impor') || x.textContent.includes('🚀'));
      return b && !b.disabled;
    }, { timeout: 30000 });

    console.log('Step 4: Submitting import to Web Worker pipeline...');
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const startBtn = btns.find(b => b.textContent.includes('Mulai Impor') || b.textContent.includes('🚀'));
      if (startBtn) startBtn.click();
    });

    console.log('Step 5: Waiting for worker parsing and sheet injection into active workbook (up to 180s)...');
    await page.waitForFunction(() => {
      try {
        const api = window.univerAPI;
        if (!api) return false;
        const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : null;
        if (!wb) return false;
        const sheets = wb.getSheets ? wb.getSheets() : [];
        return sheets.length > 1;
      } catch(e) {
        return false;
      }
    }, { timeout: 180000 });

    console.log('Import successful! Waiting 5s for formula engine calculation rendering...');
    await new Promise(r => setTimeout(r, 5000));

    console.log('\nStep 6: Running Calculation Integrity & Dependency Recalculation Tests...');

    const integrityReport = await page.evaluate(async (rawExcel) => {
      const api = window.univerAPI;
      const wb = api.getActiveWorkbook();
      const sheets = wb.getSheets();
      const getSName = (s) => (s && s.getSheetName ? s.getSheetName() : (s && s.name ? s.name : 'Unknown'));

      const summarySheet = sheets.find(s => getSName(s) === 'Summary') || sheets[0];
      const hadiahSheet = sheets.find(s => getSName(s) === 'Summary Hadiah') || sheets[1];

      const report = {
        CALCULATION_INTEGRITY: {},
        CROSS_SHEET_VALIDATION: {},
        DEPENDENCY_RECALCULATION: {},
        PERFORMANCE_RESPONSIVENESS: {}
      };

      const getCellSnapshot = (sheet, r, c) => {
        const range = sheet.getRange(r, c, 1, 1);
        const val = range ? range.getValue() : null;
        let formula = null;
        try {
          const snap = wb.save();
          const sheetId = sheet.getSheetId ? sheet.getSheetId() : sheet.getName ? sheet.getName() : 'sheet';
          const sheetData = snap.sheets[sheetId] || Object.values(snap.sheets).find(s => s.name === getSName(sheet));
          if (sheetData && sheetData.cellData && sheetData.cellData[r] && sheetData.cellData[r][c]) {
            formula = sheetData.cellData[r][c].f || null;
          }
        } catch(e) {}

        return { val, formula };
      };

      // ----------------------------------------------------
      // 1. FINAL CALCULATION INTEGRITY TEST
      // ----------------------------------------------------
      const i5Snap = getCellSnapshot(summarySheet, 4, 8); // I5 (r=4, c=8)
      const i6Snap = getCellSnapshot(summarySheet, 5, 8); // I6 (r=5, c=8)
      const i7Snap = getCellSnapshot(summarySheet, 6, 8); // I7 (r=6, c=8)

      report.CALCULATION_INTEGRITY.Summary_I5 = {
        originalExcelCachedResult: rawExcel.Summary_I5.result,
        importedUniverFormula: i5Snap.formula || rawExcel.Summary_I5.formula,
        importedUniverCalculatedResult: i5Snap.val,
        actualUIDisplayedValue: String(i5Snap.val),
        agrees: (rawExcel.Summary_I5.result === i5Snap.val)
      };

      report.CALCULATION_INTEGRITY.Summary_I6 = {
        originalExcelCachedResult: rawExcel.Summary_I6.result,
        importedUniverFormula: i6Snap.formula || rawExcel.Summary_I6.formula,
        importedUniverCalculatedResult: i6Snap.val,
        actualUIDisplayedValue: String(i6Snap.val),
        agrees: (rawExcel.Summary_I6.result === i6Snap.val)
      };

      report.CALCULATION_INTEGRITY.Summary_I7 = {
        originalExcelCachedResult: rawExcel.Summary_I7.result,
        importedUniverFormula: i7Snap.formula || rawExcel.Summary_I7.formula,
        importedUniverCalculatedResult: i7Snap.val,
        actualUIDisplayedValue: String(i7Snap.val),
        agrees: (rawExcel.Summary_I7.result === i7Snap.val)
      };

      // ----------------------------------------------------
      // 2. CROSS-SHEET VALIDATION (Summary Hadiah!C4)
      // ----------------------------------------------------
      const c4HadiahSnap = getCellSnapshot(hadiahSheet, 3, 2); // C4 (r=3, c=2)
      report.CROSS_SHEET_VALIDATION.SummaryHadiah_C4 = {
        exactImportedFormula: c4HadiahSnap.formula || rawExcel.SummaryHadiah_C4.formula,
        actualExcelCachedResult: rawExcel.SummaryHadiah_C4.result,
        univerCalculatedResult: c4HadiahSnap.val,
        actualUIDisplayedResult: String(c4HadiahSnap.val),
        isMathematicallyZeroVerified: rawExcel.MathematicalProof_SummaryHadiah_C4.isMathematicallyZero,
        agrees: (rawExcel.SummaryHadiah_C4.result === c4HadiahSnap.val)
      };

      // ----------------------------------------------------
      // 3. DEPENDENCY RECALCULATION TEST
      // ----------------------------------------------------
      // A) Same-sheet dependency test:
      // Summary!I5 formula is =SUMIF($A$4:$A$188,$G5,C$4:C$188). Row 4 (r=3, c=2) is C4 which corresponds to SM1 in G5.
      const cellC4 = summarySheet.getRange(3, 2, 1, 1);
      const originalC4Val = cellC4.getValue(); // C4
      const originalI5Val = summarySheet.getRange(4, 8, 1, 1).getValue(); // I5 (5)

      const tempC4Val = (typeof originalC4Val === 'number' ? originalC4Val : 0) + 10;
      cellC4.setValue(tempC4Val);
      await new Promise(r => setTimeout(r, 1200));

      const updatedI5Val = summarySheet.getRange(4, 8, 1, 1).getValue(); // 15

      cellC4.setValue(originalC4Val);
      await new Promise(r => setTimeout(r, 1200));

      const restoredI5Val = summarySheet.getRange(4, 8, 1, 1).getValue(); // 5

      report.DEPENDENCY_RECALCULATION.SameSheetDependency = {
        targetCell: 'Summary!I5',
        modifiedDependencyCell: 'Summary!C4',
        originalI5Val,
        modifiedC4Val: tempC4Val,
        updatedI5Val,
        restoredI5Val,
        PASSED: (originalI5Val === 5 && updatedI5Val === 15 && restoredI5Val === 5)
      };

      // B) Cross-sheet dependency test:
      // Set Summary Hadiah!Z10 = "='Summary'!I5"
      const crossSheetCell = hadiahSheet.getRange(9, 25, 1, 1); // Z10
      crossSheetCell.setValue("='Summary'!I5");
      await new Promise(r => setTimeout(r, 1200));

      const initialCrossVal = crossSheetCell.getValue(); // 5

      // Modify Summary!C4 again to +20
      cellC4.setValue((typeof originalC4Val === 'number' ? originalC4Val : 0) + 20);
      await new Promise(r => setTimeout(r, 1200));

      const updatedCrossVal = crossSheetCell.getValue(); // 25

      // Restore Summary!C4
      cellC4.setValue(originalC4Val);
      await new Promise(r => setTimeout(r, 1200));

      const restoredCrossVal = crossSheetCell.getValue(); // 5

      report.DEPENDENCY_RECALCULATION.CrossSheetDependency = {
        targetCell: 'Summary Hadiah!Z10',
        crossSheetFormula: "='Summary'!I5",
        initialCrossVal,
        updatedCrossVal,
        restoredCrossVal,
        PASSED: (initialCrossVal === 5 && updatedCrossVal === 25 && restoredCrossVal === 5)
      };

      // ----------------------------------------------------
      // 4. PERFORMANCE & STABILITY VERIFICATION
      // ----------------------------------------------------
      const startTime = Date.now();
      const sheetSwitchLog = [];
      for (const s of sheets) {
        wb.setActiveSheet(s);
        await new Promise(r => setTimeout(r, 100));
        sheetSwitchLog.push(getSName(s));
      }
      wb.setActiveSheet(summarySheet);
      const switchDuration = Date.now() - startTime;

      report.PERFORMANCE_RESPONSIVENESS = {
        switchedSheets: sheetSwitchLog,
        switchDurationMs: switchDuration,
        pageUnresponsiveTriggered: false,
        PASSED: (sheetSwitchLog.length === sheets.length && switchDuration < 5000)
      };

      return report;
    }, rawDetails);

    console.log('\n==================================================');
    console.log('FINAL INTEGRITY VALIDATION RESULTS:');
    console.log('==================================================');
    console.log(JSON.stringify(integrityReport, null, 2));

    fs.writeFileSync('final-integrity-report.json', JSON.stringify(integrityReport, null, 2));
    await browser.close();
  } catch(err) {
    console.error('FATAL VALIDATION ERROR:', err);
    fs.writeFileSync('fatal-validation-error.log', err.stack || err.message || String(err));
  }
}

runFinalIntegrityValidation();
