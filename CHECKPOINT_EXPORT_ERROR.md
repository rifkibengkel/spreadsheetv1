# ROLLBACK CHECKPOINT: KNOWN-GOOD-FORMULA-IMPORT-EXPORT-ERROR

**Date/Time**: 2026-08-31T16:48:00+07:00  
**Base Commit**: `4022836 (feat : custom csv export)`  
**Branch**: `main`  

---

## 1. CURRENT PROJECT STATE SUMMARY

```text
XLSX Import             : ✅ PASS (Working on standard & giant >500MB workbooks)
Formula Editing         : ✅ PASS (Working & responsive)
Formula Calculation     : ✅ PASS (Akurat di Sheet 1 s.d. Sheet 13)
Sheet Switching/Editing : ✅ PASS (Grid aktif & selection box movable)
Browser UI Stability    : ✅ PASS (Bebas dari freeze / Page Unresponsive)
TypeScript Compilation  : ✅ PASS (0 errors, Next.js build compiled)

XLSX Export             : ❌ FAIL (Isolated to Export Worker)
```

---

## 2. CURRENT EXPORT ERROR

```text
File: src/lib/export/excel.ts:145
Console Log: [EXPORT] Worker ERROR: Export worker failed / out of memory / structure mismatch
Stack:
  resolve(blob);
  } else if (type === 'ERROR') {
  > console.error('[EXPORT] Worker ERROR:', error);
    if (worker) worker.terminate();
    reject(new Error(error || 'Export worker failed.'));
```

---

## 3. WORKING TREE MODIFICATIONS AT THIS CHECKPOINT

Only two production files modified relative to commit `4022836`:
1. `src/components/spreadsheet/Spreadsheet.tsx`:
   - In-process native formula calculation (`initialFormulaComputing: 2`).
   - Official `SheetInterceptorService` for visual continuity (`INTERCEPTOR_POINT.CELL_CONTENT`).
   - Zero monkey-patching of Univer internals.
2. `src/lib/workers/excel.worker.ts`:
   - Dual-path importer: Standard ExcelJS for normal workbooks + Targeted Streaming Fallback for giant worksheets (>300MB uncompressed XML).

---

## 4. PROTECTED FILES AUDIT (100% UNTOUCHED — 0 DIFFS)

- `src/lib/import/excel.ts`
- `src/lib/export/excel.ts`
- `src/lib/workers/export.worker.ts`
- `src/lib/export/csv.ts`
- `src/lib/workers/csv.worker.ts`
- `src/lib/session/workbookSession.ts`
- `node_modules/@univerjs/*`

---

## 5. MANDATORY ROLLBACK REQUIREMENT

If any future export experiment introduces regressions to:
- XLSX Import
- Formula Calculation (Sheet 1–13)
- Cell Selection / UI Responsiveness
- Memory stability before export

**IMMEDIATELY ROLL BACK TO THIS EXACT CHECKPOINT.**
