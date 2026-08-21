import ExcelJS from 'exceljs';
import Papa from 'papaparse';

self.onmessage = async (e: MessageEvent) => {
  let { snapshot, format, activeSheetId } = e.data;

  try {
    if (format === 'csv') {
      self.postMessage({ type: 'PROGRESS', percent: 10, message: 'Menyiapkan data CSV...' });
      const targetSheetId = activeSheetId || (snapshot.sheetOrder ? snapshot.sheetOrder[0] : Object.keys(snapshot.sheets)[0]);
      const sheetData = snapshot.sheets[targetSheetId];

      if (!sheetData || !sheetData.cellData) {
        throw new Error('Sheet data is empty or invalid.');
      }

      const rows: string[][] = [];
      const cellDataRaw = sheetData.cellData;
      const rowKeys = Object.keys(cellDataRaw).map(Number).sort((a, b) => a - b);

      for (let i = 0; i < rowKeys.length; i++) {
        const rIdx = rowKeys[i];
        const cols = cellDataRaw[rIdx];
        if (!cols) continue;

        const maxCol = Math.max(...Object.keys(cols).map(Number));
        const rowArr: string[] = new Array(maxCol + 1).fill('');

        Object.keys(cols).forEach((cStr) => {
          const cIdx = parseInt(cStr, 10);
          const cell = cols[cIdx];
          if (cell) {
            let strVal = '';
            if (cell.v !== undefined && cell.v !== null) {
              strVal = String(cell.v);
            } else if (cell.p && cell.p.body && typeof cell.p.body.dataStream === 'string') {
              strVal = cell.p.body.dataStream.replace(/[\r\n]+$/, '');
            }
            rowArr[cIdx] = strVal;
          }
        });

        rows.push(rowArr);

        if (i % 2000 === 0 || i === rowKeys.length - 1) {
          const percent = Math.min(85, Math.round(10 + (i / rowKeys.length) * 75));
          self.postMessage({ type: 'PROGRESS', percent, message: `Mengonversi baris CSV (${i + 1}/${rowKeys.length})...` });
          await new Promise((resolve) => setTimeout(resolve, 0));
        }
      }

      // Early release snapshot memory
      snapshot = null;

      self.postMessage({ type: 'PROGRESS', percent: 90, message: 'Meng-encode teks CSV...' });
      const csvText = Papa.unparse(rows);
      const encoder = new TextEncoder();
      const uint8 = encoder.encode(csvText);
      const buffer = uint8.buffer;

      (self as any).postMessage({ type: 'PROGRESS', percent: 100, message: 'File CSV selesai dibuat!' });
      (self as any).postMessage({ type: 'COMPLETE', buffer }, [buffer]);
      return;
    }

    // Default: Excel (.xlsx) export
    (self as any).postMessage({ type: 'PROGRESS', percent: 5, message: 'Membuat workbook Excel di Worker...' });

    let exportWb: ExcelJS.Workbook | null = new ExcelJS.Workbook();
    exportWb.creator = 'Univer Spreadsheets App';
    exportWb.lastModifiedBy = 'Univer Spreadsheets App';
    exportWb.calcProperties.fullCalcOnLoad = true;

    const sortedSheetIds = snapshot.sheetOrder || Object.keys(snapshot.sheets);

    // Collect all available worksheet names for precise sheet-name quoting in formulas
    const availableSheetNames: string[] = sortedSheetIds
      .map((id: string) => snapshot.sheets[id]?.name)
      .filter((name: any): name is string => typeof name === 'string' && name.length > 0);

    function sanitizeSheetNames(formula: string): string {
      if (!formula) return formula;
      let result = formula;
      for (let i = 0; i < availableSheetNames.length; i++) {
        const name = availableSheetNames[i];
        if (/[\s&\-\.\/\\]/.test(name)) {
          const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          const regex = new RegExp(`(?<!')${escaped}(?!')!`, 'g');
          result = result.replace(regex, `'${name}'!`);
        }
      }
      return result;
    }

    let totalRowsOverall = 0;

    sortedSheetIds.forEach((sheetId: string) => {
      const sData = snapshot.sheets[sheetId];
      if (sData && sData.cellData) {
        totalRowsOverall += Object.keys(sData.cellData).length;
      }
    });
    if (totalRowsOverall === 0) totalRowsOverall = 1;

    let processedRowsOverall = 0;
    const CHUNK_SIZE = 1000;

    for (let sIdx = 0; sIdx < sortedSheetIds.length; sIdx++) {
      const sheetId = sortedSheetIds[sIdx];
      const sheetData = snapshot.sheets[sheetId];
      if (!sheetData || !exportWb) continue;

      const ws = exportWb.addWorksheet(sheetData.name || sheetId);
      const cellDataRaw = sheetData.cellData || {};
      const rowKeys = Object.keys(cellDataRaw).map(Number).sort((a, b) => a - b);

      // Track shared formulas per sheet (si -> master formula string)
      const sharedFormulaMap = new Map<string, string>();

      // First pass: register all master shared formulas
      for (let rIdx = 0; rIdx < rowKeys.length; rIdx++) {
        const rowIndex = rowKeys[rIdx];
        const cols = cellDataRaw[rowIndex];
        if (cols) {
          const colKeys = Object.keys(cols).map(Number);
          for (let cIdx = 0; cIdx < colKeys.length; cIdx++) {
            const cellPay = cols[colKeys[cIdx]];
            if (cellPay && cellPay.f && cellPay.si !== undefined && cellPay.si !== null) {
              sharedFormulaMap.set(String(cellPay.si), cellPay.f);
            }
          }
        }
      }

      for (let rIdx = 0; rIdx < rowKeys.length; rIdx++) {
        const rowIndex = rowKeys[rIdx];
        const cols = cellDataRaw[rowIndex];
        if (cols) {
          const colKeys = Object.keys(cols).map(Number).sort((a, b) => a - b);
          for (let cIdx = 0; cIdx < colKeys.length; cIdx++) {
            const colIndex = colKeys[cIdx];
            const cellPay = cols[colIndex];
            if (!cellPay) continue;

            const cell = ws.getCell(rowIndex + 1, colIndex + 1);

            let formulaText = cellPay.f;
            if (!formulaText && cellPay.si !== undefined && cellPay.si !== null) {
              formulaText = sharedFormulaMap.get(String(cellPay.si));
            }

            if (formulaText) {
              let parsedFormula = formulaText;
              if (parsedFormula.startsWith('=')) {
                parsedFormula = parsedFormula.substring(1);
              }
              // Sanitize unquoted sheet names containing spaces, &, -, etc.
              parsedFormula = sanitizeSheetNames(parsedFormula);

              const isArrayFormula = cellPay.t === 1 || 
                /SUM\s*\(\s*IF\b|COUNT\s*\(\s*IF\b|AVERAGE\s*\(\s*IF\b|MODE\.MULT\b|MAX\s*\(\s*IF\b|MIN\s*\(\s*IF\b/i.test(parsedFormula) ||
                /:\$[A-Z]+\$\d+\s*=|\$[A-Z]+\$\d+:\$[A-Z]+\$\d+\s*=/i.test(parsedFormula) ||
                (parsedFormula.startsWith('{') && parsedFormula.endsWith('}'));

              if (parsedFormula.startsWith('{') && parsedFormula.endsWith('}')) {
                parsedFormula = parsedFormula.substring(1, parsedFormula.length - 1);
              }

              let formulaRes = cellPay.v !== undefined && cellPay.v !== null 
                ? cellPay.v 
                : (cellPay.p?.body?.dataStream ? cellPay.p.body.dataStream.replace(/[\r\n]+$/, '') : undefined);

              if (typeof formulaRes === 'string' && formulaRes.trim() !== '' && !isNaN(Number(formulaRes))) {
                formulaRes = Number(formulaRes);
              }

              const cellObj: any = {
                formula: parsedFormula,
                result: formulaRes,
              };

              if (isArrayFormula) {
                cellObj.shareType = 'array';
                cellObj.ref = cell.address;
              }

              cell.value = cellObj;
            } else if (cellPay.v !== undefined && cellPay.v !== null) {
              if (typeof cellPay.v === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(cellPay.v)) {
                const dateObj = new Date(cellPay.v);
                if (!isNaN(dateObj.getTime())) {
                  cell.value = dateObj;
                  cell.numFmt = 'yyyy-mm-dd';
                } else {
                  cell.value = cellPay.v;
                }
              } else {
                cell.value = cellPay.v;
              }
            } else if (cellPay.p && cellPay.p.body && typeof cellPay.p.body.dataStream === 'string') {
              const rawText = cellPay.p.body.dataStream.replace(/[\r\n]+$/, '');
              if (rawText.length > 0) {
                cell.value = rawText;
              }
            }
          }
        }

        processedRowsOverall++;
        if (rIdx % CHUNK_SIZE === 0 || rIdx === rowKeys.length - 1) {
          const percent = Math.min(80, Math.round(5 + (processedRowsOverall / totalRowsOverall) * 75));
          self.postMessage({
            type: 'PROGRESS',
            percent,
            message: `Memproses sheet '${sheetData.name || sheetId}' (${rIdx + 1}/${rowKeys.length} baris)...`,
          });
          await new Promise((resolve) => setTimeout(resolve, 0));
        }
      }
    }

    // Release snapshot reference immediately after copying into ExcelJS DOM
    snapshot = null;

    self.postMessage({ type: 'PROGRESS', percent: 85, message: 'Mengompresi data ke format .xlsx (Worker)...' });
    await new Promise((resolve) => setTimeout(resolve, 0));

    if (!exportWb) throw new Error('Export workbook is invalid.');
    const arrayBuffer = await exportWb.xlsx.writeBuffer();

    // Release ExcelJS Workbook DOM immediately after writeBuffer
    exportWb = null;

    const transferableBuffer = arrayBuffer instanceof ArrayBuffer ? arrayBuffer : (arrayBuffer as any)?.buffer;

    (self as any).postMessage({ type: 'PROGRESS', percent: 100, message: 'File Excel (.xlsx) selesai dibuat!' });
    if (transferableBuffer instanceof ArrayBuffer) {
      (self as any).postMessage({ type: 'COMPLETE', buffer: arrayBuffer }, [transferableBuffer]);
    } else {
      (self as any).postMessage({ type: 'COMPLETE', buffer: arrayBuffer });
    }
  } catch (err: any) {
    (self as any).postMessage({
      type: 'ERROR',
      error: err?.message || 'Gagal mengekspor file di Background Worker',
    });
  }
};
