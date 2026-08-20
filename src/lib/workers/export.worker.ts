import ExcelJS from 'exceljs';
import Papa from 'papaparse';

self.onmessage = async (e: MessageEvent) => {
  const { snapshot, format, activeSheetId } = e.data;

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
            rowArr[cIdx] = cell.v !== undefined && cell.v !== null ? String(cell.v) : '';
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
      (snapshot as any) = null;

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

    const sortedSheetIds = snapshot.sheetOrder || Object.keys(snapshot.sheets);
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

      for (let rIdx = 0; rIdx < rowKeys.length; rIdx++) {
        const rowIndex = rowKeys[rIdx];
        const cols = cellDataRaw[rowIndex];
        if (cols) {
          const colKeys = Object.keys(cols).map(Number);
          for (let cIdx = 0; cIdx < colKeys.length; cIdx++) {
            const colIndex = colKeys[cIdx];
            const cellPay = cols[colIndex];
            if (!cellPay) continue;

            const cell = ws.getCell(rowIndex + 1, colIndex + 1);
            if (cellPay.f) {
              let parsedFormula = cellPay.f;
              if (parsedFormula.startsWith('=')) {
                parsedFormula = parsedFormula.substring(1);
              }
              cell.value = {
                formula: parsedFormula,
                result: cellPay.v,
              };
            } else if (cellPay.v !== undefined && cellPay.v !== null) {
              cell.value = cellPay.v;
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
    (snapshot as any) = null;

    self.postMessage({ type: 'PROGRESS', percent: 85, message: 'Mengompresi data ke format .xlsx (Worker)...' });
    await new Promise((resolve) => setTimeout(resolve, 0));

    if (!exportWb) throw new Error('Export workbook is invalid.');
    const arrayBuffer = await exportWb.xlsx.writeBuffer();

    // Release ExcelJS Workbook DOM immediately after writeBuffer
    exportWb = null;

    (self as any).postMessage({ type: 'PROGRESS', percent: 100, message: 'File Excel (.xlsx) selesai dibuat!' });
    (self as any).postMessage({ type: 'COMPLETE', buffer: arrayBuffer }, [arrayBuffer]);
  } catch (err: any) {
    (self as any).postMessage({
      type: 'ERROR',
      error: err?.message || 'Gagal mengekspor file di Background Worker',
    });
  }
};
