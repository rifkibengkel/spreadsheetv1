import type { IWorkbookData } from '@univerjs/core';
import { ROWS_PER_VIRTUAL_SHEET } from '@/lib/viewport/types';

const CSV_ROW_CHUNK_SIZE = 2500;

function logMemory(tag: string) {
  if (typeof window !== 'undefined' && (window.performance as any)?.memory) {
    const mem = (window.performance as any).memory;
    const usedMB = (mem.usedJSHeapSize / 1024 / 1024).toFixed(2);
    const totalMB = (mem.totalJSHeapSize / 1024 / 1024).toFixed(2);
    console.log(`[MEMORY][CSV][${tag}] Used: ${usedMB} MB / Total: ${totalMB} MB`);
  }
}

export interface ExportCSVOptions {
  scope?: 'sheet' | 'all';
  sheetIndex?: number;
}

function createCsvWorker(): Worker {
  try {
    const workerFunction = () => {
      const textEncoder = new TextEncoder();
      function formatCsvRow(row: string[]): string {
        return row
          .map((cell) => {
            if (cell === '' || cell === undefined || cell === null) return '';
            const str = String(cell);
            if (
              str.includes('"') ||
              str.includes(',') ||
              str.includes('\n') ||
              str.includes('\r')
            ) {
              return `"${str.replace(/"/g, '""')}"`;
            }
            return str;
          })
          .join(',');
      }

      let outputChunks: Uint8Array[] = [];
      let totalRowsExpected = 0;
      let processedRowsCount = 0;
      let sheetName = '';

      self.onmessage = async (e: MessageEvent) => {
        try {
          const data = e.data;
          if (!data) return;

          const { type, chunkIndex, totalChunks, chunkRows, totalRows } = data;

          if (type === 'INIT_CSV') {
            outputChunks = [];
            totalRowsExpected = totalRows || 0;
            processedRowsCount = 0;
            sheetName = data.sheetName || 'Sheet';

            self.postMessage({ type: 'ACK' });
            return;
          }

          if (type === 'APPEND_CSV_CHUNK') {
            if (!chunkRows) {
              self.postMessage({ type: 'CHUNK_ACK', chunkIndex });
              return;
            }

            const chunkRowKeys = Object.keys(chunkRows).map(Number).sort((a, b) => a - b);
            const csvLines: string[] = [];

            for (let i = 0; i < chunkRowKeys.length; i++) {
              const rIndex = chunkRowKeys[i];
              const cols = chunkRows[rIndex];
              if (!cols) continue;

              const colIndices = Object.keys(cols).map(Number).sort((a, b) => a - b);
              if (colIndices.length === 0) continue;
              const maxCol = colIndices[colIndices.length - 1];

              const rowArr: string[] = new Array(maxCol + 1).fill('');
              for (let j = 0; j < colIndices.length; j++) {
                const cIndex = colIndices[j];
                const cell = cols[cIndex];
                if (cell) {
                  let strVal = '';
                  if (cell.v !== undefined && cell.v !== null) {
                    strVal = String(cell.v);
                  } else if (cell.p && cell.p.body && typeof cell.p.body.dataStream === 'string') {
                    strVal = cell.p.body.dataStream.replace(/[\r\n]+$/, '');
                  }
                  rowArr[cIndex] = strVal;
                }
              }

              csvLines.push(formatCsvRow(rowArr));
            }

            if (csvLines.length > 0) {
              const chunkCsvString = csvLines.join('\r\n') + '\r\n';
              const chunkUint8 = textEncoder.encode(chunkCsvString);
              outputChunks.push(chunkUint8);
            }

            processedRowsCount += chunkRowKeys.length;

            self.postMessage({ type: 'CHUNK_ACK', chunkIndex });
            return;
          }

          if (type === 'FINALIZE_CSV') {
            self.postMessage({ type: 'PROGRESS', percent: 95, message: 'Mengemas file CSV...' });

            const totalLength = outputChunks.reduce((sum, c) => sum + c.length, 0);
            const result = new Uint8Array(totalLength);
            let offset = 0;
            for (const chunk of outputChunks) {
              result.set(chunk, offset);
              offset += chunk.length;
            }

            outputChunks = [];

            self.postMessage({ type: 'PROGRESS', percent: 100, message: 'File CSV selesai dibuat!' });
            (self as any).postMessage(
              { type: 'COMPLETE', buffer: result.buffer },
              [result.buffer]
            );
            return;
          }
        } catch (err: any) {
          console.error('[CSV WORKER] Error:', err);
          outputChunks = [];
          self.postMessage({
            type: 'ERROR',
            error: err?.message || 'Gagal mengekspor CSV di Background Worker',
          });
        }
      };
    };

    const code = `(${workerFunction.toString()})()`;
    const blob = new Blob([code], { type: 'application/javascript' });
    const blobUrl = URL.createObjectURL(blob);
    const worker = new Worker(blobUrl);
    setTimeout(() => {
      try {
        URL.revokeObjectURL(blobUrl);
      } catch {}
    }, 2000);
    return worker;
  } catch (err) {
    try {
      return new Worker(new URL('../workers/csv.worker.ts', import.meta.url), { type: 'module' });
    } catch {
      return new Worker('/_next/static/chunks/src_lib_workers_csv_worker_ts.js', { type: 'module' });
    }
  }
}

/**
 * Offloads CSV export to a dedicated streaming background Web Worker.
 * Bounded by O(chunk_size) memory, eliminating main-thread V8 heap spikes
 * and preventing browser tab crashes on large datasets.
 * Supports OPFS virtual datasets, server datasets, and standard Univer sheets.
 */
export async function exportActiveSheetToCSV(
  univerAPI: any,
  onProgress?: (percent: number, message: string) => void,
  options?: ExportCSVOptions
): Promise<Blob> {
  const api = univerAPI || (typeof window !== 'undefined' ? (window as any).univerAPI : null);
  if (!api) throw new Error('Univer API engine is not ready.');

  onProgress?.(2, 'Menyiapkan data sheet aktif untuk CSV...');
  logMemory('START');

  // Step 1: Commit any active cell edit before snapshot (matches Excel export)
  if (
    typeof document !== 'undefined' &&
    document.activeElement &&
    typeof (document.activeElement as HTMLElement).blur === 'function'
  ) {
    (document.activeElement as HTMLElement).blur();
  }
  try {
    const sheet = api.getActiveWorkbook ? api.getActiveWorkbook().getActiveSheet() : null;
    if (sheet && sheet.endEditing) {
      sheet.endEditing();
    }
    const cmd = api?.getCommandService ? api.getCommandService() : api?._commandService;
    if (cmd) {
      try {
        cmd.syncExecuteCommand?.('sheet.operation.set-cell-edit-visible', {
          visible: false,
          eventType: 0,
        });
      } catch {}
      try {
        cmd.syncExecuteCommand?.('sheet.command.set-cell-edit-visible', {
          active: false,
          save: true,
        });
      } catch {}
    }
  } catch (e) {
    console.warn('[CSV EXPORT] endEdit warning:', e);
  }

  // =========================================================================
  // Path A: Server-Backed Virtual Dataset Controller
  // =========================================================================
  if (typeof window !== 'undefined' && (window as any).__datasetViewportController) {
    const dvc = (window as any).__datasetViewportController;
    const serverDatasetId = (dvc as any).datasetId;
    if (serverDatasetId && typeof serverDatasetId === 'string') {
      const scope = options?.scope || 'sheet';
      const sheetIndex =
        options?.sheetIndex !== undefined
          ? options.sheetIndex
          : dvc.getActiveSheetIndex?.() || 0;
      const params = new URLSearchParams();
      params.set('scope', scope);
      params.set('sheetIndex', String(sheetIndex));
      if (dvc.currentQueryId) params.set('queryId', dvc.currentQueryId);

      onProgress?.(25, 'Menyiapkan streaming data CSV dari server...');
      const res = await fetch(`/api/v1/datasets/${serverDatasetId}/export/csv?${params.toString()}`);
      if (!res.ok) {
        const errText = await res.text().catch(() => '');
        throw new Error(`Gagal mengekspor CSV dari server: ${errText || res.statusText}`);
      }
      onProgress?.(80, 'Memproses data CSV...');
      const blob = await res.blob();
      onProgress?.(100, 'File CSV siap diunduh!');
      return blob;
    }
  }

  // =========================================================================
  // Path B: OPFS Virtual / Partitioned Dataset Backing Reader
  // =========================================================================
  const opfsController = typeof window !== 'undefined' ? (window as any).__viewportController : null;
  const opfsReader =
    (typeof window !== 'undefined' ? (window as any).__datasetReader : null) ||
    opfsController?.reader;
  const opfsMeta = opfsController?.getMetadata?.() || opfsReader?.getMetadata?.();

  const isOpfsActive = !!(
    opfsMeta &&
    opfsMeta.status === 'COMMITTED' &&
    typeof opfsMeta.rowCount === 'number' &&
    opfsMeta.rowCount > 0 &&
    opfsReader &&
    typeof opfsReader.readRowRange === 'function'
  );

  if (isOpfsActive) {
    const partitions: any[] = opfsController?.getPartitions?.() || [];
    const activeSheetIndex =
      options?.sheetIndex !== undefined
        ? options.sheetIndex
        : opfsController?.getActiveSheetIndex?.() || 0;

    const activeWorkbook = api?.getActiveWorkbook ? api.getActiveWorkbook() : null;
    const activeSheet = activeWorkbook?.getActiveSheet ? activeWorkbook.getActiveSheet() : null;
    const rawSheet = activeSheet?.getSheet ? activeSheet.getSheet() : activeSheet?._worksheet || activeSheet;
    const liveMatrix = rawSheet?.getCellMatrix?.();

    let startDataRow = 0;
    let endDataRow = opfsMeta.rowCount - 1;
    let exportSheetName = opfsMeta.fileName || 'Sheet';

    if (options?.scope === 'sheet' && partitions.length > 0) {
      const part = partitions[activeSheetIndex] || partitions[0];
      startDataRow = part.startRow;
      endDataRow = part.endRow;
      exportSheetName = part.name || `Sheet ${activeSheetIndex + 1}`;
    } else if (!options?.scope && partitions.length > 1) {
      const part = opfsController?.getActiveSheetPartition?.() || partitions[0];
      startDataRow = part.startRow;
      endDataRow = part.endRow;
      exportSheetName = part.name || `Sheet ${activeSheetIndex + 1}`;
    }

    const totalDataRows = Math.max(0, endDataRow - startDataRow + 1);
    const headers: string[] = opfsMeta.headers || [];
    const hasHeaders = headers.length > 0;

    onProgress?.(5, `Menyiapkan ekspor CSV "${exportSheetName}" (${totalDataRows.toLocaleString()} baris)...`);

    const dataFile: File | null = opfsReader.getDataFile?.() || (opfsReader as any).dataFile;
    const indexArray: BigUint64Array | null = opfsReader.getIndexArray?.() || (opfsReader as any).indexArray;

    function formatCsvCell(val: any): string {
      if (val === null || val === undefined) return '';
      const s = String(val);
      if (s.includes('"') || s.includes(',') || s.includes('\n') || s.includes('\r')) {
        return `"${s.replace(/"/g, '""')}"`;
      }
      return s;
    }

    // 1. Identify all user edits that fall within the exported row range
    // Maps globalDataRow -> Map<colIndex, value>
    const editedRowsMap = new Map<number, Map<number, any>>();

    if (opfsController) {
      const allEdits: Map<string, any> =
        opfsController.getAllUserEdits?.() || (opfsController as any).userEditedCells || new Map();
      for (const edit of allEdits.values()) {
        if (!edit || edit.value === undefined) continue;

        let partition = partitions.find((p: any) => p.sheetId === edit.sheetId);
        if (!partition && partitions.length > 0) {
          const sIdxMatch = String(edit.sheetId).match(/(\d+)/);
          if (sIdxMatch) {
            const idx = parseInt(sIdxMatch[1], 10) - 1;
            partition = partitions[idx];
          }
        }
        if (!partition) {
          partition = partitions[activeSheetIndex] || partitions[0] || { startRow: 0 };
        }

        const localDataRow = edit.row - (hasHeaders ? 1 : 0);
        const globalDataRow = (partition.startRow ?? 0) + localDataRow;

        if (globalDataRow >= startDataRow && globalDataRow <= endDataRow) {
          if (!editedRowsMap.has(globalDataRow)) {
            editedRowsMap.set(globalDataRow, new Map());
          }
          // Formula edit with calculated result: export edit.value
          // If edit.formula exists and edit.value is null/undefined, fallback to live matrix if available
          let valToExport = edit.value;
          if ((valToExport === null || valToExport === undefined) && edit.formula && liveMatrix) {
            const liveVal = liveMatrix.getValue ? liveMatrix.getValue(edit.row, edit.col)?.v : undefined;
            if (liveVal !== undefined && liveVal !== null) {
              valToExport = liveVal;
            }
          }
          // Explicit clear (edit.value === null without formula) is formatted as an empty CSV string
          if (valToExport === null || valToExport === undefined) {
            valToExport = '';
          }
          editedRowsMap.get(globalDataRow)!.set(edit.col, valToExport);
        }
      }
    }

    console.log(
      `[CSV EXPORT][COPY-THROUGH] Exporting "${exportSheetName}" (${totalDataRows.toLocaleString()} rows [${startDataRow}..${endDataRow}], ${editedRowsMap.size} edited rows)`
    );

    // If physical dataFile & indexArray are available, use blazing-fast Copy-Through (Excel Export Reference Architecture)
    if (dataFile && indexArray) {
      // Detect line ending from sample bytes of dataFile
      let eol = '\n';
      try {
        const sampleSlice = dataFile.slice(0, 512);
        const sampleBuf = await sampleSlice.arrayBuffer();
        const sampleStr = new TextDecoder().decode(sampleBuf);
        if (sampleStr.includes('\r\n')) {
          eol = '\r\n';
        } else if (sampleStr.includes('\n')) {
          eol = '\n';
        }
      } catch (_) {}

      const headerStr = hasHeaders ? headers.map(formatCsvCell).join(',') + eol : '';

      // ---------------------------------------------------------------------
      // FAST PATH A: Zero-Copy Copy-Through (Clean Sheet - 0 edits)
      // Reference: Excel export Copy-Through mode (0ms CPU)
      // ---------------------------------------------------------------------
      if (editedRowsMap.size === 0) {
        onProgress?.(50, 'Menyalin data langsung dari disk OPFS (Zero-Copy)...');
        const startByte = Number(indexArray[startDataRow]);
        const endByte = Number(indexArray[endDataRow + 1]);
        const dataSlice = dataFile.slice(startByte, endByte);

        onProgress?.(95, 'Mengemas file CSV...');
        const blob = hasHeaders
          ? new Blob([headerStr, dataSlice], { type: 'text/csv;charset=utf-8;' })
          : new Blob([dataSlice], { type: 'text/csv;charset=utf-8;' });

        onProgress?.(100, 'File CSV selesai diekspor!');
        logMemory('COMPLETE_ZERO_COPY');
        return blob;
      }

      // ---------------------------------------------------------------------
      // FAST PATH B: Spliced Copy-Through (Few edits <= 5,000 rows)
      // Slices unmodified binary spans directly from disk, formats only edited lines!
      // ---------------------------------------------------------------------
      if (editedRowsMap.size <= 5000) {
        onProgress?.(30, `Menggabungkan ${editedRowsMap.size} baris perubahan dengan data asli...`);
        const sortedEditedRows = Array.from(editedRowsMap.keys()).sort((a, b) => a - b);
        const parts: (Blob | string)[] = [];
        if (hasHeaders) {
          parts.push(headerStr);
        }

        let currentByte = Number(indexArray[startDataRow]);

        for (let i = 0; i < sortedEditedRows.length; i++) {
          const editedRowIndex = sortedEditedRows[i];
          const rowStartByte = Number(indexArray[editedRowIndex]);
          const rowEndByte = Number(indexArray[editedRowIndex + 1]);

          // 1. Clean slice of all unmodified rows preceding this edited row
          if (rowStartByte > currentByte) {
            parts.push(dataFile.slice(currentByte, rowStartByte));
          }

          // 2. Read and format only this single edited row
          const rowSlice = await opfsReader.readRow(editedRowIndex);
          const colEdits = editedRowsMap.get(editedRowIndex);
          const formattedCells = [...rowSlice.cells];
          if (colEdits) {
            for (const [colIdx, val] of colEdits.entries()) {
              formattedCells[colIdx] = val;
            }
          }
          const formattedRowStr = formattedCells.map(formatCsvCell).join(',') + eol;
          parts.push(formattedRowStr);

          currentByte = rowEndByte;

          if (i % 50 === 0 || i === sortedEditedRows.length - 1) {
            const pct = Math.round(30 + ((i + 1) / sortedEditedRows.length) * 55);
            onProgress?.(pct, `Mengonversi baris edit (${i + 1}/${sortedEditedRows.length})...`);
          }
        }

        // 3. Clean slice from last edited row to the end of partition/dataset
        const finalEndByte = Number(indexArray[endDataRow + 1]);
        if (finalEndByte > currentByte) {
          parts.push(dataFile.slice(currentByte, finalEndByte));
        }

        onProgress?.(95, 'Mengemas file CSV...');
        const blob = new Blob(parts, { type: 'text/csv;charset=utf-8;' });
        onProgress?.(100, 'File CSV selesai diekspor!');
        logMemory('COMPLETE_SPLICED');
        return blob;
      }
    }

    // -----------------------------------------------------------------------
    // FALLBACK: Streaming Worker for Massive Edits (> 5,000 rows)
    // -----------------------------------------------------------------------
    const MASSIVE_CHUNK_SIZE = 50000;
    const totalChunks = Math.ceil(totalDataRows / MASSIVE_CHUNK_SIZE) || 1;
    const totalExportRows = totalDataRows + (hasHeaders ? 1 : 0);

    return new Promise<Blob>((resolve, reject) => {
      let worker: Worker | null = null;
      let pendingAckResolver: (() => void) | null = null;

      try {
        worker = createCsvWorker();
      } catch (err: any) {
        reject(err);
        return;
      }

      if (!worker) {
        reject(new Error('[CSV EXPORT] Worker could not be created.'));
        return;
      }

      worker.onmessage = (e: MessageEvent) => {
        const { type, percent, message, buffer, error } = e.data;

        if (type === 'ACK' || type === 'CHUNK_ACK') {
          if (pendingAckResolver) {
            const res = pendingAckResolver;
            pendingAckResolver = null;
            res();
          }
        } else if (type === 'PROGRESS') {
          onProgress?.(percent, message);
        } else if (type === 'COMPLETE') {
          logMemory('COMPLETE');
          onProgress?.(100, 'File CSV selesai dibuat!');
          const blob = new Blob([buffer], { type: 'text/csv;charset=utf-8;' });
          if (worker) worker.terminate();
          resolve(blob);
        } else if (type === 'ERROR') {
          console.error('[CSV EXPORT] Worker error:', error);
          if (worker) worker.terminate();
          reject(new Error(error || 'CSV export worker failed.'));
        }
      };

      worker.onerror = (err) => {
        console.error('[CSV EXPORT] Worker runtime error:', err);
        if (worker) worker.terminate();
        reject(new Error(err.message || 'CSV export worker runtime error.'));
      };

      const sendWithAck = (msg: any): Promise<void> => {
        return new Promise<void>((res) => {
          pendingAckResolver = res;
          worker!.postMessage(msg);
        });
      };

      (async () => {
        try {
          await sendWithAck({
            type: 'INIT_CSV',
            sheetId: `sheet-${activeSheetIndex + 1}`,
            sheetName: exportSheetName,
            totalRows: totalExportRows,
          });

          let currentExportRowIndex = 0;
          const targetSheetId = partitions[activeSheetIndex]?.sheetId || 'sheet-1';

          for (let cIdx = 0; cIdx < totalChunks; cIdx++) {
            const chunkStart = startDataRow + cIdx * MASSIVE_CHUNK_SIZE;
            const chunkEnd = Math.min(chunkStart + MASSIVE_CHUNK_SIZE - 1, endDataRow);
            const chunkRows: Record<number, Record<number, any>> = {};

            if (cIdx === 0 && hasHeaders) {
              const headerRowObj: Record<number, any> = {};
              for (let c = 0; c < headers.length; c++) {
                headerRowObj[c] = { v: headers[c] };
              }
              chunkRows[currentExportRowIndex++] = headerRowObj;
            }

            const rowSlices = await opfsReader.readRowRange(chunkStart, chunkEnd);

            for (let i = 0; i < rowSlices.length; i++) {
              const slice = rowSlices[i];
              const rowObj: Record<number, any> = {};

              const partitionIndex = partitions.length > 1
                ? Math.min(partitions.length - 1, Math.floor(slice.rowIndex / ROWS_PER_VIRTUAL_SHEET))
                : activeSheetIndex;
              const currentPartition = partitions[partitionIndex] || partitions[0];
              const currentSheetId = currentPartition?.sheetId || targetSheetId;
              const currentPartitionStartRow = currentPartition?.startRow ?? startDataRow;
              const localDataRow = slice.rowIndex - currentPartitionStartRow;
              const univerDisplayRow = localDataRow + (hasHeaders ? 1 : 0);

              for (let c = 0; c < slice.cells.length; c++) {
                const edit = opfsController?.getCellEdit?.(univerDisplayRow, c, currentSheetId);
                if (edit && edit.value !== undefined) {
                  let valToExport = edit.value;
                  if ((valToExport === null || valToExport === undefined) && edit.formula && liveMatrix) {
                    const liveVal = liveMatrix.getValue ? liveMatrix.getValue(univerDisplayRow, c)?.v : undefined;
                    if (liveVal !== undefined && liveVal !== null) {
                      valToExport = liveVal;
                    }
                  }
                  rowObj[c] = { v: valToExport === null || valToExport === undefined ? '' : valToExport };
                } else {
                  rowObj[c] = { v: slice.cells[c] };
                }
              }
              chunkRows[currentExportRowIndex++] = rowObj;
            }

            const percent = Math.min(85, Math.round(5 + ((cIdx + 1) / totalChunks) * 80));
            onProgress?.(
              percent,
              `Mengonversi CSV "${exportSheetName}" (${cIdx + 1}/${totalChunks})...`
            );

            await sendWithAck({
              type: 'APPEND_CSV_CHUNK',
              chunkIndex: cIdx + 1,
              totalChunks,
              chunkRows,
            });

            await new Promise((r) => setTimeout(r, 0));
          }

          onProgress?.(90, 'Mengemas file CSV...');
          logMemory('BEFORE_FINALIZE');

          worker!.postMessage({ type: 'FINALIZE_CSV' });
        } catch (streamErr: any) {
          console.error('[CSV EXPORT] OPFS Streaming error:', streamErr);
          if (worker) worker.terminate();
          reject(streamErr);
        }
      })();
    });
  }

  // =========================================================================
  // Path C: Standard Univer Worksheet (Excel Export Reference Pattern)
  // =========================================================================
  const activeWorkbook = api.getActiveWorkbook
    ? api.getActiveWorkbook()
    : api.getActiveUniverSheet
      ? api.getActiveUniverSheet()
      : null;

  if (!activeWorkbook) throw new Error('No active workbook to export.');

  // Support both Facade (FWorkbook) and Core (Workbook) matching Excel export pattern
  const coreWorkbook = activeWorkbook.getWorkbook
    ? activeWorkbook.getWorkbook()
    : (activeWorkbook as any)._workbook || activeWorkbook;

  const worksheetMap: Map<string, any> = coreWorkbook.getWorksheets
    ? coreWorkbook.getWorksheets()
    : new Map();

  const activeSheet = activeWorkbook.getActiveSheet ? activeWorkbook.getActiveSheet() : null;
  const activeSheetId = activeSheet
    ? (activeSheet.getSheetId ? activeSheet.getSheetId() : activeSheet.id)
    : (worksheetMap.keys().next().value);

  const fWorksheet = activeWorkbook.getSheetBySheetId
    ? activeWorkbook.getSheetBySheetId(activeSheetId)
    : activeSheet;

  const coreWorksheet = fWorksheet?.getSheet
    ? fWorksheet.getSheet()
    : (fWorksheet as any)?._worksheet ||
      (worksheetMap.has(activeSheetId)
        ? worksheetMap.get(activeSheetId)
        : fWorksheet);

  if (!coreWorksheet && !fWorksheet) {
    throw new Error('Active worksheet not found.');
  }

  const sData = coreWorksheet?.getSnapshot
    ? coreWorksheet.getSnapshot()
    : coreWorksheet?.getConfig
      ? coreWorksheet.getConfig()
      : (fWorksheet as any)?.getSnapshot?.() ||
        (fWorksheet as any)?.getConfig?.();

  const cellDataRaw = sData?.cellData || {};
  const matrix = coreWorksheet?.getCellMatrix
    ? coreWorksheet.getCellMatrix()
    : fWorksheet?.getSheet?.()?.getCellMatrix?.() ||
      (fWorksheet as any)?.getCellMatrix?.();
  const liveMatrix =
    matrix?.getMatrix ? matrix.getMatrix() : (matrix as any)?._matrix || null;

  // Gather complete union of rows from stale snapshot AND live matrix
  const allRowKeysSet = new Set<number>([
    ...Object.keys(cellDataRaw).map(Number),
    ...(liveMatrix ? Object.keys(liveMatrix).map(Number) : []),
  ]);
  const rowKeys = Array.from(allRowKeysSet).sort((a, b) => a - b);
  const totalRows = rowKeys.length;

  const sheetName =
    coreWorksheet?.getName?.() ||
    fWorksheet?.getSheetName?.() ||
    fWorksheet?.getName?.() ||
    sData?.name ||
    activeSheetId;

  const totalChunks = Math.ceil(totalRows / CSV_ROW_CHUNK_SIZE) || 1;

  console.log(`[CSV EXPORT] Exporting sheet "${sheetName}" (${totalRows} rows across ${totalChunks} chunks)`);

  return new Promise<Blob>((resolve, reject) => {
    let worker: Worker | null = null;
    let pendingAckResolver: (() => void) | null = null;

    try {
      worker = createCsvWorker();
    } catch (err: any) {
      reject(err);
      return;
    }

    if (!worker) {
      reject(new Error('[CSV EXPORT] Worker could not be created.'));
      return;
    }

    worker.onmessage = (e: MessageEvent) => {
      const { type, percent, message, buffer, error } = e.data;

      if (type === 'ACK' || type === 'CHUNK_ACK') {
        if (pendingAckResolver) {
          const res = pendingAckResolver;
          pendingAckResolver = null;
          res();
        }
      } else if (type === 'PROGRESS') {
        onProgress?.(percent, message);
      } else if (type === 'COMPLETE') {
        logMemory('COMPLETE');
        onProgress?.(100, 'File CSV selesai dibuat!');
        const blob = new Blob([buffer], { type: 'text/csv;charset=utf-8;' });
        if (worker) worker.terminate();
        resolve(blob);
      } else if (type === 'ERROR') {
        console.error('[CSV EXPORT] Worker error:', error);
        if (worker) worker.terminate();
        reject(new Error(error || 'CSV export worker failed.'));
      }
    };

    worker.onerror = (err) => {
      console.error('[CSV EXPORT] Worker runtime error:', err);
      if (worker) worker.terminate();
      reject(new Error(err.message || 'CSV export worker runtime error.'));
    };

    const sendWithAck = (msg: any): Promise<void> => {
      return new Promise<void>((res) => {
        pendingAckResolver = res;
        worker!.postMessage(msg);
      });
    };

    // Incremental chunk processing with backpressure (matching Excel export)
    (async () => {
      try {
        await sendWithAck({
          type: 'INIT_CSV',
          sheetId: activeSheetId,
          sheetName,
          totalRows,
        });

        for (let cIdx = 0; cIdx < totalChunks; cIdx++) {
          const start = cIdx * CSV_ROW_CHUNK_SIZE;
          const end = Math.min(start + CSV_ROW_CHUNK_SIZE, totalRows);

          // Materialize slice for this chunk only
          const chunkRows: Record<number, Record<number, any>> = {};
          for (let r = start; r < end; r++) {
            const rowIndex = rowKeys[r];
            const rawRow = cellDataRaw[rowIndex] || {};
            const liveRow = liveMatrix ? liveMatrix[rowIndex] : null;

            if (rawRow || liveRow) {
              const allColKeysSet = new Set<number>([
                ...Object.keys(rawRow).map(Number),
                ...(liveRow ? Object.keys(liveRow).map(Number) : []),
              ]);
              const colKeys = Array.from(allColKeysSet).sort((a, b) => a - b);
              const rowObj: Record<number, any> = {};

              for (const colIndex of colKeys) {
                // Live cell from matrix is authoritative
                const liveCell = matrix?.getValue
                  ? matrix.getValue(rowIndex, colIndex)
                  : liveRow?.[colIndex];
                const rawCell = rawRow[colIndex];

                if (liveCell !== undefined && liveCell !== null) {
                  const merged = rawCell ? { ...rawCell, ...liveCell } : { ...liveCell };
                  if (liveCell.v !== undefined) {
                    merged.v = liveCell.v;
                  }
                  if (liveCell.t !== undefined) {
                    merged.t = liveCell.t;
                  }
                  rowObj[colIndex] = merged;
                } else if (!liveMatrix && rawCell !== undefined && rawCell !== null) {
                  rowObj[colIndex] = { ...rawCell };
                }
              }

              if (Object.keys(rowObj).length > 0) {
                chunkRows[rowIndex] = rowObj;
              }
            }
          }

          const percent = Math.min(85, Math.round(5 + ((cIdx + 1) / totalChunks) * 80));
          onProgress?.(
            percent,
            `Mengonversi CSV "${sheetName}" (${cIdx + 1}/${totalChunks})...`
          );

          // Wait for worker ACK before sending next chunk
          await sendWithAck({
            type: 'APPEND_CSV_CHUNK',
            chunkIndex: cIdx + 1,
            totalChunks,
            chunkRows,
          });

          // Yield event loop to allow garbage collection and UI updates
          await new Promise((r) => setTimeout(r, 0));
        }

        console.log('[CSV EXPORT] All chunks sent, finalizing CSV...');
        onProgress?.(90, 'Mengemas file CSV...');
        logMemory('BEFORE_FINALIZE');

        worker!.postMessage({ type: 'FINALIZE_CSV' });
      } catch (streamErr: any) {
        console.error('[CSV EXPORT] Streaming error:', streamErr);
        if (worker) worker.terminate();
        reject(streamErr);
      }
    })();
  });
}

export interface ExportServerCSVOptions {
  datasetId: string;
  scope?: 'sheet' | 'all';
  sheetIndex?: number;
  queryId?: string | null;
  filename?: string;
  saveToDesktop?: boolean;
}

/**
 * Initiates native streaming browser download for server-backed CSV datasets.
 * Bypasses V8 heap limits, Blob allocations, and JSON serialization.
 */
export function triggerServerDatasetCsvDownload(options: ExportServerCSVOptions): void {
  const { datasetId, scope = 'sheet', sheetIndex = 0, queryId, filename } = options;
  const params = new URLSearchParams();
  params.set('scope', scope);
  params.set('sheetIndex', String(sheetIndex));
  if (queryId) params.set('queryId', queryId);
  if (filename) params.set('filename', filename);

  const url = `/api/v1/datasets/${datasetId}/export/csv?${params.toString()}`;

  const a = document.createElement('a');
  a.style.display = 'none';
  a.href = url;
  if (filename) a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    try {
      if (document.body.contains(a)) {
        document.body.removeChild(a);
      }
    } catch {}
  }, 2000);
}

/**
 * Directly writes server-backed CSV dataset to the user's Desktop on filesystem.
 * Handles millions of rows in seconds without transmitting multi-GB data over HTTP.
 */
export async function saveServerDatasetCsvToDesktop(options: ExportServerCSVOptions): Promise<{
  success: boolean;
  filename: string;
  filePath: string;
  rowCount: number;
  message?: string;
}> {
  const { datasetId, scope = 'sheet', sheetIndex = 0, queryId, filename } = options;
  const res = await fetch(`/api/v1/datasets/${datasetId}/export/csv`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      scope,
      sheetIndex,
      queryId,
      filename,
      saveToDesktop: true,
    }),
  });

  const data = await res.json();
  if (!res.ok || data.error) {
    throw new Error(data.error || 'Gagal menyimpan file CSV ke Desktop.');
  }
  return data;
}

if (typeof window !== 'undefined') {
  (window as any).__exportActiveSheetToCSV = exportActiveSheetToCSV;
}
