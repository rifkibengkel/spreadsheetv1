import type { IWorkbookData } from '@univerjs/core';

const CSV_ROW_CHUNK_SIZE = 2500;

function logMemory(tag: string) {
  if (typeof window !== 'undefined' && (window.performance as any)?.memory) {
    const mem = (window.performance as any).memory;
    const usedMB = (mem.usedJSHeapSize / 1024 / 1024).toFixed(2);
    const totalMB = (mem.totalJSHeapSize / 1024 / 1024).toFixed(2);
    console.log(`[MEMORY][CSV][${tag}] Used: ${usedMB} MB / Total: ${totalMB} MB`);
  }
}

/**
 * Offloads CSV export to a dedicated streaming background Web Worker.
 * Bounded by O(chunk_size) memory, eliminating main-thread V8 heap spikes
 * and preventing browser tab crashes on large datasets.
 */
export async function exportActiveSheetToCSV(
  univerAPI: any,
  onProgress?: (percent: number, message: string) => void
): Promise<Blob> {
  const api = univerAPI || (typeof window !== 'undefined' ? (window as any).univerAPI : null);
  if (!api) throw new Error('Univer API engine is not ready.');

  onProgress?.(2, 'Menyiapkan data sheet aktif untuk CSV...');
  logMemory('START');

  // Step 1: Commit any active cell edit before snapshot
  try {
    const sheet = api.getActiveWorkbook ? api.getActiveWorkbook().getActiveSheet() : null;
    if (sheet && sheet.endEditing) {
      sheet.endEditing();
    } else if (api.getCommandService) {
      api.getCommandService().syncExecuteCommand?.('sheet.command.set-cell-edit-visible', {
        active: false,
        save: true,
      });
    }
  } catch (e) {
    console.warn('[CSV EXPORT] endEdit warning:', e);
  }

  // Step 2: Get active workbook and snapshot
  const activeWorkbook = api.getActiveWorkbook
    ? api.getActiveWorkbook()
    : api.getActiveUniverSheet
      ? api.getActiveUniverSheet()
      : null;

  if (!activeWorkbook) throw new Error('No active workbook to export.');

  const workbookSnapshot: IWorkbookData = activeWorkbook.getSnapshot
    ? activeWorkbook.getSnapshot()
    : activeWorkbook.save();

  const activeSheet = activeWorkbook.getActiveSheet ? activeWorkbook.getActiveSheet() : null;
  const activeSheetId = activeSheet
    ? (activeSheet.getSheetId ? activeSheet.getSheetId() : activeSheet.id)
    : (workbookSnapshot.sheetOrder ? workbookSnapshot.sheetOrder[0] : Object.keys(workbookSnapshot.sheets || {})[0]);

  const sheetData = workbookSnapshot.sheets?.[activeSheetId];
  if (!sheetData || !sheetData.cellData) {
    throw new Error('Sheet data is empty or invalid.');
  }

  const sheetName = sheetData.name || (activeSheet?.getSheetName ? activeSheet.getSheetName() : activeSheetId);
  const cellDataRaw = sheetData.cellData;
  const rowKeys = Object.keys(cellDataRaw).map(Number).sort((a, b) => a - b);
  const totalRows = rowKeys.length;
  const totalChunks = Math.ceil(totalRows / CSV_ROW_CHUNK_SIZE) || 1;

  console.log(`[CSV EXPORT] Exporting sheet "${sheetName}" (${totalRows} rows across ${totalChunks} chunks)`);

  return new Promise<Blob>((resolve, reject) => {
    let worker: Worker | null = null;
    let pendingAckResolver: (() => void) | null = null;

    try {
      worker = new Worker(new URL('../workers/csv.worker.ts', import.meta.url), { type: 'module' });
    } catch (err) {
      try {
        worker = new Worker('/_next/static/chunks/src_lib_workers_csv_worker_ts.js', { type: 'module' });
      } catch (fallbackErr) {
        reject(fallbackErr);
        return;
      }
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
        console.log('[CSV EXPORT] COMPLETE message received from Worker');
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

    // Incremental chunk processing with backpressure
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
            if (cellDataRaw[rowIndex]) {
              chunkRows[rowIndex] = cellDataRaw[rowIndex];
            }
          }

          const percent = Math.min(90, Math.round(5 + ((cIdx + 1) / totalChunks) * 85));
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
        onProgress?.(92, 'Mengemas file CSV...');
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
