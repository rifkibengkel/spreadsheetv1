import type { IWorkbookData } from '@univerjs/core';
import { workbookSession } from '@/lib/session/workbookSession';

const EXPORT_ROW_CHUNK_SIZE = 2500;

function logMemory(tag: string) {
  if (typeof window !== 'undefined' && (window.performance as any)?.memory) {
    const mem = (window.performance as any).memory;
    const usedMB = (mem.usedJSHeapSize / 1024 / 1024).toFixed(2);
    const totalMB = (mem.totalJSHeapSize / 1024 / 1024).toFixed(2);
    console.log(`[MEMORY][${tag}] Used: ${usedMB} MB / Total: ${totalMB} MB`);
  }
}

/**
 * High-performance browser-side Excel XLSX exporter.
 * Uses Copy-Through architecture when original workbook is available (0ms CPU for clean sheets),
 * falling back to streaming Web Worker generation when needed.
 */
export async function exportWorkbookToExcel(
  univerAPI: any,
  onProgress?: (percent: number, message: string) => void
): Promise<Blob> {
  const api = univerAPI || (typeof window !== 'undefined' ? (window as any).univerAPI : null);
  if (!api) throw new Error('Univer API is not initialized.');

  onProgress?.(2, 'Menyiapkan proses ekspor Excel...');
  logMemory('START');

  // Step 1: Force active cell editing commit
  try {
    const sheet = api.getActiveWorkbook ? api.getActiveWorkbook().getActiveSheet() : null;
    if (sheet && sheet.endEditing) {
      sheet.endEditing();
    } else if (api.getCommandService) {
      api
        .getCommandService()
        .syncExecuteCommand(
          'sheet.command.set-cell-edit-visible',
          { active: false, save: true }
        );
    }
  } catch (e) {
    console.warn('[EXPORT] endEdit warning:', e);
  }

  // Step 2: Get active workbook
  const activeWorkbook = api?.getActiveWorkbook
    ? api.getActiveWorkbook()
    : api?.getActiveUniverSheet
      ? api.getActiveUniverSheet()
      : null;

  if (!activeWorkbook) {
    throw new Error('No active workbook found in Univer.');
  }

  // Step 3: Get snapshot reference
  const workbookSnapshot: IWorkbookData = activeWorkbook.getSnapshot
    ? activeWorkbook.getSnapshot()
    : activeWorkbook.save();

  const sortedSheetIds: string[] = activeWorkbook.getSheetOrders
    ? (activeWorkbook.getSheetOrders() as string[])
    : (workbookSnapshot.sheetOrder || Object.keys(workbookSnapshot.sheets || {}));

  const availableSheetNames: string[] = sortedSheetIds
    .map((id: string) => workbookSnapshot.sheets?.[id]?.name)
    .filter((name: any): name is string => typeof name === 'string' && name.length > 0);

  // Check Copy-Through capability
  const hasOriginal = workbookSession.hasOriginalWorkbook();
  const dirtySheetIds = workbookSession.getDirtySheetIds();
  const originalBuffer = hasOriginal ? workbookSession.getOriginalBuffer() : null;

  const dirtySheetIndices: number[] = [];
  sortedSheetIds.forEach((id, idx) => {
    const sName = workbookSnapshot.sheets?.[id]?.name || id;
    if (!hasOriginal || dirtySheetIds.has(id) || dirtySheetIds.has(sName)) {
      dirtySheetIndices.push(idx + 1);
    }
  });

  const isCopyThrough = hasOriginal && originalBuffer !== null;
  console.log(`[EXPORT] Mode: ${isCopyThrough ? 'COPY_THROUGH' : 'GREENFIELD'}, Dirty Sheets: ${dirtySheetIndices.length}/${sortedSheetIds.length}`);

  // Calculate total rows across all sheets for smooth progress reporting
  let globalTotalRows = 0;
  sortedSheetIds.forEach((sId) => {
    const s = workbookSnapshot.sheets?.[sId];
    if (s?.cellData) {
      globalTotalRows += Object.keys(s.cellData).length;
    }
  });
  if (globalTotalRows === 0) globalTotalRows = 1;

  // Step 4: Initialize Web Worker
  return new Promise<Blob>((resolve, reject) => {
    let worker: Worker | null = null;
    let pendingAckResolver: (() => void) | null = null;

    try {
      worker = new Worker(
        new URL('../workers/export.worker.ts', import.meta.url),
        { type: 'module' }
      );
    } catch (err) {
      try {
        worker = new Worker(
          '/_next/static/chunks/src_lib_workers_export_worker_ts.js',
          { type: 'module' }
        );
      } catch (fallbackErr) {
        reject(fallbackErr);
        return;
      }
    }

    if (!worker) {
      reject(new Error('[EXPORT] Worker could not be created.'));
      return;
    }

    worker.onmessage = (e: MessageEvent) => {
      const { type, percent, message, buffer, error } = e.data;

      if (type === 'ACK' || type === 'CHUNK_ACK' || type === 'SHEET_ACK') {
        if (pendingAckResolver) {
          const res = pendingAckResolver;
          pendingAckResolver = null;
          res();
        }
      } else if (type === 'PROGRESS') {
        onProgress?.(percent, message);
      } else if (type === 'COMPLETE') {
        console.log('[EXPORT] COMPLETE message received from Worker');
        logMemory('COMPLETE');
        onProgress?.(90, 'File Excel siap diunduh...');
        const blob = new Blob([buffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        if (worker) worker.terminate();
        resolve(blob);
      } else if (type === 'ERROR') {
        console.error('[EXPORT] Worker ERROR:', error);
        if (worker) worker.terminate();
        reject(new Error(error || 'Export worker failed.'));
      }
    };

    worker.onerror = (err) => {
      console.error('[EXPORT] WORKER RUNTIME ERROR:', err);
      if (worker) worker.terminate();
      reject(new Error(err.message || 'Export worker runtime error.'));
    };

    const sendWithAck = (msg: any): Promise<void> => {
      return new Promise<void>((res) => {
        pendingAckResolver = res;
        worker!.postMessage(msg);
      });
    };

    // Step 5: Incremental Processing & Copy-Through Orchestration
    (async () => {
      try {
        let processedGlobalRows = 0;

        console.log('[EXPORT] Initializing worker workbook');
        await sendWithAck({
          type: 'INIT_WORKBOOK',
          mode: isCopyThrough ? 'COPY_THROUGH' : 'GREENFIELD',
          originalBuffer: isCopyThrough ? originalBuffer : undefined,
          dirtySheetIndices,
          availableSheetNames,
          format: 'xlsx',
          sheetsMetadata: sortedSheetIds.map((id, index) => ({
            id,
            index: index + 1,
            name: workbookSnapshot.sheets?.[id]?.name || id,
          })),
        });

        for (let sIdx = 0; sIdx < sortedSheetIds.length; sIdx++) {
          const sheetId = sortedSheetIds[sIdx];
          const sData = workbookSnapshot.sheets?.[sheetId];
          const sheetName = sData?.name || sheetId;
          const sheetNum = sIdx + 1;
          const isDirty = dirtySheetIndices.includes(sheetNum);

          const cellDataRaw = sData?.cellData || {};
          const rowKeys = Object.keys(cellDataRaw).map(Number).sort((a, b) => a - b);
          const totalRows = rowKeys.length;

          // In Copy-Through mode, skip clean sheets entirely!
          if (isCopyThrough && !isDirty) {
            console.log(`[EXPORT] Sheet ${sheetName} (${sheetNum}) is CLEAN -> Copied directly from original ZIP (0ms CPU)`);
            processedGlobalRows += totalRows;
            const globalPercent = Math.min(85, Math.round(5 + (processedGlobalRows / globalTotalRows) * 80));
            onProgress?.(globalPercent, `Menyalin sheet "${sheetName}" (Zero-Copy)...`);
            continue;
          }

          if (!sData || !sData.cellData) continue;

          console.log(`[EXPORT] Serializing dirty sheet: ${sheetName} (${sheetNum}/${sortedSheetIds.length})`);
          logMemory(`SHEET_${sheetNum}_START`);

          await sendWithAck({
            type: 'INIT_SHEET',
            sheetId,
            sheetIndex: sheetNum,
            sheetName,
            columnData: sData?.columnData || {},
          });

          const totalChunks = Math.ceil(totalRows / EXPORT_ROW_CHUNK_SIZE) || 1;

          for (let cIdx = 0; cIdx < totalChunks; cIdx++) {
            const start = cIdx * EXPORT_ROW_CHUNK_SIZE;
            const end = Math.min(start + EXPORT_ROW_CHUNK_SIZE, totalRows);

            // Materialize slice for this chunk only
            const chunkRows: Record<number, Record<number, any>> = {};
            const worksheet = activeWorkbook.getSheetBySheetId
              ? activeWorkbook.getSheetBySheetId(sheetId)
              : null;
            const matrix = worksheet?.getCellMatrix
              ? worksheet.getCellMatrix()
              : (worksheet as any)?.getSheet?.()?.getCellMatrix?.();

            for (let r = start; r < end; r++) {
              const rowIndex = rowKeys[r];
              if (cellDataRaw[rowIndex]) {
                const rowObj: Record<number, any> = { ...cellDataRaw[rowIndex] };
                if (matrix) {
                  const colKeys = Object.keys(rowObj);
                  for (let cIdx = 0; cIdx < colKeys.length; cIdx++) {
                    const colIndex = parseInt(colKeys[cIdx], 10);
                    const cell = rowObj[colIndex];
                    if (cell && (cell.f || cell.si !== undefined)) {
                      const liveCell = matrix.getValue(rowIndex, colIndex);
                      if (liveCell && liveCell.v !== undefined && liveCell.v !== null) {
                        rowObj[colIndex] = { ...cell, v: liveCell.v, t: liveCell.t };
                      }
                    }
                  }
                }
                chunkRows[rowIndex] = rowObj;
              }
            }

            processedGlobalRows += Object.keys(chunkRows).length;
            const globalPercent = Math.min(85, Math.round(5 + (processedGlobalRows / globalTotalRows) * 80));

            onProgress?.(
              globalPercent,
              `Memproses sheet "${sheetName}" (${cIdx + 1}/${totalChunks})...`
            );

            // Backpressure wait
            await sendWithAck({
              type: 'APPEND_CHUNK',
              sheetId,
              sheetIndex: sheetNum,
              chunkIndex: cIdx + 1,
              totalChunks,
              chunkRows,
            });

            await new Promise((r) => setTimeout(r, 0));
          }

          console.log(`[EXPORT] sheet complete: ${sheetName}`);
          await sendWithAck({
            type: 'END_SHEET',
            sheetId,
            sheetIndex: sheetNum,
          });
        }

        console.log('[EXPORT] Finalizing XLSX package...');
        onProgress?.(88, 'Mengemas file Excel...');
        logMemory('BEFORE_FINALIZE');

        worker!.postMessage({
          type: 'FINALIZE_WORKBOOK',
        });
      } catch (streamErr: any) {
        console.error('[EXPORT] Streaming error:', streamErr);
        if (worker) worker.terminate();
        reject(streamErr);
      }
    })();
  });
}

if (typeof window !== 'undefined') {
  (window as any).exportWorkbookToExcel = exportWorkbookToExcel;
}