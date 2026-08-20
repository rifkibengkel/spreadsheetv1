import type { IWorkbookData } from '@univerjs/core';

/**
 * Offloads CSV export to a dedicated background Web Worker
 * to eliminate main-thread V8 heap memory spikes and prevent browser crashes on large datasets.
 */
export async function exportActiveSheetToCSV(
  univerAPI: any,
  onProgress?: (percent: number, message: string) => void
): Promise<Blob> {
  const api = univerAPI || (typeof window !== 'undefined' ? (window as any).univerAPI : null);
  if (!api) throw new Error('Univer API engine is not ready.');

  onProgress?.(2, 'Membaca snapshot active sheet...');
  await new Promise((resolve) => setTimeout(resolve, 0));

  const activeWorkbook = api.getActiveWorkbook ? api.getActiveWorkbook() : (api.getActiveUniverSheet ? api.getActiveUniverSheet() : null);
  if (!activeWorkbook) throw new Error('No active workbook to export.');

  const activeSheet = activeWorkbook.getActiveSheet ? activeWorkbook.getActiveSheet() : null;
  const activeSheetId = activeSheet ? activeSheet.getSheetId() : null;
  const snapshot: IWorkbookData = activeWorkbook.save();

  return new Promise<Blob>((resolve, reject) => {
    let worker: Worker | null = null;
    try {
      worker = new Worker(new URL('../workers/export.worker.ts', import.meta.url), { type: 'module' });
    } catch (err) {
      worker = new Worker('/_next/static/chunks/src_lib_workers_export_worker_ts.js', { type: 'module' });
    }

    worker.onmessage = (e: MessageEvent) => {
      const { type, percent, message, buffer, error } = e.data;

      if (type === 'PROGRESS') {
        onProgress?.(percent, message);
      } else if (type === 'COMPLETE') {
        onProgress?.(100, 'File CSV selesai dibuat!');
        const blob = new Blob([buffer], { type: 'text/csv;charset=utf-8;' });
        if (worker) worker.terminate();
        resolve(blob);
      } else if (type === 'ERROR') {
        if (worker) worker.terminate();
        reject(new Error(error || 'CSV export worker failed.'));
      }
    };

    worker.onerror = (err) => {
      if (worker) worker.terminate();
      reject(new Error(err.message || 'Export worker runtime error.'));
    };

    worker.postMessage({ snapshot, format: 'csv', activeSheetId });
  });
}
