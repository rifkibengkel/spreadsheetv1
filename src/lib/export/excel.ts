import type { IWorkbookData } from '@univerjs/core';

/**
 * Offloads Excel (.xlsx) export to a dedicated background Web Worker
 * to eliminate main-thread V8 heap memory spikes and prevent "Aw Snap! Out of Memory" crashes.
 */
export async function exportWorkbookToExcel(
  univerAPI: any,
  onProgress?: (percent: number, message: string) => void
): Promise<Blob> {
  const api = univerAPI || (typeof window !== 'undefined' ? (window as any).univerAPI : null);
  if (typeof window !== 'undefined') {
    (window as any).exportWorkbookToExcel = exportWorkbookToExcel;
  }

  // Commit any active cell edit before snapshot so newly typed values are saved
  try {
    if (api.endEdit) {
      api.endEdit();
    } else if (api.getCommandService) {
      api.getCommandService().executeCommand('sheet.command.set-activate-cell-edit', { active: false, save: true });
    }
  } catch (e) {}

  const activeWorkbook = api.getActiveWorkbook ? api.getActiveWorkbook() : (api.getActiveUniverSheet ? api.getActiveUniverSheet() : null);
  if (!activeWorkbook) throw new Error('No active workbook found in Univer.');

  const snapshot: IWorkbookData = activeWorkbook.save();

  return new Promise<Blob>((resolve, reject) => {
    let worker: Worker | null = null;
    try {
      worker = new Worker(new URL('../workers/export.worker.ts', import.meta.url), { type: 'module' });
    } catch (err) {
      // Fallback path for Next.js Web Worker resolving
      worker = new Worker('/_next/static/chunks/src_lib_workers_export_worker_ts.js', { type: 'module' });
    }

    worker.onmessage = (e: MessageEvent) => {
      const { type, percent, message, buffer, error } = e.data;

      if (type === 'PROGRESS') {
        onProgress?.(percent, message);
      } else if (type === 'COMPLETE') {
        onProgress?.(100, 'File Excel (.xlsx) selesai dibuat!');
        const blob = new Blob([buffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        if (worker) worker.terminate();
        resolve(blob);
      } else if (type === 'ERROR') {
        if (worker) worker.terminate();
        reject(new Error(error || 'Export worker failed.'));
      }
    };

    worker.onerror = (err) => {
      if (worker) worker.terminate();
      reject(new Error(err.message || 'Export worker runtime error.'));
    };

    worker.postMessage({ snapshot, format: 'xlsx' });
  });
}
