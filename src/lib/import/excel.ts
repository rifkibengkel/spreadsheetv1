import type { IWorkbookData } from '@univerjs/core';
import { CalculationMode } from '@univerjs/sheets-formula';
import { workbookSession } from '@/lib/session/workbookSession';

/**
 * Worker-based non-blocking async Excel importer.
 * Memisahkan pemrosesan Excel XML ke excel.worker.ts
 */
export async function importExcelToWorkbookDataAsync(
  file: File,
  onProgress?: (percent: number, message?: string) => void
): Promise<{ workbookData: Partial<IWorkbookData> }> {
  onProgress?.(2, 'Menyiapkan Web Worker untuk Impor Excel...');

  if (typeof window !== 'undefined' && window.Worker) {
    try {
      const arrayBuffer = await file.arrayBuffer();

      // Retain a copy for copy-through transit export
      const sessionCopy = new Uint8Array(arrayBuffer.slice(0));
      workbookSession.setOriginalBuffer(sessionCopy, file.name);

      return await new Promise<{ workbookData: Partial<IWorkbookData> }>((resolve, reject) => {
        const worker = new Worker(new URL('../workers/excel.worker.ts', import.meta.url));

        worker.onmessage = (e: MessageEvent) => {
          const { type, percent, message, workbookData, error } = e.data;

          if (type === 'PROGRESS') {
            onProgress?.(percent, message);
          } else if (type === 'COMPLETE') {
            worker.terminate();
            resolve({ workbookData });
          } else if (type === 'ERROR') {
            worker.terminate();
            reject(new Error(error));
          }
        };

        worker.onerror = (err) => {
          worker.terminate();
          reject(err);
        };

        worker.postMessage({ fileArrayBuffer: arrayBuffer, fileName: file.name }, [arrayBuffer]);
      });
    } catch (workerErr) {
      console.warn('Worker gagal, membatalkan import', workerErr);
      throw workerErr;
    }
  }

  throw new Error('Web Worker tidak didukung di browser ini.');
}

export async function importExcelToWorkbookData(file: File): Promise<{ workbookData: Partial<IWorkbookData> }> {
  return importExcelToWorkbookDataAsync(file);
}

/**
 * Legacy compatibility stub. All sheets are already fully populated in the Univer workbook data store.
 */
export async function loadSheetOnDemand(_univerAPI: any, _targetSheetId: string): Promise<boolean> {
  return true;
}

export function replaceUniverWorkbook(univerAPI: any, workbookData: Partial<IWorkbookData>) {
  const api = univerAPI || (typeof window !== 'undefined' ? (window as any).univerAPI : null);
  if (!api) throw new Error('Univer API engine is not ready.');

  // Ensure every sheet has mandatory type: 2 (SheetType.GRID) for FormulaEngine registration
  if (workbookData.sheets) {
    Object.keys(workbookData.sheets).forEach((sheetId) => {
      const sheet = workbookData.sheets![sheetId];
      if (sheet) {
        (sheet as any).type = 2; // SheetType.GRID
        if (!sheet.defaultColumnWidth) sheet.defaultColumnWidth = 88;
        if (!sheet.defaultRowHeight) sheet.defaultRowHeight = 24;
      }
    });
  }

  try {
    const formulaEngine = api.getFormula ? api.getFormula() : null;
    if (formulaEngine && formulaEngine.setInitialFormulaComputing) {
      formulaEngine.setInitialFormulaComputing(CalculationMode.NO_CALCULATION);
    }
  } catch (err) {
    console.warn('Set initial formula computing warning:', err);
  }

  try {
    if (api.getAllWorkbooks) {
      const workbooks = api.getAllWorkbooks();
      if (Array.isArray(workbooks)) {
        workbooks.forEach((wb) => {
          try {
            api.disposeUnit(wb.getId());
          } catch (e) {
            console.warn('Dispose error:', e);
          }
        });
      }
    } else {
      const currentWorkbook = api.getActiveWorkbook ? api.getActiveWorkbook() : null;
      if (currentWorkbook) {
        api.disposeUnit(currentWorkbook.getId());
      }
    }
  } catch (e) {
    console.warn('Dispose unit warning:', e);
  }

  api.createUniverSheet(workbookData);
}

if (typeof window !== 'undefined') {
  (window as any).importExcelToWorkbookDataAsync = importExcelToWorkbookDataAsync;
  (window as any).replaceUniverWorkbook = replaceUniverWorkbook;
}


