/**
 * Main-Thread Large Dataset CSV Import Dispatcher
 * Invariant 4 & 5: Dispatches file processing to dedicated csvImport.worker.ts.
 * Main thread receives only progress messages and completion metadata.
 */

import { ImportProgress, ImportResult, OPFSMetadata } from '../storage/opfs/types';
import { CommitManager } from '../storage/opfs/commitManager';

export async function importLargeCSVDataset(
  file: File,
  onProgress?: (progress: ImportProgress) => void
): Promise<ImportResult> {
  if (typeof window === 'undefined' || !window.Worker) {
    throw new Error('Web Workers are not supported in this environment.');
  }

  return new Promise<ImportResult>((resolve, reject) => {
    let worker: Worker | null = null;

    try {
      worker = new Worker(new URL('../workers/csvImport.worker.ts', import.meta.url), {
        type: 'module',
      });
    } catch (workerInitErr) {
      reject(new Error(`Failed to instantiate CSV Import Worker: ${workerInitErr}`));
      return;
    }

    worker.onmessage = (e: MessageEvent) => {
      const data = e.data;
      if (!data) return;

      if (data.type === 'PROGRESS') {
        if (onProgress) onProgress(data.progress);
      } else if (data.type === 'COMPLETE') {
        if (worker) {
          worker.terminate();
          worker = null;
        }
        resolve(data.result as ImportResult);
      } else if (data.type === 'ERROR') {
        if (worker) {
          worker.terminate();
          worker = null;
        }
        reject(new Error(data.error || 'CSV Import failed in worker.'));
      }
    };

    worker.onerror = (err) => {
      if (worker) {
        worker.terminate();
        worker = null;
      }
      reject(err);
    };

    // Send file handle to worker
    worker.postMessage({ type: 'START_IMPORT', file });
  });
}

/**
 * Checks whether an existing committed OPFS dataset is ready to use on startup/reopen.
 */
export async function getCommittedDatasetStatus(): Promise<{
  ready: boolean;
  metadata: OPFSMetadata | null;
  reason?: string;
}> {
  return await CommitManager.checkDatasetStatus();
}
