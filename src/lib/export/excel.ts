import type { IWorkbookData } from "@univerjs/core";
import { workbookSession } from "@/lib/session/workbookSession";
import { EXPORT_WORKER_CODE } from "@/lib/workers/exportWorkerBundle";

const EXPORT_ROW_CHUNK_SIZE = 2500;

function logMemory(tag: string) {
  if (typeof window !== "undefined" && (window.performance as any)?.memory) {
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
  onProgress?: (percent: number, message: string) => void,
): Promise<Blob> {
  const t0 = Date.now();
  console.log(`[EXPORT][T0] Export clicked (+0ms)`);

  const api =
    univerAPI ||
    (typeof window !== "undefined" ? (window as any).univerAPI : null);
  if (!api) throw new Error("Univer API is not initialized.");

  onProgress?.(2, "Menyiapkan proses ekspor Excel...");
  logMemory("START");

  // Step 1: Force active cell editing commit
  try {
    const sheet = api.getActiveWorkbook
      ? api.getActiveWorkbook().getActiveSheet()
      : null;
    if (sheet && sheet.endEditing) {
      sheet.endEditing();
    } else if (api.getCommandService) {
      api
        .getCommandService()
        .syncExecuteCommand("sheet.command.set-cell-edit-visible", {
          active: false,
          save: true,
        });
    }
  } catch (e) {
    console.warn("[EXPORT] endEdit warning:", e);
  }

  // Step 2: Get active workbook
  const activeWorkbook = api?.getActiveWorkbook
    ? api.getActiveWorkbook()
    : api?.getActiveUniverSheet
      ? api.getActiveUniverSheet()
      : null;

  if (!activeWorkbook) {
    throw new Error("No active workbook found in Univer.");
  }

  // Support both Facade (FWorkbook) and Core (Workbook)
  const coreWorkbook = activeWorkbook.getWorkbook
    ? activeWorkbook.getWorkbook()
    : (activeWorkbook as any)._workbook || activeWorkbook;

  const worksheetMap: Map<string, any> = coreWorkbook.getWorksheets
    ? coreWorkbook.getWorksheets()
    : new Map();

  const facadeSheets: any[] = activeWorkbook.getSheets
    ? activeWorkbook.getSheets()
    : [];

  let sortedSheetIds: string[] = [];
  if (
    coreWorkbook.getSheetOrders &&
    typeof coreWorkbook.getSheetOrders === "function"
  ) {
    sortedSheetIds = Array.from(coreWorkbook.getSheetOrders() as string[]);
  } else if (facadeSheets.length > 0) {
    sortedSheetIds = facadeSheets.map((s: any) => s.getSheetId());
  } else if (worksheetMap.size > 0) {
    sortedSheetIds = Array.from(worksheetMap.keys());
  }

  const sheetNamesMap = new Map<string, string>();
  sortedSheetIds.forEach((id: string) => {
    let name = id;
    if (activeWorkbook.getSheetBySheetId) {
      const fWs = activeWorkbook.getSheetBySheetId(id);
      if (fWs) {
        name = fWs.getSheetName?.() || fWs.getName?.() || fWs.name || id;
      }
    }
    if (name === id && worksheetMap.has(id)) {
      const cWs = worksheetMap.get(id);
      name = cWs?.getName?.() || cWs?.getSheetName?.() || cWs?.name || id;
    }
    sheetNamesMap.set(id, name);
  });

  const availableSheetNames: string[] = sortedSheetIds
    .map((id: string) => sheetNamesMap.get(id) || id)
    .filter(
      (name: any): name is string =>
        typeof name === "string" && name.length > 0,
    );

  // Check Copy-Through capability
  const hasOriginal = workbookSession.hasOriginalWorkbook();
  const dirtySheetIds = workbookSession.getDirtySheetIds();
  const originalBuffer = hasOriginal
    ? workbookSession.getOriginalBuffer()
    : null;

  const dirtySheetIndices: number[] = [];
  sortedSheetIds.forEach((id, idx) => {
    const sName = sheetNamesMap.get(id) || id;
    if (!hasOriginal || dirtySheetIds.has(id) || dirtySheetIds.has(sName)) {
      dirtySheetIndices.push(idx + 1);
    }
  });

  const isCopyThrough = hasOriginal && originalBuffer !== null;
  console.log(
    `[EXPORT][T1] Mode: ${isCopyThrough ? "COPY_THROUGH" : "GREENFIELD"}, Dirty Sheets: ${dirtySheetIndices.length}/${sortedSheetIds.length} (+${Date.now() - t0}ms)`,
  );

  // Step 4: Initialize Web Worker (via Inlined Blob URL to avoid HTTP socket exhaustion)
  return new Promise<Blob>((resolve, reject) => {
    let worker: Worker | null = null;
    let blobUrl: string | null = null;
    let pendingAckResolver: (() => void) | null = null;
    let pendingAckRejecter: ((err: any) => void) | null = null;

    const cleanupWorker = () => {
      if (worker) {
        try {
          worker.terminate();
        } catch (_) {}
        worker = null;
      }
      if (blobUrl) {
        try {
          URL.revokeObjectURL(blobUrl);
        } catch (_) {}
        blobUrl = null;
      }
    };

    try {
      console.log(
        `[EXPORT][T2] WORKER_CREATE: Initializing Web Worker via self-contained Blob URL (+${Date.now() - t0}ms)`,
      );
      const workerBlob = new Blob([EXPORT_WORKER_CODE], {
        type: "application/javascript",
      });
      blobUrl = URL.createObjectURL(workerBlob);
      worker = new Worker(blobUrl);
      console.log(
        `[EXPORT][T3] WORKER_READY: Worker object instantiated successfully via Blob URL (+${Date.now() - t0}ms)`,
      );
    } catch (err) {
      console.error("[EXPORT][T3] Worker synchronous creation failed:", err);
      cleanupWorker();
      reject(err);
      return;
    }

    if (!worker) {
      cleanupWorker();
      reject(new Error("[EXPORT] Worker could not be created."));
      return;
    }

    worker.onmessage = (e: MessageEvent) => {
      const { type, percent, message, buffer, error } = e.data;
      console.log(
        `[EXPORT] Worker onmessage: type=${type}, percent=${percent}, message=${message} (+${Date.now() - t0}ms)`,
      );

      if (type === "ACK" || type === "CHUNK_ACK" || type === "SHEET_ACK") {
        if (pendingAckResolver) {
          const res = pendingAckResolver;
          pendingAckResolver = null;
          pendingAckRejecter = null;
          res();
        }
      } else if (type === "PROGRESS") {
        onProgress?.(percent, message);
      } else if (type === "COMPLETE") {
        console.log(
          `[EXPORT][T7] FINALIZE_COMPLETE: ArrayBuffer returned from worker (+${Date.now() - t0}ms)`,
        );
        logMemory("COMPLETE");
        onProgress?.(90, "File Excel siap diunduh...");
        const blob = new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        console.log(
          `[EXPORT][T8] DOWNLOAD_READY: Blob created (+${Date.now() - t0}ms), size: ${(blob.size / 1024 / 1024).toFixed(2)} MB`,
        );
        cleanupWorker();
        resolve(blob);
      } else if (type === "ERROR") {
        console.error("[EXPORT] Worker ERROR:", error);
        if (pendingAckRejecter) {
          const rej = pendingAckRejecter;
          pendingAckResolver = null;
          pendingAckRejecter = null;
          rej(new Error(error || "Export worker failed."));
        }
        cleanupWorker();
        reject(new Error(error || "Export worker failed."));
      }
    };

    worker.onerror = (err: any) => {
      const errDetails = {
        message: err?.message,
        filename: err?.filename,
        lineno: err?.lineno,
        colno: err?.colno,
        type: err?.type,
        isTrusted: err?.isTrusted,
        error: err?.error ? err.error.message || String(err.error) : undefined,
      };
      console.error(
        "[EXPORT] WORKER RUNTIME ERROR:",
        errDetails.message || "Unknown worker error",
        errDetails,
      );
      if (pendingAckRejecter) {
        const rej = pendingAckRejecter;
        pendingAckResolver = null;
        pendingAckRejecter = null;
        rej(
          new Error(
            errDetails.message ||
              err?.error?.message ||
              "Export worker runtime error.",
          ),
        );
      }
      cleanupWorker();
      reject(
        new Error(
          errDetails.message ||
            err?.error?.message ||
            "Export worker runtime error.",
        ),
      );
    };

    worker.onmessageerror = (err: any) => {
      console.error("[EXPORT] WORKER MESSAGE ERROR:", err);
    };

    const sendWithAck = (
      msg: any,
      transferList?: Transferable[],
    ): Promise<void> => {
      return new Promise<void>((res, rej) => {
        pendingAckResolver = res;
        pendingAckRejecter = rej;
        if (transferList && transferList.length > 0) {
          worker!.postMessage(msg, transferList);
        } else {
          worker!.postMessage(msg);
        }
      });
    };

    // Step 5: Incremental Processing & Copy-Through Orchestration
    (async () => {
      try {
        const transferables: Transferable[] = [];
        let bufferToSend: ArrayBuffer | undefined = undefined;
        if (isCopyThrough && originalBuffer) {
          // Slice a detached copy to transfer in 0ms without structured clone overhead
          const slice = originalBuffer.slice(0);
          bufferToSend = slice.buffer;
          transferables.push(slice.buffer);
        }

        console.log(
          `[EXPORT][T4] INIT_WORKBOOK_SENT: Payload sent to worker (INIT_WORKBOOK, byteLength: ${bufferToSend?.byteLength}, dirtySheetIndices: ${JSON.stringify(dirtySheetIndices)}) (+${Date.now() - t0}ms)`,
        );
        await sendWithAck(
          {
            type: "INIT_WORKBOOK",
            mode: isCopyThrough ? "COPY_THROUGH" : "GREENFIELD",
            originalBuffer: bufferToSend,
            dirtySheetIndices,
            availableSheetNames,
            format: "xlsx",
            sheetsMetadata: sortedSheetIds.map((id, index) => ({
              id,
              index: index + 1,
              name: sheetNamesMap.get(id) || id,
            })),
          },
          transferables,
        );
        console.log(
          `[EXPORT][T5] Worker received payload & ACKed INIT_WORKBOOK (+${Date.now() - t0}ms)`,
        );
        console.log(
          `[EXPORT][T6] Sheet processing started (+${Date.now() - t0}ms)`,
        );

        for (let sIdx = 0; sIdx < sortedSheetIds.length; sIdx++) {
          const sheetId = sortedSheetIds[sIdx];
          const sheetName = sheetNamesMap.get(sheetId) || sheetId;
          const sheetNum = sIdx + 1;
          const isDirty = dirtySheetIndices.includes(sheetNum);

          // In Copy-Through mode, skip clean sheets entirely! Zero-copy directly from original ZIP (0ms CPU)
          if (isCopyThrough && !isDirty) {
            console.log(
              `[EXPORT][T7] Sheet ${sheetNum}/${sortedSheetIds.length} ("${sheetName}") is CLEAN -> Copied directly from original ZIP (0ms CPU) (+${Date.now() - t0}ms)`,
            );
            const globalPercent = Math.min(
              85,
              Math.round(5 + ((sIdx + 1) / sortedSheetIds.length) * 80),
            );
            onProgress?.(
              globalPercent,
              `Menyalin sheet "${sheetName}" (Zero-Copy)...`,
            );
            continue;
          }

          console.log(
            `[EXPORT] Processing dirty sheet ${sheetNum}/${sortedSheetIds.length}: "${sheetName}"`,
          );
          const fWorksheet = activeWorkbook.getSheetBySheetId
            ? activeWorkbook.getSheetBySheetId(sheetId)
            : null;

          const coreWorksheet = fWorksheet?.getSheet
            ? fWorksheet.getSheet()
            : (fWorksheet as any)?._worksheet ||
              (worksheetMap.has(sheetId)
                ? worksheetMap.get(sheetId)
                : fWorksheet);

          if (!coreWorksheet && !fWorksheet) {
            console.warn(
              `[EXPORT] Sheet "${sheetName}" has no coreWorksheet or fWorksheet! Skipping.`,
            );
            continue;
          }

          const tSnapStart = Date.now();
          const sData = coreWorksheet?.getSnapshot
            ? coreWorksheet.getSnapshot()
            : coreWorksheet?.getConfig
              ? coreWorksheet.getConfig()
              : (fWorksheet as any)?.getSnapshot?.() ||
                (fWorksheet as any)?.getConfig?.();

          const cellDataRaw = sData?.cellData || {};
          const rowKeys = Object.keys(cellDataRaw)
            .map(Number)
            .sort((a, b) => a - b);
          const totalRows = rowKeys.length;

          console.log(
            `[EXPORT] Serializing dirty sheet: ${sheetName} (${sheetNum}/${sortedSheetIds.length}, ${totalRows} rows, snapshot in ${Date.now() - tSnapStart}ms)`,
          );
          logMemory(`SHEET_${sheetNum}_START`);

          await sendWithAck({
            type: "INIT_SHEET",
            sheetId,
            sheetIndex: sheetNum,
            sheetName,
            columnData:
              sData?.columnData ||
              coreWorksheet?.getConfig?.()?.columnData ||
              {},
          });

          const totalChunks = Math.ceil(totalRows / EXPORT_ROW_CHUNK_SIZE) || 1;

          for (let cIdx = 0; cIdx < totalChunks; cIdx++) {
            const start = cIdx * EXPORT_ROW_CHUNK_SIZE;
            const end = Math.min(start + EXPORT_ROW_CHUNK_SIZE, totalRows);

            // Materialize slice for this chunk only
            const chunkRows: Record<number, Record<number, any>> = {};
            const matrix = coreWorksheet?.getCellMatrix
              ? coreWorksheet.getCellMatrix()
              : fWorksheet?.getSheet?.()?.getCellMatrix?.() ||
                (fWorksheet as any)?.getCellMatrix?.();

            for (let r = start; r < end; r++) {
              const rowIndex = rowKeys[r];
              if (cellDataRaw[rowIndex]) {
                const rowObj: Record<number, any> = {
                  ...cellDataRaw[rowIndex],
                };
                if (matrix) {
                  const colKeys = Object.keys(rowObj);
                  for (let colIdx = 0; colIdx < colKeys.length; colIdx++) {
                    const colIndex = parseInt(colKeys[colIdx], 10);
                    const cell = rowObj[colIndex];
                    if (cell && (cell.f || cell.si !== undefined)) {
                      const liveCell = matrix.getValue(rowIndex, colIndex);
                      if (
                        liveCell &&
                        liveCell.v !== undefined &&
                        liveCell.v !== null
                      ) {
                        rowObj[colIndex] = {
                          ...cell,
                          v: liveCell.v,
                          t: liveCell.t,
                        };
                      }
                    }
                  }
                }
                chunkRows[rowIndex] = rowObj;
              }
            }

            const chunkPercent = Math.min(
              85,
              Math.round(
                5 +
                  ((sIdx + (cIdx + 1) / totalChunks) / sortedSheetIds.length) *
                    80,
              ),
            );
            onProgress?.(
              chunkPercent,
              `Memproses sheet "${sheetName}" (${cIdx + 1}/${totalChunks})...`,
            );

            // Backpressure wait
            await sendWithAck({
              type: "APPEND_CHUNK",
              sheetId,
              sheetIndex: sheetNum,
              chunkIndex: cIdx + 1,
              totalChunks,
              chunkRows,
            });

            await new Promise((r) => setTimeout(r, 0));
          }

          await sendWithAck({
            type: "END_SHEET",
            sheetId,
            sheetIndex: sheetNum,
            sheetName,
          });
          if (sIdx === 0) {
            console.log(
              `[EXPORT][T5] FIRST_SHEET_COMPLETE: Sheet ${sheetNum}/${sortedSheetIds.length} ("${sheetName}") completed (+${Date.now() - t0}ms)`,
            );
          } else {
            console.log(
              `[EXPORT] Sheet ${sheetNum}/${sortedSheetIds.length} ("${sheetName}") completed (+${Date.now() - t0}ms)`,
            );
          }
        }

        console.log(
          `[EXPORT][T6] FINALIZE_START: XLSX workbook packaging started (FINALIZE_WORKBOOK) (+${Date.now() - t0}ms)`,
        );
        onProgress?.(88, "Mengemas file Excel...");
        logMemory("BEFORE_FINALIZE");

        worker!.postMessage({
          type: "FINALIZE_WORKBOOK",
        });
      } catch (streamErr: any) {
        console.error("[EXPORT] Streaming error:", streamErr);
        cleanupWorker();
        reject(streamErr);
      }
    })();
  });
}

if (typeof window !== "undefined") {
  (window as any).exportWorkbookToExcel = exportWorkbookToExcel;
}
