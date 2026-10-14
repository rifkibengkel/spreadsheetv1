/**
 * Dataset Viewport Controller (Phase 1B)
 * Orchestrates dynamic bounded viewport window hydration for partitioned 11.8M-row datasets.
 * 
 * Pipeline:
 * Univer Scroll/Jump -> Generation Token Check -> Dedicated Worker -> Range API -> 
 * Bounded LRU Cache -> Viewport Window Calculation -> Univer Viewport Adapter Hydration.
 */

import { UniverViewportAdapter } from './univerViewportAdapter';
import { CellEditRecord, VirtualSheetPartition, ROWS_PER_VIRTUAL_SHEET } from './types';
import { RowSlice } from '../storage/opfs/types';
import { DatasetRow } from '../largeCsv/reader/datasetRangeReader';

export interface DatasetViewportConfig {
  /** Number of visible rows in typical screen viewport (default: 50) */
  visibleRows: number;
  /** Extra rows hydrated before and after visible rows (default: 100) */
  overscanRows: number;
  /** Scroll margin threshold to trigger sliding window prefetch (default: 30) */
  marginThreshold: number;
  /** Base URL for API requests (default: window.location.origin or '') */
  apiBaseUrl?: string;
}

export const DEFAULT_DATASET_VIEWPORT_CONFIG: DatasetViewportConfig = {
  visibleRows: 50,
  overscanRows: 100,
  marginThreshold: 30,
};

export interface ViewportActiveWindow {
  startRow: number;
  endRow: number;
  rowCount: number;
  generationToken: number;
}

export interface ViewportTimingMetrics {
  t0RequestReceived?: number;
  t1PartitionResolutionMs?: number;
  t2IndexLookupMs?: number;
  t3FileReadMs?: number;
  t4DecodeMs?: number;
  t5ServerTotalMs?: number;
  t6WorkerReceiveMs?: number;
  t7WorkerDurationMs?: number;
  t8UniverHydrationMs?: number;
  e2eDurationMs?: number;
}

export class DatasetViewportController {
  private datasetId: string = '';
  private adapter: UniverViewportAdapter;
  private config: DatasetViewportConfig;
  private worker: Worker | null = null;
  private generationToken: number = 0;
  private activeWindow: ViewportActiveWindow | null = null;
  private totalDataRows: number = 0;
  private columns: string[] = [];
  private columnCount: number = 0;
  private baseDatasetTotalRows: number = 0;
  private currentQueryId: string | null = null;
  private userEditedCells: Map<string, CellEditRecord> = new Map();
  private isLoading: boolean = false;
  private isMounted: boolean = false;
  private lastMetrics: ViewportTimingMetrics | null = null;
  private onMetricsUpdated?: (metrics: ViewportTimingMetrics) => void;

  /** Virtual sheet partitions (4,000,000 data rows per sheet) */
  private partitions: VirtualSheetPartition[] = [];
  private activeSheetIndex: number = 0;

  constructor(
    adapter: UniverViewportAdapter,
    config: Partial<DatasetViewportConfig> = {}
  ) {
    this.adapter = adapter;
    this.config = { ...DEFAULT_DATASET_VIEWPORT_CONFIG, ...config };
    this.initWorker();
  }

  private calculatePartitions(): void {
    const sheetCount = Math.max(1, Math.ceil(this.totalDataRows / ROWS_PER_VIRTUAL_SHEET));
    this.partitions = [];
    for (let s = 0; s < sheetCount; s++) {
      const startRow = s * ROWS_PER_VIRTUAL_SHEET;
      const endRow = Math.min(this.totalDataRows - 1, (s + 1) * ROWS_PER_VIRTUAL_SHEET - 1);
      const rowCount = Math.max(0, endRow - startRow + 1);
      this.partitions.push({
        sheetId: `sheet-${s + 1}`,
        sheetIndex: s,
        startRow,
        endRow,
        rowCount,
      });
    }
  }

  public getPartitions(): VirtualSheetPartition[] {
    return this.partitions;
  }

  public getActiveSheetPartition(): VirtualSheetPartition | null {
    return this.partitions[this.activeSheetIndex] || null;
  }

  public getActiveSheetIndex(): number {
    return this.activeSheetIndex;
  }

  public recordUserEdit(sheetId: string, row: number, col: number, value: any, formula?: string): void {
    const key = `${sheetId}_${row}_${col}`;
    this.userEditedCells.set(key, {
      sheetId,
      row,
      col,
      value,
      formula,
      timestamp: Date.now(),
    });
  }

  public getAllUserEdits(): Map<string, CellEditRecord> {
    return new Map(this.userEditedCells);
  }

  public clearUserEdits(): void {
    this.userEditedCells.clear();
  }


  private initWorker(): void {
    if (typeof window !== 'undefined' && window.Worker) {
      try {
        this.worker = new Worker(
          new URL('../workers/datasetViewport.worker.ts', import.meta.url),
          { type: 'module' }
        );
        this.worker.onmessage = this.handleWorkerMessage.bind(this);
        this.worker.onerror = (err) => {
          console.error('[DatasetViewportController] Worker error:', err);
        };
      } catch (err) {
        console.warn('[DatasetViewportController] Failed to initialize Dedicated Worker:', err);
      }
    }
  }

  public setAdapter(adapter: UniverViewportAdapter): void {
    this.adapter = adapter;
  }

  public setMetricsListener(listener: (m: ViewportTimingMetrics) => void): void {
    this.onMetricsUpdated = listener;
  }

  public getActiveWindow(): ViewportActiveWindow | null {
    return this.activeWindow;
  }

  public getLastMetrics(): ViewportTimingMetrics | null {
    return this.lastMetrics;
  }

  public getTotalRows(): number {
    return this.totalDataRows;
  }

  public getDatasetId(): string {
    return this.datasetId;
  }

  public isReady(): boolean {
    return this.isMounted && this.adapter.isReady();
  }

  public getCurrentQueryId(): string | null {
    return this.currentQueryId;
  }

  public getHeaders(): string[] {
    return this.columns;
  }

  public getBaseTotalRows(): number {
    return this.baseDatasetTotalRows || this.totalDataRows;
  }

  /**
   * Mounts the server-partitioned dataset to Univer.
   * Sets logical sheet dimensions (e.g. 11,809,234 rows x 54 columns)
   * and hydrates the initial bounded viewport window.
   */
  public async mount(datasetId: string): Promise<boolean> {
    this.datasetId = datasetId;
    const token = ++this.generationToken;
    this.isLoading = true;

    const tStart = performance.now();

    try {
      // 1. Fetch initial window (rows 0..initialCount) directly or via worker
      const initialCount = this.config.visibleRows + this.config.overscanRows * 2;
      const initialData = await this.fetchRange(0, initialCount, token);

      if (token !== this.generationToken || !initialData) {
        return false;
      }

      this.totalDataRows = initialData.totalRows;
      this.baseDatasetTotalRows = initialData.totalRows;
      this.columns = initialData.columns;
      this.columnCount = initialData.columnCount;

      // 2. Mount virtual sheet dimensions into Univer
      this.calculatePartitions();
      this.activeSheetIndex = 0;

      if (this.adapter.mountVirtualSheets) {
        await this.adapter.mountVirtualSheets(this.partitions, this.columnCount, this.columns);
      } else {
        await this.adapter.mountVirtualSheet({
          magic: 'SHEET_OPFS_V1' as any,
          version: 1,
          fileName: `dataset-${datasetId}.csv`,
          rowCount: this.totalDataRows,
          columnCount: this.columnCount,
          headers: this.columns,
          dataSizeBytes: 0,
          indexSizeBytes: 0,
          status: 'COMMITTED',
          createdAt: new Date().toISOString(),
        });
      }

      // 3. Hydrate initial window into primary virtual sheet
      const primaryPartition = this.partitions[0];
      const targetSubId = primaryPartition?.sheetId || this.adapter.getActiveSubUnitId();

      const tHydrate0 = performance.now();
      const rowSlices: RowSlice[] = initialData.rows.map((r: DatasetRow, idx: number) => ({
        rowIndex: idx, // Local row index within Sheet 1 (0-based)
        rawBytes: new Uint8Array(),
        rawText: '',
        cells: r.cells,
      }));

      await this.adapter.hydrateWindow(rowSlices, this.userEditedCells, targetSubId);
      const tHydrate1 = performance.now();

      const hasHeaders = this.columns.length > 0;
      const headerOffset = hasHeaders ? 1 : 0;

      this.activeWindow = {
        startRow: 0 + headerOffset,
        endRow: initialData.rowCount - 1 + headerOffset,
        rowCount: initialData.rowCount,
        generationToken: token,
      };

      const tEnd = performance.now();

      this.lastMetrics = {
        t1PartitionResolutionMs: initialData.metrics?.partitionResolutionMs,
        t2IndexLookupMs: initialData.metrics?.indexLookupMs,
        t3FileReadMs: initialData.metrics?.fileReadMs,
        t4DecodeMs: initialData.metrics?.decodeMs,
        t5ServerTotalMs: initialData.metrics?.serverTotalMs,
        t7WorkerDurationMs: initialData.metrics?.workerDurationMs,
        t8UniverHydrationMs: tHydrate1 - tHydrate0,
        e2eDurationMs: tEnd - tStart,
      };

      if (this.onMetricsUpdated && this.lastMetrics) {
        this.onMetricsUpdated(this.lastMetrics);
      }

      this.isMounted = true;
      return true;
    } finally {
      if (token === this.generationToken) {
        this.isLoading = false;
      }
    }
  }

  /**
   * Switches active virtual sheet tab.
   * Evicts active window from previous sheet, updates adapter subUnitId,
   * and hydrates initial visible viewport of new sheet without materializing 4M rows.
   */
  public async setActiveSheet(sheetIndexOrId: number | string): Promise<boolean> {
    if (!this.isMounted || this.partitions.length === 0) return false;

    let targetIndex = -1;
    if (typeof sheetIndexOrId === 'number') {
      targetIndex = Math.max(0, Math.min(sheetIndexOrId, this.partitions.length - 1));
    } else {
      targetIndex = this.partitions.findIndex((p) => p.sheetId === sheetIndexOrId);
      if (targetIndex === -1) {
        const match = sheetIndexOrId.match(/\d+/);
        if (match) {
          const parsed = parseInt(match[0], 10) - 1;
          if (parsed >= 0 && parsed < this.partitions.length) {
            targetIndex = parsed;
          }
        }
      }
    }

    if (targetIndex === -1) return false;
    if (targetIndex === this.activeSheetIndex && this.activeWindow !== null) {
      return true; // Already active with hydrated window
    }

    // 1. Evict active window from previous sheet to prevent memory accumulation across repeated switches
    if (this.activeWindow) {
      const oldPart = this.partitions[this.activeSheetIndex];
      if (oldPart) {
        await this.adapter.evictWindow(this.activeWindow.startRow, this.activeWindow.endRow, oldPart.sheetId);
      }
      this.activeWindow = null;
    }

    this.activeSheetIndex = targetIndex;
    const targetPartition = this.partitions[targetIndex];
    if (this.adapter.setActiveSubUnitId) {
      this.adapter.setActiveSubUnitId(targetPartition.sheetId);
    }
    if (this.adapter.activateSheet) {
      await this.adapter.activateSheet(targetPartition.sheetId);
    }

    // 2. Hydrate initial visible viewport on this sheet (local data rows 0..windowSize-1)
    const token = ++this.generationToken;
    this.isLoading = true;
    const tStart = performance.now();

    try {
      const windowSize = this.config.visibleRows + this.config.overscanRows * 2;
      const rowCountToFetch = Math.min(targetPartition.rowCount, windowSize);
      const globalStartRow = targetPartition.startRow;

      const data = await this.fetchRange(globalStartRow, rowCountToFetch, token);
      if (token !== this.generationToken || !data) return false;

      const hasHeaders = this.columns.length > 0;
      const headerOffset = hasHeaders ? 1 : 0;

      const rowSlices: RowSlice[] = data.rows.map((r: DatasetRow, idx: number) => ({
        rowIndex: idx, // Local row index within this virtual sheet (0-based)
        rawBytes: new Uint8Array(),
        rawText: '',
        cells: r.cells,
      }));

      const tHydrate0 = performance.now();
      await this.adapter.hydrateWindow(rowSlices, this.userEditedCells, targetPartition.sheetId);
      const tHydrate1 = performance.now();

      this.activeWindow = {
        startRow: 0 + headerOffset,
        endRow: data.rowCount - 1 + headerOffset,
        rowCount: data.rowCount,
        generationToken: token,
      };

      await this.adapter.scrollToRow(hasHeaders ? 1 : 0);

      const tEnd = performance.now();
      this.lastMetrics = {
        t1PartitionResolutionMs: data.metrics?.partitionResolutionMs,
        t2IndexLookupMs: data.metrics?.indexLookupMs,
        t3FileReadMs: data.metrics?.fileReadMs,
        t4DecodeMs: data.metrics?.decodeMs,
        t5ServerTotalMs: data.metrics?.serverTotalMs,
        t7WorkerDurationMs: data.metrics?.workerDurationMs,
        t8UniverHydrationMs: tHydrate1 - tHydrate0,
        e2eDurationMs: tEnd - tStart,
      };

      if (this.onMetricsUpdated && this.lastMetrics) {
        this.onMetricsUpdated(this.lastMetrics);
      }

      return true;
    } finally {
      if (token === this.generationToken) {
        this.isLoading = false;
      }
    }
  }

  /**
   * Applies a filtered or sorted query index to the active viewport.
   * Recalculates virtual sheets for the query result count and re-hydrates window starting at row 0.
   */
  public async setQuery(queryId: string | null, totalMatchedRows?: number): Promise<boolean> {
    this.currentQueryId = queryId;
    if (queryId && totalMatchedRows !== undefined) {
      this.totalDataRows = totalMatchedRows;
    } else if (!queryId) {
      this.totalDataRows = this.baseDatasetTotalRows;
    }

    if (!this.isMounted) return false;

    // Evict previous active window before switching query/dataset context
    if (this.activeWindow) {
      const oldPart = this.partitions[this.activeSheetIndex];
      if (oldPart) {
        await this.adapter.evictWindow(this.activeWindow.startRow, this.activeWindow.endRow, oldPart.sheetId);
      }
      this.activeWindow = null;
    }

    // Recalculate virtual sheet partitions for the query result count
    this.calculatePartitions();
    if (this.adapter.mountVirtualSheets) {
      await this.adapter.mountVirtualSheets(this.partitions, this.columnCount, this.columns);
    }

    // Reset active sheet to Sheet 1
    this.activeSheetIndex = 0;
    this.activeWindow = null;

    // Jump to row 0 / rehydrate initial window of Sheet 1
    return this.jumpToRow(0);
  }

  /**
   * Clears any active query filter/sort and restores base dataset viewport.
   */
  public async clearQuery(): Promise<boolean> {
    return this.setQuery(null);
  }

  /**
   * Jumps to an arbitrary spreadsheet row anywhere in the dataset (or query result).
   * Normal mode: accepts global DATA row number (0-based or 1-based, clamped to 0..totalDataRows-1).
   * Query mode: accepts query result position (0-based or 1-based, clamped to 0..totalMatchedRows-1).
   * Automatically determines virtual sheet partition and translates to local coordinates.
   */
  public async jumpToRow(targetSpreadsheetRow: number): Promise<boolean> {
    if (!this.isMounted || this.totalDataRows === 0 || this.partitions.length === 0) return false;

    // Global data row (or query result position)
    const globalDataRow = Math.max(0, Math.min(targetSpreadsheetRow, this.totalDataRows - 1));
    const targetSheetIndex = Math.floor(globalDataRow / ROWS_PER_VIRTUAL_SHEET);
    const clampedSheetIndex = Math.max(0, Math.min(targetSheetIndex, this.partitions.length - 1));
    const targetPartition = this.partitions[clampedSheetIndex];

    const localDataRow = globalDataRow % ROWS_PER_VIRTUAL_SHEET;
    const clampedLocalDataRow = Math.max(0, Math.min(localDataRow, targetPartition.rowCount - 1));

    const hasHeaders = this.columns.length > 0;
    const headerOffset = hasHeaders ? 1 : 0;
    const localUniverRow = clampedLocalDataRow + headerOffset;

    // If target is on a different virtual sheet, switch active sheet
    if (clampedSheetIndex !== this.activeSheetIndex) {
      if (this.activeWindow) {
        const oldPart = this.partitions[this.activeSheetIndex];
        if (oldPart) {
          await this.adapter.evictWindow(this.activeWindow.startRow, this.activeWindow.endRow, oldPart.sheetId);
        }
        this.activeWindow = null;
      }
      this.activeSheetIndex = clampedSheetIndex;
      if (this.adapter.setActiveSubUnitId) {
        this.adapter.setActiveSubUnitId(targetPartition.sheetId);
      }
      if (this.adapter.activateSheet) {
        await this.adapter.activateSheet(targetPartition.sheetId);
      }
    }

    const windowSize = this.config.visibleRows + this.config.overscanRows * 2;
    let startLocalRow = Math.max(0, clampedLocalDataRow - this.config.overscanRows);
    let endLocalRow = Math.min(targetPartition.rowCount - 1, startLocalRow + windowSize - 1);

    if (endLocalRow - startLocalRow + 1 < windowSize && startLocalRow > 0) {
      startLocalRow = Math.max(0, endLocalRow - windowSize + 1);
    }

    const token = ++this.generationToken;
    this.isLoading = true;
    const tStart = performance.now();

    try {
      const globalStartRow = targetPartition.startRow + startLocalRow;
      const rowCountToFetch = endLocalRow - startLocalRow + 1;

      const data = await this.fetchRange(globalStartRow, rowCountToFetch, token);
      if (token !== this.generationToken || !data) return false;

      // 1. Evict stale non-overlapping rows from Univer on this sheet
      if (this.activeWindow) {
        const oldStart = this.activeWindow.startRow;
        const oldEnd = this.activeWindow.endRow;
        const newStart = startLocalRow + headerOffset;
        const newEnd = endLocalRow + headerOffset;

        if (newStart > oldEnd || newEnd < oldStart) {
          await this.adapter.evictWindow(oldStart, oldEnd, targetPartition.sheetId);
        } else {
          if (oldStart < newStart) await this.adapter.evictWindow(oldStart, newStart - 1, targetPartition.sheetId);
          if (oldEnd > newEnd) await this.adapter.evictWindow(newEnd + 1, oldEnd, targetPartition.sheetId);
        }
      }

      // 2. Hydrate new window
      const tHydrate0 = performance.now();
      const rowSlices: RowSlice[] = data.rows.map((r: DatasetRow, idx: number) => ({
        rowIndex: startLocalRow + idx, // Local row index within active virtual sheet
        rawBytes: new Uint8Array(),
        rawText: '',
        cells: r.cells,
      }));

      await this.adapter.hydrateWindow(rowSlices, this.userEditedCells, targetPartition.sheetId);
      const tHydrate1 = performance.now();

      this.activeWindow = {
        startRow: startLocalRow + headerOffset,
        endRow: endLocalRow + headerOffset,
        rowCount: data.rowCount,
        generationToken: token,
      };

      // 3. Scroll Univer canvas
      await this.adapter.scrollToRow(localUniverRow);

      const tEnd = performance.now();
      this.lastMetrics = {
        t1PartitionResolutionMs: data.metrics?.partitionResolutionMs,
        t2IndexLookupMs: data.metrics?.indexLookupMs,
        t3FileReadMs: data.metrics?.fileReadMs,
        t4DecodeMs: data.metrics?.decodeMs,
        t5ServerTotalMs: data.metrics?.serverTotalMs,
        t7WorkerDurationMs: data.metrics?.workerDurationMs,
        t8UniverHydrationMs: tHydrate1 - tHydrate0,
        e2eDurationMs: tEnd - tStart,
      };

      if (this.onMetricsUpdated && this.lastMetrics) {
        this.onMetricsUpdated(this.lastMetrics);
      }

      return true;
    } finally {
      if (token === this.generationToken) {
        this.isLoading = false;
      }
    }
  }

  /**
   * Handles user scrolling on Univer canvas.
   * Coordinates are relative to the active virtual sheet.
   * If scroll position approaches the edge of the active window, slides the window smoothly.
   */
  public async handleScroll(firstVisibleRow: number): Promise<void> {
    if (this.isLoading || !this.activeWindow || this.totalDataRows === 0 || this.partitions.length === 0) return;

    const currentPartition = this.partitions[this.activeSheetIndex];
    if (!currentPartition) return;

    const hasHeaders = this.columns.length > 0;
    const headerOffset = hasHeaders ? 1 : 0;

    const currentStart = this.activeWindow.startRow;
    const currentEnd = this.activeWindow.endRow;

    const nearTop =
      firstVisibleRow <= currentStart + this.config.marginThreshold && currentStart > headerOffset;
    const nearBottom =
      firstVisibleRow >= currentEnd - this.config.marginThreshold &&
      currentEnd < currentPartition.rowCount + headerOffset - 1;

    if (!nearTop && !nearBottom) {
      return; // Comfortably inside active hydrated window
    }

    const localDataRow = Math.max(0, firstVisibleRow - headerOffset);
    const windowSize = this.config.visibleRows + this.config.overscanRows * 2;
    let startLocalRow = Math.max(0, localDataRow - this.config.overscanRows);
    let endLocalRow = Math.min(currentPartition.rowCount - 1, startLocalRow + windowSize - 1);

    if (endLocalRow - startLocalRow + 1 < windowSize && startLocalRow > 0) {
      startLocalRow = Math.max(0, endLocalRow - windowSize + 1);
    }

    const token = ++this.generationToken;
    this.isLoading = true;
    const tStart = performance.now();

    try {
      const globalStartRow = currentPartition.startRow + startLocalRow;
      const rowCountToFetch = endLocalRow - startLocalRow + 1;

      const data = await this.fetchRange(globalStartRow, rowCountToFetch, token);
      if (token !== this.generationToken || !data) return;

      const newStart = startLocalRow + headerOffset;
      const newEnd = endLocalRow + headerOffset;

      // Evict stale rows from this sheet
      if (currentStart < newStart) {
        await this.adapter.evictWindow(currentStart, newStart - 1, currentPartition.sheetId);
      }
      if (currentEnd > newEnd) {
        await this.adapter.evictWindow(newEnd + 1, currentEnd, currentPartition.sheetId);
      }

      // Hydrate new rows into this sheet
      const tHydrate0 = performance.now();
      const rowSlices: RowSlice[] = data.rows.map((r: DatasetRow, idx: number) => ({
        rowIndex: startLocalRow + idx,
        rawBytes: new Uint8Array(),
        rawText: '',
        cells: r.cells,
      }));

      await this.adapter.hydrateWindow(rowSlices, this.userEditedCells, currentPartition.sheetId);
      const tHydrate1 = performance.now();

      this.activeWindow = {
        startRow: newStart,
        endRow: newEnd,
        rowCount: data.rowCount,
        generationToken: token,
      };

      const tEnd = performance.now();
      this.lastMetrics = {
        t1PartitionResolutionMs: data.metrics?.partitionResolutionMs,
        t2IndexLookupMs: data.metrics?.indexLookupMs,
        t3FileReadMs: data.metrics?.fileReadMs,
        t4DecodeMs: data.metrics?.decodeMs,
        t5ServerTotalMs: data.metrics?.serverTotalMs,
        t7WorkerDurationMs: data.metrics?.workerDurationMs,
        t8UniverHydrationMs: tHydrate1 - tHydrate0,
        e2eDurationMs: tEnd - tStart,
      };

      if (this.onMetricsUpdated && this.lastMetrics) {
        this.onMetricsUpdated(this.lastMetrics);
      }
    } finally {
      if (token === this.generationToken) {
        this.isLoading = false;
      }
    }
  }

  /**
   * Internal range fetch dispatcher: delegates to Dedicated Worker if available,
   * or falls back to direct fetch.
   */
  private fetchRange(
    startRow: number,
    rowCount: number,
    generationToken: number
  ): Promise<any> {
    const worker = this.worker;
    if (worker) {
      return new Promise((resolve) => {
        const handler = (e: MessageEvent) => {
          const msg = e.data;
          if (!msg) return;

          if (msg.generationToken === generationToken) {
            if (msg.type === 'VIEWPORT_DATA') {
              worker.removeEventListener('message', handler);
              resolve(msg);
            } else if (msg.type === 'ERROR') {
              worker.removeEventListener('message', handler);
              resolve(null);
            }
          }
        };

        worker.addEventListener('message', handler);
        worker.postMessage({
          type: 'FETCH_VIEWPORT_RANGE',
          datasetId: this.datasetId,
          queryId: this.currentQueryId || undefined,
          startRow,
          rowCount,
          generationToken,
          apiBaseUrl: this.config.apiBaseUrl || '',
        });
      });
    }

    // Direct fetch fallback for non-worker contexts
    const baseUrl = this.config.apiBaseUrl || '';
    const queryParam = this.currentQueryId ? `&queryId=${encodeURIComponent(this.currentQueryId)}` : '';
    const url = `${baseUrl}/api/v1/datasets/${this.datasetId}/range?startRow=${startRow}&rowCount=${rowCount}${queryParam}`;
    return fetch(url).then((res) => (res.ok ? res.json() : null));
  }

  private handleWorkerMessage(e: MessageEvent): void {
    const data = e.data;
    if (!data) return;
    // Log worker lifecycle notices or diagnostics
  }

  public destroy(): void {
    if (this.worker) {
      this.worker.terminate();
      this.worker = null;
    }
    this.userEditedCells.clear();
    this.activeWindow = null;
    this.isMounted = false;
  }
}
