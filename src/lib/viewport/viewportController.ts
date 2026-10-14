/**
 * Viewport Controller (Architecture D)
 * Orchestrates dynamic sliding window over OPFS large-dataset storage.
 * Provides generation-token race protection and user edit persistence across eviction cycles.
 */

import { DatasetReader } from '../storage/opfs/datasetReader';
import { UniverViewportAdapter } from './univerViewportAdapter';
import { CellEditRecord, DEFAULT_VIEWPORT_CONFIG, ViewportConfig, ViewportWindow, VirtualSheetPartition, ROWS_PER_VIRTUAL_SHEET } from './types';
import { OPFSMetadata, RowSlice } from '../storage/opfs/types';

export class ViewportController {
  private reader: DatasetReader;
  private adapter: UniverViewportAdapter;
  private config: ViewportConfig;
  private activeWindow: ViewportWindow | null = null;
  private generationToken: number = 0;
  private userEditedCells: Map<string, CellEditRecord> = new Map();
  private isLoading: boolean = false;
  private metadata: OPFSMetadata | null = null;
  private partitions: VirtualSheetPartition[] = [];
  private activeSheetIndex: number = 0;
  private isMounted: boolean = false;

  constructor(
    reader: DatasetReader,
    adapter: UniverViewportAdapter,
    config: ViewportConfig = DEFAULT_VIEWPORT_CONFIG
  ) {
    this.reader = reader;
    this.adapter = adapter;
    this.config = config;
  }

  public setAdapter(adapter: UniverViewportAdapter): void {
    this.adapter = adapter;
  }

  public getMetadata(): OPFSMetadata | null {
    return this.metadata || this.reader.getMetadata();
  }

  public getActiveWindow(): ViewportWindow | null {
    return this.activeWindow;
  }

  public getUserEditCount(): number {
    return this.userEditedCells.size;
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

  public isReady(): boolean {
    return this.isMounted;
  }

  /**
   * Mounts the dataset to Univer and hydrates the initial window (rows 0..999).
   */
  public async mount(metadata?: OPFSMetadata): Promise<boolean> {
    const meta = metadata || this.reader.getMetadata();
    if (!meta || meta.status !== 'COMMITTED') {
      console.warn('[ViewportController] Cannot mount uncommitted dataset.');
      return false;
    }
    this.metadata = meta;

    if (meta.rowCount > ROWS_PER_VIRTUAL_SHEET) {
      const sheetCount = Math.max(1, Math.ceil(meta.rowCount / ROWS_PER_VIRTUAL_SHEET));
      this.partitions = [];
      for (let s = 0; s < sheetCount; s++) {
        const startRow = s * ROWS_PER_VIRTUAL_SHEET;
        const endRow = Math.min(meta.rowCount - 1, (s + 1) * ROWS_PER_VIRTUAL_SHEET - 1);
        const rowCount = Math.max(0, endRow - startRow + 1);
        this.partitions.push({
          sheetId: `sheet-${s + 1}`,
          sheetIndex: s,
          startRow,
          endRow,
          rowCount,
        });
      }
    } else {
      this.partitions = [{
        sheetId: 'sheet-1',
        sheetIndex: 0,
        startRow: 0,
        endRow: meta.rowCount - 1,
        rowCount: meta.rowCount,
      }];
    }
    this.activeSheetIndex = 0;

    // 1. Mount virtual sheet dimensions in Univer
    if (this.partitions.length > 1 && this.adapter.mountVirtualSheets) {
      await this.adapter.mountVirtualSheets(this.partitions, meta.columnCount, meta.headers);
    } else {
      await this.adapter.mountVirtualSheet(meta);
    }

    // 2. Fetch initial rows from OPFS for Sheet 1 (data rows 0..999)
    const initialCount = Math.min(this.config.bufferSize, meta.rowCount);
    const rows = await this.reader.readRowRange(0, initialCount - 1);

    // 3. Hydrate initial window into Univer
    const primaryTargetSubId = this.partitions[0]?.sheetId || this.adapter.getActiveSubUnitId();
    await this.adapter.hydrateWindow(rows, this.userEditedCells, primaryTargetSubId);

    const hasHeaders = (meta.headers || []).length > 0;
    const headerOffset = hasHeaders ? 1 : 0;

    const token = ++this.generationToken;
    this.activeWindow = {
      startRow: headerOffset,
      endRow: headerOffset + rows.length - 1,
      bufferSize: rows.length,
      generationToken: token,
    };

    this.isMounted = true;
    return true;
  }

  /**
   * Switches active virtual sheet tab.
   */
  public async setActiveSheet(sheetIndexOrId: number | string): Promise<boolean> {
    if (!this.isMounted) return false;
    const meta = this.getMetadata();
    if (!meta || this.partitions.length === 0) return false;

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
      return true;
    }

    // Evict active window from previous sheet
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

    // Hydrate initial visible viewport on this sheet
    const token = ++this.generationToken;
    this.isLoading = true;

    try {
      const initialCount = Math.min(this.config.bufferSize, targetPartition.rowCount);
      const globalStartRow = targetPartition.startRow;
      const rows = await this.reader.readRowRange(globalStartRow, globalStartRow + initialCount - 1);
      if (token !== this.generationToken || !rows) return false;

      const hasHeaders = (meta.headers || []).length > 0;
      const headerOffset = hasHeaders ? 1 : 0;

      // Adjust rowIndex to local virtual sheet coordinates
      const localRows: RowSlice[] = rows.map((r, idx) => ({
        ...r,
        rowIndex: idx,
      }));

      await this.adapter.hydrateWindow(localRows, this.userEditedCells, targetPartition.sheetId);

      this.activeWindow = {
        startRow: headerOffset,
        endRow: headerOffset + rows.length - 1,
        bufferSize: rows.length,
        generationToken: token,
      };

      await this.adapter.scrollToRow(headerOffset);
      return true;
    } finally {
      if (token === this.generationToken) {
        this.isLoading = false;
      }
    }
  }

  /**
   * Jumps to an arbitrary row anywhere across the 4M+ dataset.
   * targetRow is the spreadsheet 1-based or 0-based index.
   */
  public async jumpToRow(targetSpreadsheetRow: number): Promise<boolean> {
    if (!this.isMounted) return false;
    const meta = this.getMetadata();
    if (!meta || !this.adapter.isReady()) return false;

    const hasHeaders = (meta.headers || []).length > 0;
    const headerOffset = hasHeaders ? 1 : 0;
    const totalSheetRows = meta.rowCount + headerOffset;

    const clampedTarget = Math.max(headerOffset, Math.min(targetSpreadsheetRow, totalSheetRows - 1));
    const targetDataRow = clampedTarget - headerOffset;

    // Determine target partition
    const targetSheetIndex = Math.floor(targetDataRow / ROWS_PER_VIRTUAL_SHEET);
    const clampedSheetIndex = Math.max(0, Math.min(targetSheetIndex, this.partitions.length - 1));
    const targetPartition = this.partitions[clampedSheetIndex] || {
      sheetId: 'sheet-1',
      sheetIndex: 0,
      startRow: 0,
      endRow: meta.rowCount - 1,
      rowCount: meta.rowCount,
    };

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

    const localDataRow = targetDataRow - targetPartition.startRow;
    let startLocalRow = Math.max(0, localDataRow - 50);
    let endLocalRow = Math.min(targetPartition.rowCount - 1, startLocalRow + this.config.bufferSize - 1);

    if (endLocalRow - startLocalRow + 1 < this.config.bufferSize && startLocalRow > 0) {
      startLocalRow = Math.max(0, endLocalRow - this.config.bufferSize + 1);
    }

    const token = ++this.generationToken;
    this.isLoading = true;

    try {
      const globalStartRow = targetPartition.startRow + startLocalRow;
      const globalEndRow = targetPartition.startRow + endLocalRow;
      const rows = await this.reader.readRowRange(globalStartRow, globalEndRow);

      if (token !== this.generationToken) {
        return false;
      }

      // Evict old window if outside new window
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

      const localRows: RowSlice[] = rows.map((r, idx) => ({
        ...r,
        rowIndex: startLocalRow + idx,
      }));

      await this.adapter.hydrateWindow(localRows, this.userEditedCells, targetPartition.sheetId);

      this.activeWindow = {
        startRow: startLocalRow + headerOffset,
        endRow: endLocalRow + headerOffset,
        bufferSize: rows.length,
        generationToken: token,
      };

      const localUniverRow = localDataRow + headerOffset;
      await this.adapter.scrollToRow(localUniverRow);

      return true;
    } finally {
      if (token === this.generationToken) {
        this.isLoading = false;
      }
    }
  }

  /**
   * Handles user scrolling. If viewport approaches boundary of active window,
   * slides the window asynchronously.
   */
  public async handleScroll(firstVisibleRow: number): Promise<void> {
    if (!this.isMounted || this.isLoading || !this.activeWindow || !this.metadata || this.partitions.length === 0) return;

    const currentPartition = this.partitions[this.activeSheetIndex];
    if (!currentPartition) return;

    const hasHeaders = (this.metadata.headers || []).length > 0;
    const headerOffset = hasHeaders ? 1 : 0;

    const currentStart = this.activeWindow.startRow;
    const currentEnd = this.activeWindow.endRow;

    // Check if scrolling near top boundary or bottom boundary
    const nearTop = firstVisibleRow <= currentStart + this.config.marginThreshold && currentStart > headerOffset;
    const nearBottom = firstVisibleRow >= currentEnd - this.config.marginThreshold && currentEnd < currentPartition.rowCount + headerOffset - 1;

    if (!nearTop && !nearBottom) {
      return; // Still comfortably within active window
    }

    // Target center around firstVisibleRow
    const targetDataRow = Math.max(0, firstVisibleRow - headerOffset);
    let startLocalRow = Math.max(0, targetDataRow - Math.floor(this.config.bufferSize / 3));
    let endLocalRow = Math.min(currentPartition.rowCount - 1, startLocalRow + this.config.bufferSize - 1);

    if (endLocalRow - startLocalRow + 1 < this.config.bufferSize && startLocalRow > 0) {
      startLocalRow = Math.max(0, endLocalRow - this.config.bufferSize + 1);
    }

    const token = ++this.generationToken;
    this.isLoading = true;

    try {
      const globalStartRow = currentPartition.startRow + startLocalRow;
      const globalEndRow = currentPartition.startRow + endLocalRow;
      const rows = await this.reader.readRowRange(globalStartRow, globalEndRow);
      if (token !== this.generationToken) return;

      const newStart = startLocalRow + headerOffset;
      const newEnd = endLocalRow + headerOffset;

      // Evict stale rows
      if (currentStart < newStart) {
        await this.adapter.evictWindow(currentStart, newStart - 1, currentPartition.sheetId);
      }
      if (currentEnd > newEnd) {
        await this.adapter.evictWindow(newEnd + 1, currentEnd, currentPartition.sheetId);
      }

      const localRows: RowSlice[] = rows.map((r, idx) => ({
        ...r,
        rowIndex: startLocalRow + idx,
      }));

      await this.adapter.hydrateWindow(localRows, this.userEditedCells, currentPartition.sheetId);

      this.activeWindow = {
        startRow: newStart,
        endRow: newEnd,
        bufferSize: rows.length,
        generationToken: token,
      };
    } finally {
      if (token === this.generationToken) {
        this.isLoading = false;
      }
    }
  }

  /**
   * Records a user edit to a cell.
   * Preserved across window eviction cycles and overlaid when rows are re-hydrated.
   */
  public recordUserEdit(
    sheetId: string,
    row: number,
    col: number,
    value: any,
    formula?: string
  ): void {
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

  public getCellEdit(row: number, col: number, sheetId?: string): CellEditRecord | undefined {
    if (sheetId) {
      return this.userEditedCells.get(`${sheetId}_${row}_${col}`);
    }
    const currentSubId = this.partitions[this.activeSheetIndex]?.sheetId || 'sheet-1';
    return this.userEditedCells.get(`${currentSubId}_${row}_${col}`) || this.userEditedCells.get(`${row}_${col}`);
  }

  public getAllUserEdits(): Map<string, CellEditRecord> {
    return new Map(this.userEditedCells);
  }

  public clearUserEdits(): void {
    this.userEditedCells.clear();
  }
}

