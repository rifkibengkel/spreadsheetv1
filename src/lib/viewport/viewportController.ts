/**
 * Viewport Controller (Architecture D)
 * Orchestrates dynamic sliding window over OPFS large-dataset storage.
 * Provides generation-token race protection and user edit persistence across eviction cycles.
 */

import { DatasetReader } from '../storage/opfs/datasetReader';
import { UniverViewportAdapter } from './univerViewportAdapter';
import { CellEditRecord, DEFAULT_VIEWPORT_CONFIG, ViewportConfig, ViewportWindow } from './types';
import { OPFSMetadata } from '../storage/opfs/types';

export class ViewportController {
  private reader: DatasetReader;
  private adapter: UniverViewportAdapter;
  private config: ViewportConfig;
  private activeWindow: ViewportWindow | null = null;
  private generationToken: number = 0;
  private userEditedCells: Map<string, CellEditRecord> = new Map();
  private isLoading: boolean = false;
  private metadata: OPFSMetadata | null = null;

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

    const token = ++this.generationToken;

    // 1. Mount virtual sheet dimensions in Univer (e.g. 4,149,891 rows)
    await this.adapter.mountVirtualSheet(meta);

    // 2. Fetch initial 1,000 rows from OPFS (data rows 0..999)
    const initialCount = Math.min(this.config.bufferSize, meta.rowCount);
    const rows = await this.reader.readRowRange(0, initialCount - 1);

    if (token !== this.generationToken) {
      // Obsolete request superseded by another action
      return false;
    }

    // 3. Hydrate initial window into Univer
    await this.adapter.hydrateWindow(rows, this.userEditedCells);

    const hasHeaders = (meta.headers || []).length > 0;
    const headerOffset = hasHeaders ? 1 : 0;

    this.activeWindow = {
      startRow: headerOffset,
      endRow: headerOffset + rows.length - 1,
      bufferSize: rows.length,
      generationToken: token,
    };

    return true;
  }

  /**
   * Jumps to an arbitrary row anywhere across the 4M+ dataset.
   * targetRow is the spreadsheet 1-based or 0-based index.
   */
  public async jumpToRow(targetSpreadsheetRow: number): Promise<boolean> {
    const meta = this.getMetadata();
    if (!meta || !this.adapter.isReady()) return false;

    const hasHeaders = (meta.headers || []).length > 0;
    const headerOffset = hasHeaders ? 1 : 0;
    const totalSheetRows = meta.rowCount + headerOffset;

    const clampedTarget = Math.max(headerOffset, Math.min(targetSpreadsheetRow, totalSheetRows - 1));
    const targetDataRow = clampedTarget - headerOffset;

    const halfWindow = Math.floor(this.config.bufferSize / 2);
    let startDataRow = Math.max(0, targetDataRow - 50); // slight lookback
    let endDataRow = Math.min(meta.rowCount - 1, startDataRow + this.config.bufferSize - 1);

    // Adjust if near end of dataset
    if (endDataRow - startDataRow + 1 < this.config.bufferSize && startDataRow > 0) {
      startDataRow = Math.max(0, endDataRow - this.config.bufferSize + 1);
    }

    const token = ++this.generationToken;
    this.isLoading = true;

    try {
      // 1. Read slice from OPFS
      const rows = await this.reader.readRowRange(startDataRow, endDataRow);

      if (token !== this.generationToken) {
        return false;
      }

      // 2. Evict old window if outside new window
      if (this.activeWindow) {
        const oldStart = this.activeWindow.startRow;
        const oldEnd = this.activeWindow.endRow;
        const newStart = startDataRow + headerOffset;
        const newEnd = endDataRow + headerOffset;

        // Evict non-overlapping sections
        if (newStart > oldEnd || newEnd < oldStart) {
          await this.adapter.evictWindow(oldStart, oldEnd);
        } else {
          if (oldStart < newStart) await this.adapter.evictWindow(oldStart, newStart - 1);
          if (oldEnd > newEnd) await this.adapter.evictWindow(newEnd + 1, oldEnd);
        }
      }

      // 3. Hydrate new window
      await this.adapter.hydrateWindow(rows, this.userEditedCells);

      // 4. Update active window tracker
      this.activeWindow = {
        startRow: startDataRow + headerOffset,
        endRow: endDataRow + headerOffset,
        bufferSize: rows.length,
        generationToken: token,
      };

      // 5. Scroll canvas to target row
      await this.adapter.scrollToRow(clampedTarget);

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
    if (this.isLoading || !this.activeWindow || !this.metadata) return;

    const meta = this.metadata;
    const hasHeaders = (meta.headers || []).length > 0;
    const headerOffset = hasHeaders ? 1 : 0;

    const currentStart = this.activeWindow.startRow;
    const currentEnd = this.activeWindow.endRow;

    // Check if scrolling near top boundary or bottom boundary
    const nearTop = firstVisibleRow <= currentStart + this.config.marginThreshold && currentStart > headerOffset;
    const nearBottom = firstVisibleRow >= currentEnd - this.config.marginThreshold && currentEnd < meta.rowCount + headerOffset - 1;

    if (!nearTop && !nearBottom) {
      return; // Still comfortably within active window
    }

    // Target center around firstVisibleRow
    const targetDataRow = Math.max(0, firstVisibleRow - headerOffset);
    let startDataRow = Math.max(0, targetDataRow - Math.floor(this.config.bufferSize / 3));
    let endDataRow = Math.min(meta.rowCount - 1, startDataRow + this.config.bufferSize - 1);

    if (endDataRow - startDataRow + 1 < this.config.bufferSize && startDataRow > 0) {
      startDataRow = Math.max(0, endDataRow - this.config.bufferSize + 1);
    }

    const token = ++this.generationToken;
    this.isLoading = true;

    try {
      const rows = await this.reader.readRowRange(startDataRow, endDataRow);
      if (token !== this.generationToken) return;

      const newStart = startDataRow + headerOffset;
      const newEnd = endDataRow + headerOffset;

      // Evict stale rows
      if (currentStart < newStart) {
        await this.adapter.evictWindow(currentStart, newStart - 1);
      }
      if (currentEnd > newEnd) {
        await this.adapter.evictWindow(newEnd + 1, currentEnd);
      }

      // Hydrate new rows
      await this.adapter.hydrateWindow(rows, this.userEditedCells);

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
    const key = `${row}_${col}`;
    this.userEditedCells.set(key, {
      sheetId,
      row,
      col,
      value,
      formula,
      timestamp: Date.now(),
    });
  }

  public getCellEdit(row: number, col: number): CellEditRecord | undefined {
    return this.userEditedCells.get(`${row}_${col}`);
  }

  public clearUserEdits(): void {
    this.userEditedCells.clear();
  }
}
