/**
 * Viewport Architecture Types (Architecture D)
 * Defines interfaces for sliding window, generation token, and user edit tracking.
 */

import { OPFSMetadata, RowSlice } from '../storage/opfs/types';

export interface ViewportWindow {
  startRow: number;
  endRow: number;
  bufferSize: number;
  generationToken: number;
}

export interface CellEditRecord {
  sheetId: string;
  row: number;
  col: number;
  value: any;
  formula?: string;
  timestamp: number;
}

export interface ViewportConfig {
  /** Total number of rows hydrated in the sparse cell matrix (default: 1000) */
  bufferSize: number;
  /** Distance from window boundary before triggering a slide (default: 200) */
  marginThreshold: number;
  /** Maximum rows to retain before evicting stale rows (default: 1500) */
  maxRetainedRows: number;
}

export const DEFAULT_VIEWPORT_CONFIG: ViewportConfig = {
  bufferSize: 1000,
  marginThreshold: 200,
  maxRetainedRows: 1500,
};

export interface IUniverViewportAdapter {
  mountVirtualSheet(metadata: OPFSMetadata): Promise<void>;
  hydrateWindow(rows: RowSlice[], userEdits: Map<string, CellEditRecord>): Promise<void>;
  evictWindow(startRow: number, endRow: number): Promise<void>;
  scrollToRow(targetRow: number): Promise<boolean>;
  getLogicalRowCount(): number;
  getActiveUnitId(): string;
  getActiveSubUnitId(): string;
  isReady(): boolean;
}
