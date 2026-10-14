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

export const ROWS_PER_VIRTUAL_SHEET = 4_000_000;

export interface VirtualSheetPartition {
  sheetId: string;
  sheetIndex: number;
  startRow: number;
  endRow: number;
  rowCount: number;
}

export interface IUniverViewportAdapter {
  mountVirtualSheet(metadata: OPFSMetadata): Promise<void>;
  mountVirtualSheets?(
    partitions: VirtualSheetPartition[],
    columnCount: number,
    headers: string[]
  ): Promise<void>;
  setActiveSubUnitId?(sheetId: string): void;
  hydrateWindow(
    rows: RowSlice[],
    userEdits: Map<string, CellEditRecord>,
    targetSheetId?: string
  ): Promise<void>;
  evictWindow(startRow: number, endRow: number, targetSheetId?: string): Promise<void>;
  scrollToRow(targetRow: number): Promise<boolean>;
  getLogicalRowCount(): number;
  getActiveUnitId(): string;
  getActiveSubUnitId(): string;
  activateSheet?(sheetId: string): Promise<void>;
  isReady(): boolean;
}

