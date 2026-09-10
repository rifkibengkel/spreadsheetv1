/**
 * OPFS Storage Layer Type Contracts & Invariants
 * Phase 4 — Architecture D
 */

export const OPFS_DATA_FILE = 'table_data.bin';
export const OPFS_INDEX_FILE = 'table_index.idx';
export const OPFS_META_FILE = 'table_meta.json';
export const OPFS_MAGIC_V1 = 'SHEET_OPFS_V1';

export type CommitStatus = 'UNINITIALIZED' | 'IMPORTING' | 'VALIDATING' | 'COMMITTED' | 'FAILED';

export interface OPFSMetadata {
  magic: typeof OPFS_MAGIC_V1;
  version: number;
  sessionId?: string; // Step 1.1: Unique session / generation identity
  fileName: string;
  rowCount: number; // Number of DATA rows (excluding header)
  columnCount: number; // Number of columns in header
  headers: string[]; // Parsed header column names
  dataSizeBytes: number; // Exact byte length of table_data.bin
  indexSizeBytes: number; // Exact byte length of table_index.idx: (rowCount + 1) * 8
  status: CommitStatus;
  createdAt: string;
  committedAt?: string;
  errorMessage?: string;
}

export interface ImportProgress {
  bytesProcessed: number;
  totalBytes: number;
  rowsProcessed: number;
  percent: number;
  phase: 'READING_HEADER' | 'STREAMING_ROWS' | 'BUILDING_INDEX' | 'VALIDATING' | 'COMMITTED';
}

export interface ImportResult {
  success: boolean;
  metadata: OPFSMetadata;
  durationMs: number;
  peakMemoryMB?: number;
  error?: string;
}

export interface RowSlice {
  rowIndex: number;
  rawBytes: Uint8Array;
  rawText: string;
  cells: string[];
}
