// ==============================================================================
// ENTERPRISE CALCULATION QUEUE TYPES (PHASE 4 IMPLEMENTATION)
// ==============================================================================

export interface CellMutationDelta {
  sheetId: string;
  row: number;
  col: number;
  f?: string | null;  // Formula text (e.g. "=SUM(A1:A10)")
  v?: any;            // Raw value or optimistic display value
  t?: number;         // CellValueType (1=String, 2=Number, 3=Boolean, 4=Formula)
}

export interface CalculationJobPayload {
  workbookId: string;
  revisionId: number;      // Monotonic mutation revision to calculate up to
  enqueuedAt: number;      // Epoch timestamp ms
  actorUserId: string;
  deltaSummary: {
    sheetCount: number;
    cellCount: number;
  };
}

export interface CalculationCellPatch {
  sheetId: string;
  row: number;
  col: number;
  v: any;                  // Computed value
  t?: number;              // Type (Number, String, Error, etc.)
}

export interface CalculationJobResult {
  status: 'SUCCESS' | 'SUPERSEDED' | 'FAILED';
  workbookId: string;
  revisionId: number;
  executionMs: number;
  patchCount: number;
  strategy: string;
  fencingToken?: string;
  error?: string;
}

export const QUEUE_PARTITIONS_COUNT = 4;
export const BASE_QUEUE_NAME = 'spreadsheet-calc-queue';

export function getQueueNameForPartition(partitionIndex: number): string {
  return `${BASE_QUEUE_NAME}-${partitionIndex}`;
}
