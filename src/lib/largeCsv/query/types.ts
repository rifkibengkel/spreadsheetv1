/**
 * Data Contracts and Interfaces for Server-Side Filtering, Sorting & Query Indexing
 * Phase 1C — Enterprise Large CSV Viewport Engine
 */

export const QUERY_INDEX_MAGIC = 'SHEET_QIDX_V1';
export const QUERY_INDEX_FORMAT_VERSION = 1;
export const QUERY_INDEX_HEADER_SIZE = 2048;
export const MAX_CONCURRENT_QUERY_JOBS = 4;
export const MAX_SORT_CHUNK_SIZE = 250_000;

export type FilterOperator =
  | 'EQUALS'
  | 'NOT_EQUALS'
  | 'CONTAINS'
  | 'NOT_CONTAINS'
  | 'STARTS_WITH'
  | 'ENDS_WITH'
  | 'GREATER_THAN'
  | 'GREATER_THAN_OR_EQUAL'
  | 'LESS_THAN'
  | 'LESS_THAN_OR_EQUAL'
  | 'IS_EMPTY'
  | 'IS_NOT_EMPTY'
  | 'IN';

export interface FilterPredicate {
  /** Target column header name (e.g. "Registration_State") */
  column: string;
  /** Filter operator */
  operator: FilterOperator;
  /** Value or array of values to match against */
  value?: string | string[];
  /** Whether string comparisons should be case sensitive (default: false) */
  caseSensitive?: boolean;
  /** If true, parse values as floating-point numbers before comparison */
  numeric?: boolean;
}

export interface FilterDefinition {
  /** Conjunction between multiple predicates: 'AND' requires all, 'OR' requires any */
  conjunction: 'AND' | 'OR';
  /** List of predicates to apply */
  predicates: FilterPredicate[];
}

export type SortDirection = 'ASC' | 'DESC';
export type SortType = 'auto' | 'string' | 'number' | 'date';

export interface SortColumn {
  /** Target column header name to sort by */
  column: string;
  /** Sort order: ASC or DESC */
  direction: SortDirection;
  /** Comparison type (default: auto) */
  type?: SortType;
}

export interface SortDefinition {
  /** Ordered list of sort columns (primary, secondary, etc.) */
  columns: SortColumn[];
}

export interface QueryDefinition {
  /** Optional filter definition */
  filter?: FilterDefinition;
  /** Optional sort definition */
  sort?: SortDefinition;
}

export type QueryStatus = 'GENERATING' | 'READY' | 'FAILED';

export interface QueryMetadata {
  queryId: string;
  datasetId: string;
  status: QueryStatus;
  totalDatasetRows: number;
  matchedRows: number;
  indexFilePath: string;
  metadataFilePath: string;
  createdAt: number;
  completedAt?: number;
  executionTimeMs?: number;
  queryDef: QueryDefinition;
  errorMessage?: string;
}

export class DatasetQueryError extends Error {
  public statusCode: number;
  public code: string;

  constructor(message: string, statusCode = 400, code = 'INVALID_QUERY_REQUEST') {
    super(message);
    this.name = 'DatasetQueryError';
    this.statusCode = statusCode;
    this.code = code;
  }
}
