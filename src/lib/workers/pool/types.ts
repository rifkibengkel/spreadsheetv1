export type JobPriority = 'HIGH' | 'LOW';

export type CalculationJobType = 
  | 'EVALUATE_FORMULA' 
  | 'VECTOR_AGGREGATE' 
  | 'SUMIF' 
  | 'COUNTIF' 
  | 'AVERAGE';

export interface CalculationJob {
  jobId: string;
  generationId: number;
  unitId: string;
  subUnitId: string;
  targetCell: { row: number; col: number };
  priority: JobPriority;
  formula: string;
  type: CalculationJobType;
  payload: {
    criteriaVector?: any[];
    sumVector?: (number | null | undefined)[];
    targetCriteria?: any;
    rangeVector?: (number | null | undefined)[];
    expression?: string;
    [key: string]: any;
  };
  timestamp: number;
}

export interface CalculationResult {
  jobId: string;
  generationId: number;
  unitId: string;
  subUnitId: string;
  targetCell: { row: number; col: number };
  value: any;
  error?: string;
  durationMs: number;
}

export type WorkerInMessage =
  | { type: 'CALCULATE'; job: CalculationJob }
  | { type: 'CANCEL_GENERATION'; generationId: number }
  | { type: 'PING'; id: number };

export type WorkerOutMessage =
  | { type: 'RESULT'; result: CalculationResult }
  | { type: 'ERROR'; jobId: string; generationId: number; error: string; durationMs: number }
  | { type: 'CANCELLED'; jobId: string; generationId: number }
  | { type: 'PONG'; id: number };

export interface WorkerPoolMetrics {
  totalJobsDispatched: number;
  totalJobsCompleted: number;
  totalJobsCancelled: number;
  totalJobsDiscardedStale: number;
  averageExecutionTimeMs: number;
  activeWorkers: number;
  queueLength: number;
}
