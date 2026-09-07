import { WorkerPool } from './workerPool';
import { CalculationJob, CalculationJobType, CalculationResult, WorkerPoolMetrics } from './types';

export interface SchedulerCallbacks {
  onCalculationCompleted: (result: CalculationResult) => void;
  onCalculationError?: (error: any, job: CalculationJob) => void;
}

export class CalculationScheduler {
  private pool: WorkerPool;
  private currentGenerationId = 1;
  private activeSubUnitId: string | null = null;
  private callbacks: SchedulerCallbacks;
  private isDisposed = false;

  constructor(callbacks: SchedulerCallbacks, poolSize = 4) {
    this.callbacks = callbacks;
    this.pool = new WorkerPool(poolSize);
  }

  /**
   * Called whenever user switches active worksheet.
   * Instantly increments generationId and cancels all pending/in-flight jobs for older sheets.
   */
  public onSheetActivated(subUnitId: string) {
    if (this.isDisposed || this.activeSubUnitId === subUnitId) return;

    this.activeSubUnitId = subUnitId;
    const oldGen = this.currentGenerationId;
    this.currentGenerationId++;

    // Broadcast cancellation for prior sheet calculations
    this.pool.cancelGeneration(oldGen);
  }

  /**
   * Schedules a formula evaluation on the worker pool.
   * Returns true if offloaded to worker pool, false if left to in-process default.
   */
  public scheduleFormulaCalculation(
    unitId: string,
    subUnitId: string,
    row: number,
    col: number,
    formula: string,
    univerAPI: any
  ): boolean {
    if (this.isDisposed || !formula || !formula.startsWith('=')) return false;

    // Check if this formula is an aggregative / heavy candidate
    const upperFormula = formula.toUpperCase();
    const isSumIf = upperFormula.startsWith('=SUMIF(');
    const isCountIf = upperFormula.startsWith('=COUNTIF(');
    const isSum = upperFormula.startsWith('=SUM(');
    const isAverage = upperFormula.startsWith('=AVERAGE(');

    if (!isSumIf && !isCountIf && !isSum && !isAverage) {
      // Small or complex non-aggregative formula: let in-process engine evaluate
      return false;
    }

    const jobType: CalculationJobType = isSumIf
      ? 'SUMIF'
      : isCountIf
      ? 'COUNTIF'
      : isAverage
      ? 'AVERAGE'
      : 'VECTOR_AGGREGATE';

    try {
      const extractionStart = performance.now();
      const payload = this.extractTargetedVectors(formula, jobType, univerAPI);
      const extractionDuration = performance.now() - extractionStart;

      if (!payload) {
        return false;
      }

      // Generation bump for new user input
      this.currentGenerationId++;
      const currentGen = this.currentGenerationId;
      const jobId = `job_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
      const priority = subUnitId === this.activeSubUnitId ? 'HIGH' : 'LOW';

      const job: CalculationJob = {
        jobId,
        generationId: currentGen,
        unitId,
        subUnitId,
        targetCell: { row, col },
        priority,
        formula,
        type: jobType,
        payload,
        timestamp: Date.now(),
      };

      this.pool
        .dispatch(job)
        .then((result) => {
          // Stale result guard: drop if user has changed formula or switched sheet
          if (result.generationId !== this.currentGenerationId) {
            return;
          }
          this.callbacks.onCalculationCompleted(result);
        })
        .catch((err) => {
          if (err?.message !== 'JOB_CANCELLED' && err?.message !== 'JOB_DISCARDED_STALE') {
            this.callbacks.onCalculationError?.(err, job);
          }
        });

      return true;
    } catch (err) {
      console.warn('[CalculationScheduler] Vector extraction error:', err);
      return false;
    }
  }

  /**
   * Fast targeted range extractor. Extracts only 1D arrays of values from the referenced range,
   * avoiding full workbook cloning or memory overhead.
   */
  private extractTargetedVectors(
    formula: string,
    jobType: CalculationJobType,
    univerAPI: any
  ): any | null {
    const workbook = univerAPI?.getActiveWorkbook?.();
    if (!workbook) return null;

    const inner = formula.substring(formula.indexOf('(') + 1, formula.lastIndexOf(')'));
    const args = this.splitFormulaArguments(inner);

    if (jobType === 'SUMIF') {
      if (args.length < 2) return null;
      const rangeRef = args[0].trim();
      const criteriaStr = args[1].trim();
      const sumRangeRef = args[2] ? args[2].trim() : rangeRef;

      const criteriaVec = this.extract1DVectorFromRef(rangeRef, workbook);
      const sumVec =
        sumRangeRef === rangeRef
          ? criteriaVec
          : this.extract1DVectorFromRef(sumRangeRef, workbook);

      const resolvedCriteria = this.resolveCriteriaValue(criteriaStr, workbook);

      return {
        criteriaVector: criteriaVec,
        sumVector: sumVec,
        targetCriteria: resolvedCriteria,
      };
    }

    if (jobType === 'COUNTIF') {
      if (args.length < 2) return null;
      const rangeRef = args[0].trim();
      const criteriaStr = args[1].trim();

      const criteriaVec = this.extract1DVectorFromRef(rangeRef, workbook);
      const resolvedCriteria = this.resolveCriteriaValue(criteriaStr, workbook);

      return {
        criteriaVector: criteriaVec,
        targetCriteria: resolvedCriteria,
      };
    }

    if (jobType === 'VECTOR_AGGREGATE' || jobType === 'AVERAGE') {
      if (args.length < 1) return null;
      const rangeRef = args[0].trim();
      const rangeVec = this.extract1DVectorFromRef(rangeRef, workbook);

      return {
        rangeVector: rangeVec,
      };
    }

    return null;
  }

  /**
   * Resolves a reference like "'Entries data'!$B$2:$B$626000" or "A1:A100" into a 1D array.
   */
  private extract1DVectorFromRef(ref: string, workbook: any): any[] {
    let sheetName = '';
    let cellRangeStr = ref.replace(/\$/g, '');

    if (cellRangeStr.includes('!')) {
      const parts = cellRangeStr.split('!');
      sheetName = parts[0].replace(/^['"]|['"]$/g, '');
      cellRangeStr = parts[1];
    }

    const worksheet = sheetName
      ? workbook.getSheetByName(sheetName)
      : workbook.getActiveSheet();

    if (!worksheet) return [];

    const matrix =
      (worksheet as any).getSheet?.()?.getCellMatrix?.() ||
      (worksheet as any).getCellMatrix?.();

    if (!matrix) return [];

    const bounds = this.parseRangeCoordinates(cellRangeStr);
    if (!bounds) return [];

    const { startRow, endRow, startCol, endCol } = bounds;
    const rowCount = endRow - startRow + 1;
    const colCount = endCol - startCol + 1;
    const totalCount = Math.max(0, rowCount * colCount);
    const vector = new Array(totalCount);
    let idx = 0;

    // For single-column multi-row references (e.g. B2:B626000):
    if (startCol === endCol) {
      for (let r = startRow; r <= endRow; r++) {
        const val = matrix.getValue(r, startCol);
        vector[idx++] = val?.v ?? null;
      }
    } else {
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const val = matrix.getValue(r, c);
          vector[idx++] = val?.v ?? null;
        }
      }
    }

    return vector;
  }

  /**
   * Parses A1-style reference like "B2:B626000" or "A1" into 0-indexed coordinates.
   */
  private parseRangeCoordinates(
    rangeStr: string
  ): { startRow: number; endRow: number; startCol: number; endCol: number } | null {
    const parts = rangeStr.split(':');
    const startCoord = this.parseSingleCell(parts[0]);
    if (!startCoord) return null;

    if (parts.length === 1) {
      return {
        startRow: startCoord.row,
        endRow: startCoord.row,
        startCol: startCoord.col,
        endCol: startCoord.col,
      };
    }

    const endCoord = this.parseSingleCell(parts[1]);
    if (!endCoord) return null;

    return {
      startRow: Math.min(startCoord.row, endCoord.row),
      endRow: Math.max(startCoord.row, endCoord.row),
      startCol: Math.min(startCoord.col, endCoord.col),
      endCol: Math.max(startCoord.col, endCoord.col),
    };
  }

  private parseSingleCell(cellStr: string): { row: number; col: number } | null {
    const match = cellStr.trim().match(/^([A-Za-z]+)(\d+)$/);
    if (!match) return null;

    const colLetters = match[1].toUpperCase();
    const rowNum = parseInt(match[2], 10) - 1; // 0-indexed

    let colNum = 0;
    for (let i = 0; i < colLetters.length; i++) {
      colNum = colNum * 26 + (colLetters.charCodeAt(i) - 64);
    }
    colNum -= 1; // 0-indexed

    return { row: rowNum, col: colNum };
  }

  private resolveCriteriaValue(criteriaArg: string, workbook: any): any {
    const trimmed = criteriaArg.trim();

    // String literal in quotes
    if (
      (trimmed.startsWith('"') && trimmed.endsWith('"')) ||
      (trimmed.startsWith("'") && trimmed.endsWith("'"))
    ) {
      return trimmed.slice(1, -1);
    }

    // Direct number
    const num = Number(trimmed);
    if (!isNaN(num)) return num;

    // Cell reference like "Summary!$B11" or "B11"
    const parsed = this.parseRangeCoordinates(trimmed.replace(/[$'"]/g, ''));
    if (parsed) {
      let sheetName = '';
      let cellStr = trimmed;
      if (trimmed.includes('!')) {
        const parts = trimmed.split('!');
        sheetName = parts[0].replace(/^['"]|['"]$/g, '');
        cellStr = parts[1];
      }
      const ws = sheetName ? workbook.getSheetByName(sheetName) : workbook.getActiveSheet();
      if (ws) {
        const matrix = ws.getSheet?.()?.getCellMatrix?.() || ws.getCellMatrix?.();
        const coord = this.parseSingleCell(cellStr.replace(/\$/g, ''));
        if (coord && matrix) {
          return matrix.getValue(coord.row, coord.col)?.v ?? trimmed;
        }
      }
    }

    return trimmed;
  }

  private splitFormulaArguments(argsStr: string): string[] {
    const args: string[] = [];
    let current = '';
    let inQuotes = false;
    let quoteChar = '';
    let depth = 0;

    for (let i = 0; i < argsStr.length; i++) {
      const char = argsStr[i];

      if ((char === '"' || char === "'") && (i === 0 || argsStr[i - 1] !== '\\')) {
        if (!inQuotes) {
          inQuotes = true;
          quoteChar = char;
        } else if (quoteChar === char) {
          inQuotes = false;
        }
      }

      if (!inQuotes) {
        if (char === '(') depth++;
        if (char === ')') depth--;
        if (char === ',' && depth === 0) {
          args.push(current.trim());
          current = '';
          continue;
        }
      }
      current += char;
    }

    if (current.trim()) {
      args.push(current.trim());
    }

    return args;
  }

  public getMetrics(): WorkerPoolMetrics {
    return this.pool.getMetrics();
  }

  public dispose() {
    this.isDisposed = true;
    this.pool.terminateAll();
  }
}
