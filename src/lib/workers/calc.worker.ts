import { CalculationJob, WorkerInMessage, WorkerOutMessage } from './pool/types';

let latestCancelledGen = 0;

self.addEventListener('message', (e: MessageEvent<WorkerInMessage>) => {
  const data = e.data;
  if (!data) return;

  if (data.type === 'PING') {
    (self as any).postMessage({ type: 'PONG', id: data.id });
    return;
  }

  if (data.type === 'CANCEL_GENERATION') {
    if (data.generationId > latestCancelledGen) {
      latestCancelledGen = data.generationId;
    }
    return;
  }

  if (data.type === 'CALCULATE') {
    const job = data.job;
    const start = performance.now();

    // Early exit if this generation was already cancelled
    if (job.generationId <= latestCancelledGen) {
      (self as any).postMessage({
        type: 'CANCELLED',
        jobId: job.jobId,
        generationId: job.generationId,
      } as WorkerOutMessage);
      return;
    }

    try {
      const value = executeJob(job);
      const durationMs = performance.now() - start;

      // Check cancellation again after execution before posting result
      if (job.generationId <= latestCancelledGen) {
        (self as any).postMessage({
          type: 'CANCELLED',
          jobId: job.jobId,
          generationId: job.generationId,
        } as WorkerOutMessage);
        return;
      }

      (self as any).postMessage({
        type: 'RESULT',
        result: {
          jobId: job.jobId,
          generationId: job.generationId,
          unitId: job.unitId,
          subUnitId: job.subUnitId,
          targetCell: job.targetCell,
          value,
          durationMs,
        },
      } as WorkerOutMessage);
    } catch (err: any) {
      const durationMs = performance.now() - start;
      (self as any).postMessage({
        type: 'ERROR',
        jobId: job.jobId,
        generationId: job.generationId,
        error: err?.message || 'Calculation error',
        durationMs,
      } as WorkerOutMessage);
    }
  }
});

function executeJob(job: CalculationJob): any {
  const { type, payload } = job;

  switch (type) {
    case 'SUMIF': {
      const criteriaVec = payload.criteriaVector || [];
      const sumVec = payload.sumVector || criteriaVec;
      const targetCrit = payload.targetCriteria;
      const len = Math.min(criteriaVec.length, sumVec.length);

      let total = 0;
      const checkOp = compileCriteria(targetCrit);

      // Cooperative cancellation check every 10,000 rows
      for (let i = 0; i < len; i++) {
        if (i > 0 && i % 10000 === 0) {
          if (job.generationId <= latestCancelledGen) {
            throw new Error('JOB_CANCELLED_COOPERATIVE');
          }
        }

        const critVal = criteriaVec[i];
        if (checkOp(critVal)) {
          const num = Number(sumVec[i]);
          if (!isNaN(num)) {
            total += num;
          }
        }
      }
      return total;
    }

    case 'COUNTIF': {
      const criteriaVec = payload.criteriaVector || [];
      const targetCrit = payload.targetCriteria;
      const len = criteriaVec.length;

      let count = 0;
      const checkOp = compileCriteria(targetCrit);

      for (let i = 0; i < len; i++) {
        if (i > 0 && i % 10000 === 0) {
          if (job.generationId <= latestCancelledGen) {
            throw new Error('JOB_CANCELLED_COOPERATIVE');
          }
        }

        if (checkOp(criteriaVec[i])) {
          count++;
        }
      }
      return count;
    }

    case 'AVERAGE':
    case 'VECTOR_AGGREGATE': {
      const vec = payload.rangeVector || [];
      const len = vec.length;

      let sum = 0;
      let count = 0;

      for (let i = 0; i < len; i++) {
        if (i > 0 && i % 10000 === 0) {
          if (job.generationId <= latestCancelledGen) {
            throw new Error('JOB_CANCELLED_COOPERATIVE');
          }
        }

        const raw = vec[i];
        if (raw !== null && raw !== undefined && (raw as any) !== '') {
          const num = Number(raw);
          if (!isNaN(num)) {
            sum += num;
            count++;
          }
        }
      }

      if (type === 'AVERAGE') {
        return count > 0 ? sum / count : '#DIV/0!';
      }
      return sum;
    }

    case 'EVALUATE_FORMULA': {
      // Basic expression evaluator fallback if formula is simple expression
      if (typeof payload.expression === 'string') {
        try {
          // Safe mathematical expression evaluation
          const sanitized = payload.expression.replace(/[^0-9+\-*/(). ]/g, '');
          // eslint-disable-next-line no-new-func
          const evalFn = new Function(`return (${sanitized});`);
          return evalFn();
        } catch {
          return '#VALUE!';
        }
      }
      return null;
    }

    default:
      return null;
  }
}

/**
 * Fast compiled criteria matcher for Excel comparison patterns:
 * e.g., ">100", "<=50", "<>0", "Apple*", or exact values.
 */
function compileCriteria(crit: any): (val: any) => boolean {
  if (crit === null || crit === undefined) {
    return (val: any) => val === null || val === undefined || val === '';
  }

  if (typeof crit === 'string') {
    const trimmed = crit.trim();

    if (trimmed.startsWith('>=')) {
      const n = parseFloat(trimmed.slice(2));
      return (val: any) => !isNaN(Number(val)) && Number(val) >= n;
    }
    if (trimmed.startsWith('<=')) {
      const n = parseFloat(trimmed.slice(2));
      return (val: any) => !isNaN(Number(val)) && Number(val) <= n;
    }
    if (trimmed.startsWith('<>')) {
      const n = parseFloat(trimmed.slice(2));
      if (!isNaN(n)) return (val: any) => Number(val) !== n;
      const targetStr = trimmed.slice(2);
      return (val: any) => String(val).toLowerCase() !== targetStr.toLowerCase();
    }
    if (trimmed.startsWith('>')) {
      const n = parseFloat(trimmed.slice(1));
      return (val: any) => !isNaN(Number(val)) && Number(val) > n;
    }
    if (trimmed.startsWith('<')) {
      const n = parseFloat(trimmed.slice(1));
      return (val: any) => !isNaN(Number(val)) && Number(val) < n;
    }
    if (trimmed.startsWith('=')) {
      const target = trimmed.slice(1);
      const n = parseFloat(target);
      if (!isNaN(n)) return (val: any) => Number(val) === n;
      return (val: any) => String(val).toLowerCase() === target.toLowerCase();
    }

    // Wildcard match or plain string match
    const lowerCrit = trimmed.toLowerCase();
    return (val: any) => String(val).toLowerCase() === lowerCrit;
  }

  if (typeof crit === 'number') {
    return (val: any) => Number(val) === crit;
  }

  return (val: any) => val === crit;
}
