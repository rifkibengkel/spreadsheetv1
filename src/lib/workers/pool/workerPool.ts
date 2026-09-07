import {
  CalculationJob,
  CalculationResult,
  WorkerInMessage,
  WorkerOutMessage,
  WorkerPoolMetrics,
} from './types';

interface PooledWorker {
  id: number;
  worker: Worker;
  busy: boolean;
  currentJob: CalculationJob | null;
  resolve: ((res: CalculationResult) => void) | null;
  reject: ((err: any) => void) | null;
}

interface QueuedJob {
  job: CalculationJob;
  resolve: (res: CalculationResult) => void;
  reject: (err: any) => void;
}

export class WorkerPool {
  private workers: PooledWorker[] = [];
  private jobQueue: QueuedJob[] = [];
  private poolSize: number;
  private isDisposed = false;

  // Metrics tracking
  private totalJobsDispatched = 0;
  private totalJobsCompleted = 0;
  private totalJobsCancelled = 0;
  private totalJobsDiscardedStale = 0;
  private totalDurationMs = 0;

  constructor(poolSize = 4) {
    this.poolSize = poolSize;
    this.initPool();
  }

  private initPool() {
    if (typeof window === 'undefined') return;

    for (let i = 0; i < this.poolSize; i++) {
      this.spawnWorker(i);
    }
  }

  private spawnWorker(id: number) {
    try {
      const worker = new Worker(new URL('../calc.worker.ts', import.meta.url), {
        type: 'module',
      });

      const pooledWorker: PooledWorker = {
        id,
        worker,
        busy: false,
        currentJob: null,
        resolve: null,
        reject: null,
      };

      worker.onmessage = (e: MessageEvent<WorkerOutMessage>) => {
        this.handleWorkerMessage(pooledWorker, e.data);
      };

      worker.onerror = (err) => {
        console.warn(`[WorkerPool] Worker #${id} error:`, err);
        if (pooledWorker.reject) {
          pooledWorker.reject(err);
        }
        this.recoverWorker(pooledWorker);
      };

      this.workers[id] = pooledWorker;
    } catch (e) {
      console.error(`[WorkerPool] Failed to spawn worker #${id}:`, e);
    }
  }

  private recoverWorker(pooledWorker: PooledWorker) {
    if (this.isDisposed) return;
    try {
      pooledWorker.worker.terminate();
    } catch {}

    const id = pooledWorker.id;
    console.log(`[WorkerPool] 🔄 Re-spawning Worker #${id} for self-recovery...`);
    this.spawnWorker(id);
    this.processQueue();
  }

  private handleWorkerMessage(pooled: PooledWorker, data: WorkerOutMessage) {
    if (!data) return;

    const resolve = pooled.resolve;
    const reject = pooled.reject;

    // Reset worker state to idle
    pooled.busy = false;
    pooled.currentJob = null;
    pooled.resolve = null;
    pooled.reject = null;

    if (data.type === 'RESULT') {
      this.totalJobsCompleted++;
      this.totalDurationMs += data.result.durationMs;
      if (resolve) resolve(data.result);
    } else if (data.type === 'CANCELLED') {
      this.totalJobsCancelled++;
      if (reject) reject(new Error('JOB_CANCELLED'));
    } else if (data.type === 'ERROR') {
      if (reject) reject(new Error(data.error));
    }

    // Immediately check if there are pending jobs waiting in queue
    this.processQueue();
  }

  /**
   * Dispatch a calculation job to the pool.
   * Returns a promise for the CalculationResult.
   */
  public dispatch(job: CalculationJob): Promise<CalculationResult> {
    if (this.isDisposed) {
      return Promise.reject(new Error('WorkerPool is disposed'));
    }

    this.totalJobsDispatched++;

    return new Promise<CalculationResult>((resolve, reject) => {
      const queuedJob: QueuedJob = { job, resolve, reject };

      // Find first available idle worker
      const idleWorker = this.workers.find((w) => !w.busy);

      if (idleWorker) {
        this.assignJobToWorker(idleWorker, queuedJob);
      } else {
        // Insert into priority queue: HIGH priority goes ahead of LOW priority
        if (job.priority === 'HIGH') {
          const firstLowIdx = this.jobQueue.findIndex((q) => q.job.priority === 'LOW');
          if (firstLowIdx === -1) {
            this.jobQueue.push(queuedJob);
          } else {
            this.jobQueue.splice(firstLowIdx, 0, queuedJob);
          }
        } else {
          this.jobQueue.push(queuedJob);
        }
      }
    });
  }

  private assignJobToWorker(pooled: PooledWorker, queued: QueuedJob) {
    pooled.busy = true;
    pooled.currentJob = queued.job;
    pooled.resolve = queued.resolve;
    pooled.reject = queued.reject;

    const msg: WorkerInMessage = {
      type: 'CALCULATE',
      job: queued.job,
    };

    pooled.worker.postMessage(msg);
  }

  private processQueue() {
    if (this.jobQueue.length === 0) return;

    const idleWorker = this.workers.find((w) => !w.busy);
    if (!idleWorker) return;

    const nextJob = this.jobQueue.shift();
    if (nextJob) {
      this.assignJobToWorker(idleWorker, nextJob);
    }
  }

  /**
   * Cooperative cancellation for an obsolete generationId.
   * 1. Broadcasts CANCEL_GENERATION to all 4 workers.
   * 2. Prunes all queued jobs belonging to older generations.
   */
  public cancelGeneration(generationId: number) {
    // Broadcast cancellation signal to all active workers
    for (const pooled of this.workers) {
      try {
        pooled.worker.postMessage({
          type: 'CANCEL_GENERATION',
          generationId,
        } as WorkerInMessage);
      } catch {}
    }

    // Filter queued jobs: cancel obsolete generation jobs
    const retainedQueue: QueuedJob[] = [];
    for (const q of this.jobQueue) {
      if (q.job.generationId <= generationId) {
        this.totalJobsCancelled++;
        this.totalJobsDiscardedStale++;
        q.reject(new Error('JOB_DISCARDED_STALE'));
      } else {
        retainedQueue.push(q);
      }
    }
    this.jobQueue = retainedQueue;
  }

  public getMetrics(): WorkerPoolMetrics {
    const activeWorkers = this.workers.filter((w) => w.busy).length;
    const avgDuration =
      this.totalJobsCompleted > 0
        ? Math.round(this.totalDurationMs / this.totalJobsCompleted)
        : 0;

    return {
      totalJobsDispatched: this.totalJobsDispatched,
      totalJobsCompleted: this.totalJobsCompleted,
      totalJobsCancelled: this.totalJobsCancelled,
      totalJobsDiscardedStale: this.totalJobsDiscardedStale,
      averageExecutionTimeMs: avgDuration,
      activeWorkers,
      queueLength: this.jobQueue.length,
    };
  }

  public terminateAll() {
    this.isDisposed = true;
    for (const q of this.jobQueue) {
      q.reject(new Error('WorkerPool disposed'));
    }
    this.jobQueue = [];

    for (const pooled of this.workers) {
      try {
        pooled.worker.terminate();
      } catch {}
    }
    this.workers = [];
  }
}
