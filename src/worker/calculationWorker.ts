import { prisma } from '../lib/db/prisma';
import Redis from 'ioredis';
import { WorkbookEngine } from './engine/dependencyGraph';
import { normalizeWorkbookUuid } from '../lib/uuid/normalizeUuid';

// ==============================================================================
// DEDICATED CALCULATION WORKER (PHASE 4 IMPLEMENTATION)
// Runs backend calculation off the browser UI thread
// ==============================================================================

const REDIS_URL = process.env.REDIS_URL || 'redis://127.0.0.1:6379';

export class CalculationWorker {
  private redisPub: Redis;
  private redisSub: Redis;
  private workbookEngines: Map<string, WorkbookEngine> = new Map();
  private isRunning = false;

  constructor() {
    this.redisPub = new Redis(REDIS_URL, {
      maxRetriesPerRequest: 1,
      enableOfflineQueue: false,
      lazyConnect: true,
      connectTimeout: 2000,
    });
    this.redisSub = new Redis(REDIS_URL, {
      maxRetriesPerRequest: 1,
      enableOfflineQueue: false,
      lazyConnect: true,
      connectTimeout: 2000,
    });
    this.redisPub.on('error', (err) => console.warn('[WORKER_HOST] redisPub warning:', err.message));
    this.redisSub.on('error', (err) => console.warn('[WORKER_HOST] redisSub warning:', err.message));
  }

  public async start(): Promise<void> {
    if (this.isRunning) return;
    this.isRunning = true;

    console.log('[WORKER_HOST] Starting Dedicated Calculation Worker...');

    // Subscribe to calculation trigger channel
    await this.redisSub.psubscribe('wb:*:channel:calc_trigger');

    this.redisSub.on('pmessage', async (_pattern, channel, message) => {
      try {
        const match = channel.match(/^wb:(.*):channel:calc_trigger$/);
        if (!match) return;

        const workbookId = match[1];
        const data = JSON.parse(message);
        await this.processWorkbookRevision(workbookId, data.revisionId);
      } catch (err: any) {
        console.error('[WORKER_HOST] Error processing trigger message:', err.message);
      }
    });

    console.log('[WORKER_HOST] Dedicated Calculation Worker is LIVE and listening for calculation triggers');
  }

  /**
   * Processes a specific revision for a workbook from PostgreSQL WAL.
   */
  public async processWorkbookRevision(workbookId: string, targetRev?: number): Promise<any> {
    const startTime = performance.now();

    // 1. Fetch workbook state
    let state = await prisma.workbookState.findUnique({
      where: { workbookId },
    });

    if (!state) {
      state = await prisma.workbookState.create({
        data: {
          workbookId,
          latestPersistedRev: BigInt(targetRev || 1),
          latestCalculatedRev: BigInt(0),
          fencingToken: BigInt(1),
        },
      });
    }

    const currentCalcRev = Number(state.latestCalculatedRev);
    const persistedRev = Number(state.latestPersistedRev);

    if (currentCalcRev >= persistedRev && !targetRev) {
      // Nothing to calculate
      return { status: 'UP_TO_DATE', calculatedRev: currentCalcRev };
    }

    // 2. Hydrate or retrieve engine
    let engine = this.workbookEngines.get(workbookId);
    if (!engine) {
      engine = new WorkbookEngine(workbookId);
      this.workbookEngines.set(workbookId, engine);
    }

    // 3. Fetch missing mutations from PostgreSQL WAL sequentially
    const missingBatches = await prisma.mutationBatch.findMany({
      where: {
        workbookId,
        revisionId: {
          gt: BigInt(engine.revision),
          lte: BigInt(persistedRev),
        },
      },
      orderBy: {
        revisionId: 'asc',
      },
    });

    // 4. Sequentially replay mutation deltas into memory
    for (const batch of missingBatches) {
      const deltas = batch.deltas as any[];
      if (Array.isArray(deltas)) {
        for (const delta of deltas) {
          if (delta.sheetId) {
            if (delta.f) {
              engine.setFormula(delta.sheetId, delta.row, delta.col, delta.f);
            }
            if (delta.v !== undefined) {
              engine.setCellValue(delta.sheetId, delta.row, delta.col, delta.v);
            }
          }
        }
      }
      engine.revision = Number(batch.revisionId);
    }

    // 5. Evaluate formulas
    const patches = engine.evaluateAll();
    const durationMs = Math.round(performance.now() - startTime);

    // 6. Commit calculation result to PostgreSQL transactionally
    await prisma.$transaction(async (tx) => {
      // Update calculated revision
      await tx.workbookState.update({
        where: { workbookId },
        data: {
          latestCalculatedRev: BigInt(persistedRev),
          updatedAt: new Date(),
        },
      });

      // Insert calculation result
      await tx.calculationResult.upsert({
        where: {
          workbookId_calculatedRev: {
            workbookId,
            calculatedRev: BigInt(persistedRev),
          },
        },
        create: {
          workbookId,
          calculatedRev: BigInt(persistedRev),
          cellPatches: patches,
          executionMs: durationMs,
          plannerStrategy: 'TIER1_VECTOR_KERNEL',
        },
        update: {
          cellPatches: patches,
          executionMs: durationMs,
        },
      });
    });

    // 7. Post-commit publication to Redis Pub/Sub
    try {
      const patchPayload = JSON.stringify({
        workbookId,
        revId: persistedRev,
        patches,
        executionMs: durationMs,
      });

      // Store in transient cache
      await this.redisPub.set(`wb:${workbookId}:latest_patch`, patchPayload, 'EX', 3600);
      // Broadcast on pub/sub channel
      await this.redisPub.publish(`wb:${workbookId}:channel:patches`, patchPayload);
    } catch (redisErr: any) {
      console.warn('[WORKER_HOST] Redis publish notice:', redisErr.message);
    }

    console.log(`[WORKER_HOST] Rev ${persistedRev} computed in ${durationMs}ms with ${patches.length} patches`);

    return {
      status: 'SUCCESS',
      workbookId,
      calculatedRev: persistedRev,
      patchCount: patches.length,
      executionMs: durationMs,
      patches,
    };
  }

  public async stop(): Promise<void> {
    this.isRunning = false;
    await this.redisSub.punsubscribe();
    this.redisSub.disconnect();
    this.redisPub.disconnect();
  }
}

// Export singleton instance for immediate in-process calculation or standalone execution
export const calculationWorker = new CalculationWorker();
