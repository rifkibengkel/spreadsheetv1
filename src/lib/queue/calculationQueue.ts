import { Queue, QueueOptions } from 'bullmq';
import Redis from 'ioredis';
import {
  CalculationJobPayload,
  QUEUE_PARTITIONS_COUNT,
  getQueueNameForPartition,
} from './queueTypes';

// ==============================================================================
// REDIS CONNECTION FACTORY FOR BULLMQ
// BullMQ requires maxRetriesPerRequest: null and isolated connections
// ==============================================================================

const REDIS_URL = process.env.REDIS_URL || 'redis://127.0.0.1:6379';

export function createRedisConnection(): Redis {
  return new Redis(REDIS_URL, {
    maxRetriesPerRequest: null,
    enableReadyCheck: false,
    lazyConnect: false,
    retryStrategy(times) {
      return Math.min(times * 100, 3000);
    },
  });
}

// Shared Redis client for metadata coordination
export const coordRedis = createRedisConnection();

// ==============================================================================
// DETERMINISTIC CRC32 PARTITIONING
// Maps workbookId -> Partition Index [0 .. QUEUE_PARTITIONS_COUNT - 1]
// ==============================================================================

export function calculateCRC32(str: string): number {
  let crc = -1;
  for (let i = 0; i < str.length; i++) {
    let byte = str.charCodeAt(i);
    for (let j = 0; j < 8; j++) {
      const mask = -(byte & 1);
      byte = (byte >>> 1) ^ (0xEDB88320 & mask);
    }
    crc = (crc >>> 8) ^ byte;
  }
  return (crc ^ -1) >>> 0;
}

export function getPartitionForWorkbook(workbookId: string): number {
  return calculateCRC32(workbookId) % QUEUE_PARTITIONS_COUNT;
}

// ==============================================================================
// REDIS COORDINATION KEY HELPERS
// ==============================================================================

export const RedisKeys = {
  latestEnqueuedRev: (wbId: string) => `wb:${wbId}:latest_enqueued_rev`,
  workerLease: (wbId: string) => `wb:${wbId}:worker_lease`,
  patchChannel: (wbId: string) => `wb:${wbId}:channel:patches`,
  abortChannel: (wbId: string) => `wb:${wbId}:channel:abort`,
  patchCache: (wbId: string, rev: number | string) => `wb:${wbId}:patch_cache:${rev}`,
};

// ==============================================================================
// BULLMQ QUEUE REGISTRY
// ==============================================================================

const queueMap = new Map<number, Queue<CalculationJobPayload>>();

export function getCalculationQueue(partitionIndex: number): Queue<CalculationJobPayload> {
  if (queueMap.has(partitionIndex)) {
    return queueMap.get(partitionIndex)!;
  }

  const queueName = getQueueNameForPartition(partitionIndex);
  const queueOptions: QueueOptions = {
    connection: createRedisConnection() as any,
    defaultJobOptions: {
      attempts: 3,
      backoff: {
        type: 'exponential',
        delay: 1000,
      },
      removeOnComplete: {
        count: 100,
        age: 3600,
      },
      removeOnFail: {
        count: 500,
      },
    },
  };

  const queue = new Queue<CalculationJobPayload>(queueName, queueOptions);
  queueMap.set(partitionIndex, queue);
  return queue;
}

/**
 * Enqueues a calculation job with monotonic revision tracking.
 * Strictly enforces unique job ID per revision: wb_{workbookId}_rev_{revisionId}
 */
export async function enqueueCalculationJob(payload: CalculationJobPayload): Promise<string> {
  const { workbookId, revisionId } = payload;
  const partition = getPartitionForWorkbook(workbookId);
  const queue = getCalculationQueue(partition);

  // 1. Update latest enqueued revision in Redis for pre-execution supersession checks
  const latestRevKey = RedisKeys.latestEnqueuedRev(workbookId);
  await coordRedis.set(latestRevKey, revisionId.toString(), 'EX', 86400); // 24-hour TTL

  // 2. Publish abort signal to cancel any active in-flight worker thread computing an older revision
  const abortKey = RedisKeys.abortChannel(workbookId);
  await coordRedis.publish(abortKey, JSON.stringify({ targetRev: revisionId }));

  // 3. Enqueue job into deterministic partition queue with unique revision Job ID
  const jobId = `wb_${workbookId}_rev_${revisionId}`;
  await queue.add('calculate', payload, {
    jobId,
    priority: 1, // Higher priority for newest user edit
  });

  return jobId;
}
