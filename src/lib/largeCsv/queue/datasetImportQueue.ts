import { Queue, QueueOptions } from 'bullmq';
import Redis from 'ioredis';

const DEFAULT_REDIS_URL = process.env.REDIS_URL || 'redis://127.0.0.1:6381';

export function createLargeCsvRedisConnection(): Redis {
  return new Redis(DEFAULT_REDIS_URL, {
    maxRetriesPerRequest: null,
    enableReadyCheck: false,
    lazyConnect: false,
    retryStrategy(times) {
      return Math.min(times * 100, 3000);
    },
  });
}

export function getLargeCsvRedisOptions() {
  const DEFAULT_REDIS_URL = process.env.REDIS_URL || 'redis://127.0.0.1:6380';
  try {
    const url = new URL(DEFAULT_REDIS_URL);
    return {
      host: url.hostname || '127.0.0.1',
      port: parseInt(url.port || '6380', 10),
      maxRetriesPerRequest: null,
    };
  } catch {
    return {
      host: '127.0.0.1',
      port: 6380,
      maxRetriesPerRequest: null,
    };
  }
}

export const LARGE_CSV_IMPORT_QUEUE_NAME = 'large-csv-import-queue';

export interface DatasetImportJobPayload {
  datasetId: string;
  enqueuedAt: string;
  attempt?: number;
}

let importQueue: Queue<DatasetImportJobPayload> | null = null;

export function getDatasetImportQueue(): Queue<DatasetImportJobPayload> {
  if (importQueue) return importQueue;

  const queueOptions: QueueOptions = {
    connection: getLargeCsvRedisOptions(),
    defaultJobOptions: {
      attempts: 3,
      backoff: {
        type: 'exponential',
        delay: 2000,
      },
      removeOnComplete: {
        count: 50,
        age: 86400,
      },
      removeOnFail: {
        count: 100,
      },
    },
  };

  importQueue = new Queue<DatasetImportJobPayload>(LARGE_CSV_IMPORT_QUEUE_NAME, queueOptions);
  return importQueue;
}

export async function enqueueDatasetImportJob(datasetId: string): Promise<string> {
  const queue = getDatasetImportQueue();
  const jobId = `import_ds_${datasetId}`;
  
  const job = await queue.add(
    'process-dataset',
    {
      datasetId,
      enqueuedAt: new Date().toISOString(),
    },
    {
      jobId, // deterministic job ID prevents duplicates
    }
  );

  return job.id!;
}

export async function closeDatasetImportQueue(): Promise<void> {
  if (importQueue) {
    await importQueue.close();
    importQueue = null;
  }
}
