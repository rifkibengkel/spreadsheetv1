import { Worker, Job } from 'bullmq';
import {
  LARGE_CSV_IMPORT_QUEUE_NAME,
  getLargeCsvRedisOptions,
  DatasetImportJobPayload,
} from '../lib/largeCsv/queue/datasetImportQueue';
import { CsvImportWorker } from '../lib/largeCsv/worker/csvImportWorker';

console.log('[LargeCsvWorkerHost] Starting Dedicated Large CSV Import Worker...');

export const csvImportWorker = new Worker<DatasetImportJobPayload>(
  LARGE_CSV_IMPORT_QUEUE_NAME,
  async (job: Job<DatasetImportJobPayload>) => {
    const { datasetId } = job.data;
    console.log(`[LargeCsvWorkerHost] Picked up job ${job.id} for dataset ${datasetId}`);

    const result = await CsvImportWorker.processDataset(datasetId, (p) => {
      job.updateProgress(Math.round(p.progress * 100));
    });

    console.log(`[LargeCsvWorkerHost] Job ${job.id} finished successfully! Rows: ${result.totalDataRows}, Partitions: ${result.partitionsCount}`);
    return result;
  },
  {
    connection: getLargeCsvRedisOptions(),
    concurrency: parseInt(process.env.IMPORT_WORKER_CONCURRENCY || '1', 10),
    lockDuration: 60000,
    lockRenewTime: 20000,
  }
);

csvImportWorker.on('completed', (job) => {
  console.log(`[LargeCsvWorkerHost] Job ${job.id} completed.`);
});

csvImportWorker.on('failed', (job, err) => {
  console.error(`[LargeCsvWorkerHost] Job ${job?.id} failed with error:`, err.message);
});

export async function stopCsvImportWorker() {
  await csvImportWorker.close();
  console.log('[LargeCsvWorkerHost] Worker stopped.');
}

if (require.main === module) {
  process.on('SIGINT', async () => {
    await stopCsvImportWorker();
    process.exit(0);
  });
  process.on('SIGTERM', async () => {
    await stopCsvImportWorker();
    process.exit(0);
  });
}
