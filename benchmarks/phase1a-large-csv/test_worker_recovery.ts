import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import prisma from '../../src/lib/db/prisma';
import { DatasetUploadService } from '../../src/lib/largeCsv/datasetUploadService';
import { CsvImportWorker } from '../../src/lib/largeCsv/worker/csvImportWorker';
import {
  getPartitionBinPath,
  getPartitionIdxPath,
  getDatasetPartitionDir,
  getDatasetSpoolFilePath,
} from '../../src/lib/largeCsv/storage/storagePaths';

async function runRecoveryTests() {
  console.log('='.repeat(80));
  console.log('   PHASE 1A WORKER RECOVERY & PARTITION INTEGRITY TESTS');
  console.log('='.repeat(80));

  let passed = 0;
  let failed = 0;

  function assert(condition: boolean, testName: string) {
    if (condition) {
      console.log(`  [PASS] ${testName}`);
      passed++;
    } else {
      console.error(`  [FAIL] ${testName}`);
      failed++;
    }
  }

  // Create a synthetic CSV with header and 5,000 rows
  const header = 'id,name,amount,timestamp\n';
  const rowLines: string[] = [];
  for (let i = 0; i < 5000; i++) {
    rowLines.push(`${i},Customer_${i},${(i * 1.5).toFixed(2)},2026-09-11T12:00:00Z\n`);
  }
  const fullCsvContent = Buffer.concat([Buffer.from(header, 'utf8'), ...rowLines.map(r => Buffer.from(r, 'utf8'))]);
  const fullChecksum = crypto.createHash('sha256').update(fullCsvContent).digest('hex');

  // Step 1: Create session and upload in 2 chunks
  const session = await DatasetUploadService.createUploadSession({
    originalFilename: 'test_recovery.csv',
    byteSize: fullCsvContent.length,
    expectedChecksum: fullChecksum,
  });

  const half = Math.floor(fullCsvContent.length / 2);
  await DatasetUploadService.uploadChunk({
    datasetId: session.datasetId,
    chunkIndex: 0,
    byteOffset: 0,
    chunkBuffer: fullCsvContent.subarray(0, half),
  });
  await DatasetUploadService.uploadChunk({
    datasetId: session.datasetId,
    chunkIndex: 1,
    byteOffset: half,
    chunkBuffer: fullCsvContent.subarray(half),
  });

  const fin = await DatasetUploadService.finalizeUpload({
    datasetId: session.datasetId,
  });
  assert(fin.status === 'QUEUED', 'Synthetic dataset finalized to QUEUED');

  // Step 2: Run worker on this dataset
  console.log('  [Step] Processing dataset with CsvImportWorker...');
  const result = await CsvImportWorker.processDataset(session.datasetId);
  assert(result.status === 'READY', 'Dataset ingestion reached READY');
  assert(result.totalDataRows === 5000, `Row count matches 5000 (got ${result.totalDataRows})`);

  // Step 3: Validate Partition & Index semantics (.idx uint64 Little-Endian)
  const partitions = await prisma.datasetPartition.findMany({
    where: { datasetId: session.datasetId },
    orderBy: { partitionNumber: 'asc' },
  });
  assert(partitions.length === 1, 'Dataset with 5,000 rows has 1 partition');
  const part0 = partitions[0];
  assert(part0.rowCount === 5000, 'Partition 0 row count is 5,000');

  const idxPath = part0.indexFilePath;
  const binPath = part0.dataFilePath;
  const idxStat = fs.statSync(idxPath);
  const binStat = fs.statSync(binPath);

  // Expected index size: (5000 rows + 1 terminal offset) * 8 = 40,008 bytes
  assert(idxStat.size === (5000 + 1) * 8, `Index file size equals (rowCount + 1) * 8 = 40,008 (got ${idxStat.size})`);

  // Verify random access row retrieval using .idx
  const idxBuf = fs.readFileSync(idxPath);
  const binBuf = fs.readFileSync(binPath);

  // Check row 0
  const row0Start = Number(idxBuf.readBigUInt64LE(0 * 8));
  const row0End = Number(idxBuf.readBigUInt64LE(1 * 8));
  const row0Str = binBuf.subarray(row0Start, row0End).toString('utf8');
  assert(row0Str === rowLines[0], `Row 0 matches source row line: "${row0Str.trim()}"`);

  // Check row 2500
  const row2500Start = Number(idxBuf.readBigUInt64LE(2500 * 8));
  const row2500End = Number(idxBuf.readBigUInt64LE(2501 * 8));
  const row2500Str = binBuf.subarray(row2500Start, row2500End).toString('utf8');
  assert(row2500Str === rowLines[2500], `Row 2500 matches source row line: "${row2500Str.trim()}"`);

  // Check last row (4999)
  const row4999Start = Number(idxBuf.readBigUInt64LE(4999 * 8));
  const row4999End = Number(idxBuf.readBigUInt64LE(5000 * 8));
  const row4999Str = binBuf.subarray(row4999Start, row4999End).toString('utf8');
  assert(row4999Str === rowLines[4999], `Row 4999 matches source row line: "${row4999Str.trim()}"`);
  assert(row4999End === binStat.size, `Terminal offset in .idx equals total .bin file size (${binStat.size})`);

  // Step 4: Simulate crash and worker retry (idempotent recovery)
  console.log('  [Step] Testing restart/retry on already-ready dataset...');
  const retryResult = await CsvImportWorker.processDataset(session.datasetId);
  assert(retryResult.status === 'READY', 'Worker retry on existing partition completes cleanly');
  assert(retryResult.totalDataRows === 5000, 'Worker retry maintains exact row count without duplication');

  // Step 5: Test corruption detection (tampered .idx size)
  console.log('  [Step] Testing corrupted .idx failure detection...');
  // Truncate .idx by 8 bytes
  fs.truncateSync(idxPath, idxStat.size - 8);
  try {
    await CsvImportWorker.verifyDatasetPartitions(session.datasetId, 5000);
    assert(false, 'Corrupted index should be rejected by verification');
  } catch (err: any) {
    assert(err.message.includes('index byte size mismatch') || err.message.includes('reconciliation failed'), 'Corrupted index rejected with error');
  }

  // Clean up synthetic dataset
  await prisma.datasetPartition.deleteMany({ where: { datasetId: session.datasetId } });
  await prisma.uploadChunk.deleteMany({ where: { datasetId: session.datasetId } });
  await prisma.dataset.delete({ where: { id: session.datasetId } });
  try {
    fs.rmSync(path.join(process.cwd(), 'storage', 'datasets', 'partitions', session.datasetId), { recursive: true, force: true });
    fs.rmSync(path.join(process.cwd(), 'storage', 'datasets', 'spool', session.datasetId), { recursive: true, force: true });
  } catch {}

  console.log(`\nWorker Recovery & Partition Tests Complete: ${passed} passed, ${failed} failed.\n`);
  const { closeDatasetImportQueue } = require('../../src/lib/largeCsv/queue/datasetImportQueue');
  await closeDatasetImportQueue();
  await prisma.$disconnect();

  process.exit(failed > 0 ? 1 : 0);
}

runRecoveryTests().catch(e => {
  console.error('Fatal test error:', e);
  process.exit(1);
});
