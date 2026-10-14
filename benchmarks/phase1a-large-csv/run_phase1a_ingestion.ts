import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import prisma from '../../src/lib/db/prisma';
import { DatasetUploadService } from '../../src/lib/largeCsv/datasetUploadService';
import { CsvImportWorker, PARTITION_ROW_LIMIT } from '../../src/lib/largeCsv/worker/csvImportWorker';
import { closeDatasetImportQueue } from '../../src/lib/largeCsv/queue/datasetImportQueue';
import {
  getPartitionBinPath,
  getPartitionIdxPath,
} from '../../src/lib/largeCsv/storage/storagePaths';

const FIXTURE_PATH = 'C:\\Users\\BKKI-1\\Downloads\\archive\\parking_violation.csv';
const EXPECTED_SIZE = 3_121_309_240;
const EXPECTED_DATA_ROWS = 11_809_233;
const EXPECTED_COLUMNS = 54;
const CHUNK_SIZE = 64 * 1024 * 1024; // 64 MB chunk size

interface ExecutionMetrics {
  uploadDurationSec: number;
  uploadSpeedMBps: number;
  importDurationSec: number;
  importRowsPerSec: number;
  importSpeedMBps: number;
  totalDurationSec: number;
  peakRssMB: number;
  peakHeapUsedMB: number;
  peakHeapTotalMB: number;
  totalDataRows: number;
  columnCount: number;
  partitionsCount: number;
  partitions: Array<{
    partitionNumber: number;
    stableRowStart: number;
    stableRowEnd: number;
    rowCount: number;
    byteSize: number;
    idxSize: number;
    checksum: string;
  }>;
  boundaryTestsPassed: boolean;
  randomAccessTestsPassed: boolean;
  reconciliationPassed: boolean;
}

async function main() {
  console.log('='.repeat(80));
  console.log('   PHASE 1A REAL-FIXTURE INGESTION ENGINE RUNNER');
  console.log('   Fixture: parking_violation.csv (3.12 GB, 11.8M rows, 54 columns)');
  console.log('='.repeat(80));

  if (!fs.existsSync(FIXTURE_PATH)) {
    throw new Error(`Fixture not found at ${FIXTURE_PATH}`);
  }
  const stat = fs.statSync(FIXTURE_PATH);
  if (stat.size !== EXPECTED_SIZE) {
    throw new Error(`Fixture size mismatch: expected ${EXPECTED_SIZE}, got ${stat.size}`);
  }
  console.log(`[Fixture] Size verified: ${stat.size.toLocaleString()} bytes.`);

  // Setup peak memory monitor
  let peakRss = 0;
  let peakHeapUsed = 0;
  let peakHeapTotal = 0;

  const memInterval = setInterval(() => {
    const mem = process.memoryUsage();
    if (mem.rss > peakRss) peakRss = mem.rss;
    if (mem.heapUsed > peakHeapUsed) peakHeapUsed = mem.heapUsed;
    if (mem.heapTotal > peakHeapTotal) peakHeapTotal = mem.heapTotal;
  }, 250);

  const overallStartTime = Date.now();

  // -------------------------------------------------------------------------
  // STAGE 1: RESUMABLE CHUNK UPLOAD
  // -------------------------------------------------------------------------
  console.log('\n--- STAGE 1: Resumable Chunked Upload ---');
  const uploadStartTime = Date.now();

  const session = await DatasetUploadService.createUploadSession({
    name: 'Parking Violations Production Fixture',
    originalFilename: 'parking_violation.csv',
    byteSize: stat.size,
  });
  console.log(`[Upload] Session created: datasetId=${session.datasetId}`);

  const fixtureFd = fs.openSync(FIXTURE_PATH, 'r');
  const chunkBuffer = Buffer.alloc(CHUNK_SIZE);
  let chunkIndex = 0;
  let currentOffset = 0;
  let bytesRemaining = stat.size;

  while (bytesRemaining > 0) {
    const toRead = Math.min(CHUNK_SIZE, bytesRemaining);
    const bytesRead = fs.readSync(fixtureFd, chunkBuffer, 0, toRead, currentOffset);
    if (bytesRead <= 0) break;

    const slice = chunkBuffer.subarray(0, bytesRead);
    await DatasetUploadService.uploadChunk({
      datasetId: session.datasetId,
      chunkIndex,
      byteOffset: currentOffset,
      chunkBuffer: slice,
    });

    currentOffset += bytesRead;
    bytesRemaining -= bytesRead;
    chunkIndex++;

    if (chunkIndex % 10 === 0 || bytesRemaining === 0) {
      const pct = ((currentOffset / stat.size) * 100).toFixed(1);
      console.log(`[Upload] Chunks: ${chunkIndex} | Uploaded: ${(currentOffset / (1024 * 1024)).toFixed(1)} MB / ${(stat.size / (1024 * 1024)).toFixed(1)} MB (${pct}%)`);
    }
  }
  fs.closeSync(fixtureFd);

  console.log(`[Upload] All ${chunkIndex} chunks written. Finalizing upload session...`);
  const finalized = await DatasetUploadService.finalizeUpload({
    datasetId: session.datasetId,
  });
  const uploadDurationSec = (Date.now() - uploadStartTime) / 1000;
  const uploadSpeedMBps = (stat.size / (1024 * 1024)) / uploadDurationSec;
  console.log(`[Upload] Finalized in ${uploadDurationSec.toFixed(2)}s (${uploadSpeedMBps.toFixed(2)} MB/s). Status: ${finalized.status}, SHA-256: ${finalized.actualChecksum}`);

  // -------------------------------------------------------------------------
  // STAGE 2: STREAMING PARTITIONING & INDEXING WORKER
  // -------------------------------------------------------------------------
  console.log('\n--- STAGE 2: Dedicated Streaming Ingestion Worker ---');
  const importStartTime = Date.now();

  let lastReportedMillion = 0;
  const importResult = await CsvImportWorker.processDataset(session.datasetId, (p) => {
    const currentMillion = Math.floor(p.rowsProcessed / 1_000_000);
    if (currentMillion > lastReportedMillion) {
      lastReportedMillion = currentMillion;
      const elapsed = (Date.now() - importStartTime) / 1000;
      const rowsPerSec = Math.round(p.rowsProcessed / Math.max(0.1, elapsed));
      console.log(`[Worker] Rows: ${p.rowsProcessed.toLocaleString()} (${rowsPerSec.toLocaleString()} rows/s) | Partition: ${p.currentPartition} | Progress: ${(p.progress * 100).toFixed(1)}% | RSS: ${(process.memoryUsage().rss / (1024 * 1024)).toFixed(1)} MB`);
    }
  });

  const importDurationSec = (Date.now() - importStartTime) / 1000;
  const importRowsPerSec = Math.round(importResult.totalDataRows / Math.max(0.1, importDurationSec));
  const importSpeedMBps = (stat.size / (1024 * 1024)) / importDurationSec;

  console.log(`\n[Worker] Ingestion complete in ${importDurationSec.toFixed(2)}s!`);
  console.log(`[Worker] Throughput: ${importRowsPerSec.toLocaleString()} rows/s (${importSpeedMBps.toFixed(2)} MB/s)`);
  console.log(`[Worker] Total Data Rows: ${importResult.totalDataRows.toLocaleString()}, Partitions: ${importResult.partitionsCount}`);

  // -------------------------------------------------------------------------
  // STAGE 3: POSTGRESQL & PARTITION METADATA RECONCILIATION
  // -------------------------------------------------------------------------
  console.log('\n--- STAGE 3: Metadata Reconciliation & Verification ---');
  const dataset = await prisma.dataset.findUnique({
    where: { id: session.datasetId },
    include: {
      partitions: { orderBy: { partitionNumber: 'asc' } },
    },
  });

  if (!dataset) throw new Error('Dataset record missing from database');

  console.log(`[DB] Dataset Status: ${dataset.status}`);
  console.log(`[DB] Dataset Row Count: ${dataset.rowCount?.toString()}`);
  console.log(`[DB] Dataset Column Count: ${dataset.columnCount}`);
  console.log(`[DB] Partitions in DB: ${dataset.partitions.length}`);

  let sumPartitionRows = 0;
  const partitionsData = [];

  for (const part of dataset.partitions) {
    const binStat = fs.statSync(part.dataFilePath);
    const idxStat = fs.statSync(part.indexFilePath);
    const expectedIdxBytes = (part.rowCount + 1) * 8;

    console.log(`  Partition ${part.partitionNumber}: rows [${part.stableRowStart}..${part.stableRowEnd}] (${part.rowCount.toLocaleString()} rows) | .bin: ${(binStat.size / (1024 * 1024)).toFixed(2)} MB | .idx: ${(idxStat.size / (1024 * 1024)).toFixed(2)} MB | status: ${part.status}`);

    if (idxStat.size !== expectedIdxBytes) {
      throw new Error(`Partition ${part.partitionNumber} index size mismatch: expected ${expectedIdxBytes}, got ${idxStat.size}`);
    }
    if (BigInt(binStat.size) !== part.byteSize) {
      throw new Error(`Partition ${part.partitionNumber} bin size mismatch: DB ${part.byteSize}, disk ${binStat.size}`);
    }

    sumPartitionRows += part.rowCount;
    partitionsData.push({
      partitionNumber: part.partitionNumber,
      stableRowStart: Number(part.stableRowStart),
      stableRowEnd: Number(part.stableRowEnd),
      rowCount: part.rowCount,
      byteSize: Number(part.byteSize),
      idxSize: idxStat.size,
      checksum: part.checksum || '',
    });
  }

  const reconciliationPassed =
    sumPartitionRows === EXPECTED_DATA_ROWS &&
    Number(dataset.rowCount) === EXPECTED_DATA_ROWS &&
    dataset.status === 'READY';

  console.log(`\n[Invariant] SUM(partition.rowCount) = ${sumPartitionRows.toLocaleString()}`);
  console.log(`[Invariant] dataset.rowCount        = ${dataset.rowCount?.toString()}`);
  console.log(`[Invariant] Expected Rows           = ${EXPECTED_DATA_ROWS.toLocaleString()}`);
  console.log(`[Invariant] Reconciliation Passed   = ${reconciliationPassed ? 'YES (TRUE)' : 'NO (FAILED)'}`);

  if (!reconciliationPassed) {
    throw new Error('Reconciliation invariant failed!');
  }

  // -------------------------------------------------------------------------
  // STAGE 4: BOUNDARY TESTS & RANDOM ACCESS VALIDATION
  // -------------------------------------------------------------------------
  console.log('\n--- STAGE 4: Boundary & Random Access Validation ---');

  // Helper to read raw row text from partition binary using .idx
  function readRowFromPartition(partNum: number, rowOffsetInPart: number): string {
    const binPath = getPartitionBinPath(session.datasetId, partNum, false);
    const idxPath = getPartitionIdxPath(session.datasetId, partNum, false);

    const idxFd = fs.openSync(idxPath, 'r');
    const idxBuf = Buffer.alloc(16);
    fs.readSync(idxFd, idxBuf, 0, 16, rowOffsetInPart * 8);
    fs.closeSync(idxFd);

    const rowStart = idxBuf.readBigUInt64LE(0);
    const rowEnd = idxBuf.readBigUInt64LE(8);
    const rowLen = Number(rowEnd - rowStart);

    const binFd = fs.openSync(binPath, 'r');
    const rowBuf = Buffer.alloc(rowLen);
    fs.readSync(binFd, rowBuf, 0, rowLen, Number(rowStart));
    fs.closeSync(binFd);

    return rowBuf.toString('utf8').replace(/\r?\n$/, '');
  }

  // Verify boundary rows:
  // Part 0 boundaries: row 0, row 1,999,999
  // Part 1 boundaries: row 2,000,000 (part 1 offset 0), row 3,999,999
  // Part 2 boundaries: row 4,000,000 (part 2 offset 0), row 5,999,999
  // Part 3 boundaries: row 6,000,000 (part 3 offset 0), row 7,999,999
  // Part 4 boundaries: row 8,000,000 (part 4 offset 0), row 9,999,999
  // Part 5 boundaries: row 10,000,000 (part 5 offset 0), row 11,809,232
  const boundaryCheckRows = [
    { stableRowId: 0, part: 0, offset: 0 },
    { stableRowId: 1_999_999, part: 0, offset: 1_999_999 },
    { stableRowId: 2_000_000, part: 1, offset: 0 },
    { stableRowId: 3_999_999, part: 1, offset: 1_999_999 },
    { stableRowId: 4_000_000, part: 2, offset: 0 },
    { stableRowId: 5_999_999, part: 2, offset: 1_999_999 },
    { stableRowId: 6_000_000, part: 3, offset: 0 },
    { stableRowId: 7_999_999, part: 3, offset: 1_999_999 },
    { stableRowId: 8_000_000, part: 4, offset: 0 },
    { stableRowId: 9_999_999, part: 4, offset: 1_999_999 },
    { stableRowId: 10_000_000, part: 5, offset: 0 },
    { stableRowId: 11_809_232, part: 5, offset: 1_809_232 },
  ];

  let boundaryTestsPassed = true;
  for (const b of boundaryCheckRows) {
    const rowStr = readRowFromPartition(b.part, b.offset);
    const cells = CsvImportWorker.parseCsvLine(rowStr);
    const isValid = cells.length === EXPECTED_COLUMNS && rowStr.length > 10;
    if (!isValid) {
      console.error(`  [FAIL] Boundary row ${b.stableRowId} in Part ${b.part}: length=${rowStr.length}, cols=${cells.length}`);
      boundaryTestsPassed = false;
    } else {
      console.log(`  [PASS] Boundary row ${b.stableRowId.toLocaleString()} (Part ${b.part}, offset ${b.offset.toLocaleString()}): ${cells.length} cols, snippet: "${rowStr.substring(0, 45)}..."`);
    }
  }

  // Random access test on 5 arbitrary stable row ordinals
  const randomRows = [42, 543_210, 2_500_000, 7_890_123, 11_111_111];
  let randomAccessTestsPassed = true;

  console.log('\n[Random Access Verification]');
  for (const r of randomRows) {
    const partNum = Math.floor(r / PARTITION_ROW_LIMIT);
    const offsetInPart = r % PARTITION_ROW_LIMIT;
    const rowStr = readRowFromPartition(partNum, offsetInPart);
    const cells = CsvImportWorker.parseCsvLine(rowStr);
    const isValid = cells.length === EXPECTED_COLUMNS;
    if (!isValid) {
      randomAccessTestsPassed = false;
      console.error(`  [FAIL] Random row ${r}: got ${cells.length} cols`);
    } else {
      console.log(`  [PASS] Random row ${r.toLocaleString()} -> Part ${partNum} offset ${offsetInPart.toLocaleString()}: ${cells.length} cols, snippet: "${rowStr.substring(0, 50)}..."`);
    }
  }

  clearInterval(memInterval);
  const totalDurationSec = (Date.now() - overallStartTime) / 1000;

  const peakRssMB = peakRss / (1024 * 1024);
  const peakHeapUsedMB = peakHeapUsed / (1024 * 1024);
  const peakHeapTotalMB = peakHeapTotal / (1024 * 1024);

  console.log('\n' + '='.repeat(80));
  console.log('   PHASE 1A BENCHMARK RESULTS SUMMARY');
  console.log('='.repeat(80));
  console.log(`Upload Duration:        ${uploadDurationSec.toFixed(2)} s (${uploadSpeedMBps.toFixed(2)} MB/s)`);
  console.log(`Import Duration:        ${importDurationSec.toFixed(2)} s (${importRowsPerSec.toLocaleString()} rows/s, ${importSpeedMBps.toFixed(2)} MB/s)`);
  console.log(`Total Pipeline Time:    ${totalDurationSec.toFixed(2)} s`);
  console.log(`Peak Worker RSS:        ${peakRssMB.toFixed(2)} MB`);
  console.log(`Peak Heap Used:         ${peakHeapUsedMB.toFixed(2)} MB`);
  console.log(`Peak Heap Total:        ${peakHeapTotalMB.toFixed(2)} MB`);
  console.log(`Boundary Tests:         ${boundaryTestsPassed ? 'PASSED' : 'FAILED'}`);
  console.log(`Random Access Tests:    ${randomAccessTestsPassed ? 'PASSED' : 'FAILED'}`);
  console.log(`Reconciliation Invariant: ${reconciliationPassed ? 'PASSED' : 'FAILED'}`);

  const metrics: ExecutionMetrics = {
    uploadDurationSec,
    uploadSpeedMBps,
    importDurationSec,
    importRowsPerSec,
    importSpeedMBps,
    totalDurationSec,
    peakRssMB,
    peakHeapUsedMB,
    peakHeapTotalMB,
    totalDataRows: importResult.totalDataRows,
    columnCount: importResult.columnCount,
    partitionsCount: importResult.partitionsCount,
    partitions: partitionsData,
    boundaryTestsPassed,
    randomAccessTestsPassed,
    reconciliationPassed,
  };

  const resultsPath = path.join(process.cwd(), 'benchmarks', 'phase1a-large-csv', 'phase1a_execution_results.json');
  fs.writeFileSync(resultsPath, JSON.stringify(metrics, null, 2), 'utf8');
  console.log(`\nMetrics written to: ${resultsPath}`);

  await closeDatasetImportQueue();
  await prisma.$disconnect();

  if (!boundaryTestsPassed || !randomAccessTestsPassed || !reconciliationPassed) {
    process.exit(1);
  }
  process.exit(0);
}

main().catch(async (err) => {
  console.error('\n[FATAL ERROR]', err);
  await closeDatasetImportQueue();
  await prisma.$disconnect();
  process.exit(1);
});
