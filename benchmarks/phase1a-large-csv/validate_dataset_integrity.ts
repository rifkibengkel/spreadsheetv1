import fs from 'fs';
import path from 'path';
import prisma from '../../src/lib/db/prisma';
import { CsvImportWorker, PARTITION_ROW_LIMIT } from '../../src/lib/largeCsv/worker/csvImportWorker';
import {
  getPartitionBinPath,
  getPartitionIdxPath,
} from '../../src/lib/largeCsv/storage/storagePaths';

const DATASET_ID = '80af4daa-3c4a-42eb-b656-dd855dc6f866';
const EXPECTED_SOURCE_SIZE = 3_121_309_240;
const EXPECTED_ROWS = 11_809_233;
const EXPECTED_COLS = 54;

function readRowFromPartition(datasetId: string, partNum: number, rowOffsetInPart: number): string {
  const binPath = getPartitionBinPath(datasetId, partNum, false);
  const idxPath = getPartitionIdxPath(datasetId, partNum, false);

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

async function validate() {
  console.log('='.repeat(80));
  console.log('   POST-IMPORT FORENSIC VALIDATION OF LARGE CSV DATASET');
  console.log(`   Dataset ID: ${DATASET_ID}`);
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

  const dataset = await prisma.dataset.findUnique({
    where: { id: DATASET_ID },
    include: {
      partitions: { orderBy: { partitionNumber: 'asc' } },
    },
  });

  assert(!!dataset, 'Dataset record exists in PostgreSQL');
  if (!dataset) return;

  assert(dataset.status === 'READY', `Dataset status is READY (actual: ${dataset.status})`);
  assert(Number(dataset.byteSize) === EXPECTED_SOURCE_SIZE, `Source size matches 3,121,309,240 bytes (got ${dataset.byteSize})`);
  assert(Number(dataset.rowCount) === EXPECTED_ROWS, `Total rows matches 11,809,233 (got ${dataset.rowCount})`);
  assert(dataset.columnCount === EXPECTED_COLS, `Column count matches 54 (got ${dataset.columnCount})`);
  assert(dataset.partitions.length === 6, `Partition count equals 6 (got ${dataset.partitions.length})`);

  let sumPartitionRows = 0;
  let expectedNextRow = BigInt(0);

  for (const part of dataset.partitions) {
    assert(part.status === 'READY', `Partition ${part.partitionNumber} status is READY`);
    assert(part.stableRowStart === expectedNextRow, `Partition ${part.partitionNumber} starts at ${expectedNextRow}`);
    expectedNextRow = part.stableRowEnd + BigInt(1);

    const binStat = fs.statSync(part.dataFilePath);
    const idxStat = fs.statSync(part.indexFilePath);
    const expectedIdxBytes = (part.rowCount + 1) * 8;

    assert(idxStat.size === expectedIdxBytes, `Part ${part.partitionNumber} .idx size matches (rowCount+1)*8 = ${expectedIdxBytes}`);
    assert(BigInt(binStat.size) === part.byteSize, `Part ${part.partitionNumber} .bin size matches DB ${part.byteSize}`);

    sumPartitionRows += part.rowCount;
  }

  assert(sumPartitionRows === EXPECTED_ROWS, `SUM(partition.rowCount) (${sumPartitionRows}) === dataset.rowCount (${dataset.rowCount}) === 11,809,233`);

  // Explicit boundary checks requested:
  // 1,999,999 / 2,000,000 / 2,000,001
  // and equivalent boundaries
  const boundaryTests = [
    { row: 0, part: 0, offset: 0, desc: 'First row of dataset' },
    { row: 1_999_999, part: 0, offset: 1_999_999, desc: 'Partition 0 boundary row (1,999,999)' },
    { row: 2_000_000, part: 1, offset: 0, desc: 'Partition 1 start row (2,000,000)' },
    { row: 2_000_001, part: 1, offset: 1, desc: 'Partition 1 second row (2,000,001)' },
    { row: 3_999_999, part: 1, offset: 1_999_999, desc: 'Partition 1 boundary row (3,999,999)' },
    { row: 4_000_000, part: 2, offset: 0, desc: 'Partition 2 start row (4,000,000)' },
    { row: 4_000_001, part: 2, offset: 1, desc: 'Partition 2 second row (4,000,001)' },
    { row: 5_999_999, part: 2, offset: 1_999_999, desc: 'Partition 2 boundary row (5,999,999)' },
    { row: 6_000_000, part: 3, offset: 0, desc: 'Partition 3 start row (6,000,000)' },
    { row: 6_000_001, part: 3, offset: 1, desc: 'Partition 3 second row (6,000,001)' },
    { row: 7_999_999, part: 3, offset: 1_999_999, desc: 'Partition 3 boundary row (7,999,999)' },
    { row: 8_000_000, part: 4, offset: 0, desc: 'Partition 4 start row (8,000,000)' },
    { row: 8_000_001, part: 4, offset: 1, desc: 'Partition 4 second row (8,000,001)' },
    { row: 9_999_999, part: 4, offset: 1_999_999, desc: 'Partition 4 boundary row (9,999,999)' },
    { row: 10_000_000, part: 5, offset: 0, desc: 'Partition 5 start row (10,000,000)' },
    { row: 10_000_001, part: 5, offset: 1, desc: 'Partition 5 second row (10,000,001)' },
    { row: 11_809_232, part: 5, offset: 1_809_232, desc: 'Final row of dataset (11,809,232)' },
  ];

  console.log('\n--- Partition Boundary Inspections ---');
  for (const b of boundaryTests) {
    const rawLine = readRowFromPartition(DATASET_ID, b.part, b.offset);
    const cells = CsvImportWorker.parseCsvLine(rawLine);
    assert(
      cells.length === EXPECTED_COLS && rawLine.length > 5,
      `Boundary [${b.desc}]: ${cells.length} columns parsed`
    );
  }

  // Random access tests
  console.log('\n--- Random Row Access Verification ---');
  const randomRowIds = [7, 12345, 123456, 3141592, 6543210, 10500000];
  for (const r of randomRowIds) {
    const part = Math.floor(r / PARTITION_ROW_LIMIT);
    const offset = r % PARTITION_ROW_LIMIT;
    const rawLine = readRowFromPartition(DATASET_ID, part, offset);
    const cells = CsvImportWorker.parseCsvLine(rawLine);
    assert(
      cells.length === EXPECTED_COLS,
      `Random Row ${r.toLocaleString()} (Part ${part}, Offset ${offset.toLocaleString()}): ${cells.length} cols`
    );
  }

  console.log(`\nValidation complete: ${passed} passed, ${failed} failed.\n`);
  await prisma.$disconnect();
  process.exit(failed > 0 ? 1 : 0);
}

validate().catch(err => {
  console.error('Validation error:', err);
  process.exit(1);
});
