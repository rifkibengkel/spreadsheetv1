import fs from 'fs';
import path from 'path';
import prisma from '../../src/lib/db/prisma';
import { DatasetRangeReader, DatasetRangeResult } from '../../src/lib/largeCsv/reader/datasetRangeReader';
import { CsvImportWorker } from '../../src/lib/largeCsv/worker/csvImportWorker';
import { BoundedLruCache } from '../../src/lib/viewport/boundedLruCache';

const DATASET_ID = '80af4daa-3c4a-42eb-b656-dd855dc6f866';
const SOURCE_CSV = 'C:\\Users\\BKKI-1\\Downloads\\archive\\parking_violation.csv';
const RESULTS_JSON_PATH = path.join(__dirname, 'phase1b_benchmark_results.json');

interface LatencyStats {
  min: number;
  p50: number;
  p95: number;
  p99: number;
  max: number;
  mean: number;
}

function computeStats(samples: number[]): LatencyStats {
  if (samples.length === 0) return { min: 0, p50: 0, p95: 0, p99: 0, max: 0, mean: 0 };
  const sorted = [...samples].sort((a, b) => a - b);
  const sum = sorted.reduce((acc, v) => acc + v, 0);
  const percentile = (p: number) => {
    const idx = Math.min(sorted.length - 1, Math.max(0, Math.floor((p / 100) * sorted.length)));
    return sorted[idx];
  };

  return {
    min: sorted[0],
    p50: percentile(50),
    p95: percentile(95),
    p99: percentile(99),
    max: sorted[sorted.length - 1],
    mean: sum / sorted.length,
  };
}

async function runPhase1bBenchmarks() {
  console.log('='.repeat(80));
  console.log('   PHASE 1B — E2E RANGE QUERY & VIEWPORT BENCHMARK ON 11.8M ROW DATASET');
  console.log(`   Dataset ID: ${DATASET_ID}`);
  console.log('='.repeat(80));

  const dataset = await prisma.dataset.findUnique({
    where: { id: DATASET_ID },
    include: {
      partitions: { orderBy: { partitionNumber: 'asc' } },
    },
  });

  if (!dataset) {
    throw new Error(`Dataset ${DATASET_ID} not found in PostgreSQL`);
  }

  console.log(`\nDataset Info:`);
  console.log(`  Name:        ${dataset.name}`);
  console.log(`  Status:      ${dataset.status}`);
  console.log(`  Total Rows:  ${dataset.rowCount}`);
  console.log(`  Columns:     ${dataset.columnCount}`);
  console.log(`  Partitions:  ${dataset.partitions.length}`);

  // ---------------------------------------------------------------------------
  // SECTION 1: Required Real Dataset Test Suites (A..F)
  // ---------------------------------------------------------------------------
  console.log('\n' + '-'.repeat(80));
  console.log('SECTION 1: EXECUTING REQUIRED REAL DATASET SUITES (A..F)');
  console.log('-'.repeat(80));

  const suiteResults: Record<string, any> = {};

  // Suite A: First rows (0..100)
  console.log('\n[Suite A] First Rows (startRow=0, rowCount=100)');
  const resA = await DatasetRangeReader.readRange(DATASET_ID, 0, 100);
  suiteResults['Suite A: First Rows'] = {
    startRow: resA.startRow,
    rowCount: resA.rowCount,
    firstStableId: resA.rows[0].stableRowId,
    lastStableId: resA.rows[99].stableRowId,
    columnCount: resA.columnCount,
    serverTotalMs: resA.metrics.serverTotalMs,
  };
  console.log(`  Read ${resA.rowCount} rows in ${resA.metrics.serverTotalMs.toFixed(2)} ms (columns=${resA.columnCount})`);

  // Suite B: Normal viewport (100,000..500)
  console.log('\n[Suite B] Normal Viewport (startRow=100,000, rowCount=500)');
  const resB = await DatasetRangeReader.readRange(DATASET_ID, 100_000, 500);
  suiteResults['Suite B: Normal Viewport'] = {
    startRow: resB.startRow,
    rowCount: resB.rowCount,
    firstStableId: resB.rows[0].stableRowId,
    lastStableId: resB.rows[499].stableRowId,
    columnCount: resB.columnCount,
    serverTotalMs: resB.metrics.serverTotalMs,
  };
  console.log(`  Read ${resB.rowCount} rows in ${resB.metrics.serverTotalMs.toFixed(2)} ms`);

  // Suite C: Boundary 0 -> 1 (1,999,500..1,000)
  console.log('\n[Suite C] Partition Boundary 0 -> 1 (startRow=1,999,500, rowCount=1,000)');
  const resC = await DatasetRangeReader.readRange(DATASET_ID, 1_999_500, 1_000);
  suiteResults['Suite C: Boundary 0->1'] = {
    startRow: resC.startRow,
    rowCount: resC.rowCount,
    boundaryRowBefore: resC.rows[499].stableRowId,
    boundaryRowAfter: resC.rows[500].stableRowId,
    serverTotalMs: resC.metrics.serverTotalMs,
  };
  console.log(`  Boundary: row ${resC.rows[499].stableRowId} -> ${resC.rows[500].stableRowId} in ${resC.metrics.serverTotalMs.toFixed(2)} ms`);

  // Suite D: Boundary 1 -> 2 (3,999,500..1,000)
  console.log('\n[Suite D] Partition Boundary 1 -> 2 (startRow=3,999,500, rowCount=1,000)');
  const resD = await DatasetRangeReader.readRange(DATASET_ID, 3_999_500, 1_000);
  suiteResults['Suite D: Boundary 1->2'] = {
    startRow: resD.startRow,
    rowCount: resD.rowCount,
    boundaryRowBefore: resD.rows[499].stableRowId,
    boundaryRowAfter: resD.rows[500].stableRowId,
    serverTotalMs: resD.metrics.serverTotalMs,
  };
  console.log(`  Boundary: row ${resD.rows[499].stableRowId} -> ${resD.rows[500].stableRowId} in ${resD.metrics.serverTotalMs.toFixed(2)} ms`);

  // Suite E: Later boundaries (2->3, 3->4, 4->5)
  console.log('\n[Suite E] Later Partition Boundaries (5,999,500, 7,999,500, 9,999,500)');
  const resE1 = await DatasetRangeReader.readRange(DATASET_ID, 5_999_500, 1_000);
  const resE2 = await DatasetRangeReader.readRange(DATASET_ID, 7_999_500, 1_000);
  const resE3 = await DatasetRangeReader.readRange(DATASET_ID, 9_999_500, 1_000);
  suiteResults['Suite E: Boundary 2->3'] = { durationMs: resE1.metrics.serverTotalMs };
  suiteResults['Suite E: Boundary 3->4'] = { durationMs: resE2.metrics.serverTotalMs };
  suiteResults['Suite E: Boundary 4->5'] = { durationMs: resE3.metrics.serverTotalMs };
  console.log(`  Boundary 2->3: row ${resE1.rows[499].stableRowId} -> ${resE1.rows[500].stableRowId} in ${resE1.metrics.serverTotalMs.toFixed(2)} ms`);
  console.log(`  Boundary 3->4: row ${resE2.rows[499].stableRowId} -> ${resE2.rows[500].stableRowId} in ${resE2.metrics.serverTotalMs.toFixed(2)} ms`);
  console.log(`  Boundary 4->5: row ${resE3.rows[499].stableRowId} -> ${resE3.rows[500].stableRowId} in ${resE3.metrics.serverTotalMs.toFixed(2)} ms`);

  // Suite F: Final rows (11,809,133..100 and final 1 row)
  console.log('\n[Suite F] Final Rows (startRow=11,809,133, rowCount=100)');
  const resF = await DatasetRangeReader.readRange(DATASET_ID, 11_809_133, 100);
  const resF1 = await DatasetRangeReader.readRange(DATASET_ID, 11_809_232, 1);
  suiteResults['Suite F: Final 100 Rows'] = {
    startRow: resF.startRow,
    rowCount: resF.rowCount,
    firstStableId: resF.rows[0].stableRowId,
    lastStableId: resF.rows[99].stableRowId,
    isNextRowNull: resF.nextRow === null,
    serverTotalMs: resF.metrics.serverTotalMs,
  };
  suiteResults['Suite F: Final 1 Row'] = {
    startRow: resF1.startRow,
    stableRowId: resF1.rows[0].stableRowId,
    serverTotalMs: resF1.metrics.serverTotalMs,
  };
  console.log(`  Final 100 rows: [${resF.rows[0].stableRowId} .. ${resF.rows[99].stableRowId}], nextRow=${resF.nextRow}`);
  console.log(`  Final 1 row: stableRowId=${resF1.rows[0].stableRowId} in ${resF1.metrics.serverTotalMs.toFixed(2)} ms`);

  // ---------------------------------------------------------------------------
  // SECTION 2: Strict Textual Data Fidelity Audit (Source CSV vs .bin vs API)
  // ---------------------------------------------------------------------------
  console.log('\n' + '-'.repeat(80));
  console.log('SECTION 2: MULTI-LAYER DATA FIDELITY COMPARISON');
  console.log('-'.repeat(80));

  const fidelityAudit: any[] = [];

  if (fs.existsSync(SOURCE_CSV)) {
    // Audit Row 0
    const sourceFd = fs.openSync(SOURCE_CSV, 'r');
    const headerAndRow0Buf = Buffer.alloc(16384);
    fs.readSync(sourceFd, headerAndRow0Buf, 0, 16384, 0);
    fs.closeSync(sourceFd);

    const sourceLines = headerAndRow0Buf.toString('utf8').split('\n');
    const sourceRow0Line = sourceLines[1].replace(/\r$/, '');
    const sourceRow0Cells = CsvImportWorker.parseCsvLine(sourceRow0Line);

    const apiRow0Cells = resA.rows[0].cells;

    let row0Mismatch = -1;
    for (let c = 0; c < 54; c++) {
      if (sourceRow0Cells[c] !== apiRow0Cells[c]) {
        row0Mismatch = c;
        break;
      }
    }

    fidelityAudit.push({
      testCase: 'Row 0 (First Dataset Row)',
      columnsCompared: 54,
      isIdentical: row0Mismatch === -1,
      firstMismatchCol: row0Mismatch,
      sampleValues: {
        col0_ordinal: { source: sourceRow0Cells[0], api: apiRow0Cells[0] },
        col1_summons: { source: sourceRow0Cells[1], api: apiRow0Cells[1] },
        col2_plate: { source: sourceRow0Cells[2], api: apiRow0Cells[2] },
        col3_state: { source: sourceRow0Cells[3], api: apiRow0Cells[3] },
        col4_type: { source: sourceRow0Cells[4], api: apiRow0Cells[4] },
      },
    });

    // Audit Boundary Row 1,999,999 (End of Partition 0)
    const part0 = dataset.partitions[0];
    const idx0Fd = fs.openSync(part0.indexFilePath, 'r');
    const idx0Buf = Buffer.alloc(16);
    fs.readSync(idx0Fd, idx0Buf, 0, 16, 1_999_999 * 8);
    fs.closeSync(idx0Fd);
    const p0Start = Number(idx0Buf.readBigUInt64LE(0));
    const p0End = Number(idx0Buf.readBigUInt64LE(8));
    const bin0Fd = fs.openSync(part0.dataFilePath, 'r');
    const bin0Buf = Buffer.alloc(p0End - p0Start);
    fs.readSync(bin0Fd, bin0Buf, 0, p0End - p0Start, p0Start);
    fs.closeSync(bin0Fd);
    const bin0Line = bin0Buf.toString('utf8').replace(/\r?\n$/, '');
    const bin0Cells = CsvImportWorker.parseCsvLine(bin0Line);

    const apiRow1999999 = resC.rows[499].cells;
    let boundary0Mismatch = -1;
    for (let c = 0; c < 54; c++) {
      if (bin0Cells[c] !== apiRow1999999[c]) {
        boundary0Mismatch = c;
        break;
      }
    }

    fidelityAudit.push({
      testCase: 'Row 1,999,999 (Partition 0 Terminal Boundary)',
      columnsCompared: 54,
      isIdentical: boundary0Mismatch === -1,
      firstMismatchCol: boundary0Mismatch,
      sampleValues: {
        col0_ordinal: { bin: bin0Cells[0], api: apiRow1999999[0] },
        col1_summons: { bin: bin0Cells[1], api: apiRow1999999[1] },
        col2_plate: { bin: bin0Cells[2], api: apiRow1999999[2] },
      },
    });

    // Audit Boundary Row 2,000,000 (Start of Partition 1)
    const part1 = dataset.partitions[1];
    const idx1Fd = fs.openSync(part1.indexFilePath, 'r');
    const idx1Buf = Buffer.alloc(16);
    fs.readSync(idx1Fd, idx1Buf, 0, 16, 0);
    fs.closeSync(idx1Fd);
    const p1Start = Number(idx1Buf.readBigUInt64LE(0));
    const p1End = Number(idx1Buf.readBigUInt64LE(8));
    const bin1Fd = fs.openSync(part1.dataFilePath, 'r');
    const bin1Buf = Buffer.alloc(p1End - p1Start);
    fs.readSync(bin1Fd, bin1Buf, 0, p1End - p1Start, p1Start);
    fs.closeSync(bin1Fd);
    const bin1Line = bin1Buf.toString('utf8').replace(/\r?\n$/, '');
    const bin1Cells = CsvImportWorker.parseCsvLine(bin1Line);

    const apiRow2000000 = resC.rows[500].cells;
    let boundary1Mismatch = -1;
    for (let c = 0; c < 54; c++) {
      if (bin1Cells[c] !== apiRow2000000[c]) {
        boundary1Mismatch = c;
        break;
      }
    }

    fidelityAudit.push({
      testCase: 'Row 2,000,000 (Partition 1 Starting Boundary)',
      columnsCompared: 54,
      isIdentical: boundary1Mismatch === -1,
      firstMismatchCol: boundary1Mismatch,
      sampleValues: {
        col0_ordinal: { bin: bin1Cells[0], api: apiRow2000000[0] },
        col1_summons: { bin: bin1Cells[1], api: apiRow2000000[1] },
        col2_plate: { bin: bin1Cells[2], api: apiRow2000000[2] },
      },
    });

    // Audit Final Row 11,809,232 (Terminal Row of Dataset)
    const part5 = dataset.partitions[5];
    const localLast = 11_809_232 - Number(part5.stableRowStart);
    const idx5Fd = fs.openSync(part5.indexFilePath, 'r');
    const idx5Buf = Buffer.alloc(16);
    fs.readSync(idx5Fd, idx5Buf, 0, 16, localLast * 8);
    fs.closeSync(idx5Fd);
    const p5Start = Number(idx5Buf.readBigUInt64LE(0));
    const p5End = Number(idx5Buf.readBigUInt64LE(8));
    const bin5Fd = fs.openSync(part5.dataFilePath, 'r');
    const bin5Buf = Buffer.alloc(p5End - p5Start);
    fs.readSync(bin5Fd, bin5Buf, 0, p5End - p5Start, p5Start);
    fs.closeSync(bin5Fd);
    const bin5Line = bin5Buf.toString('utf8').replace(/\r?\n$/, '');
    const bin5Cells = CsvImportWorker.parseCsvLine(bin5Line);

    const apiRowFinal = resF1.rows[0].cells;
    let finalMismatch = -1;
    for (let c = 0; c < 54; c++) {
      if (bin5Cells[c] !== apiRowFinal[c]) {
        finalMismatch = c;
        break;
      }
    }

    fidelityAudit.push({
      testCase: 'Row 11,809,232 (Dataset Final Row)',
      columnsCompared: 54,
      isIdentical: finalMismatch === -1,
      firstMismatchCol: finalMismatch,
      sampleValues: {
        col0_ordinal: { bin: bin5Cells[0], api: apiRowFinal[0] },
        col1_summons: { bin: bin5Cells[1], api: apiRowFinal[1] },
        col2_plate: { bin: bin5Cells[2], api: apiRowFinal[2] },
      },
    });
  }

  for (const fa of fidelityAudit) {
    console.log(`  [Fidelity] ${fa.testCase}: ${fa.isIdentical ? '100% MATCH (54/54 cols)' : 'MISMATCH'}`);
  }

  // ---------------------------------------------------------------------------
  // SECTION 3: Latency Distribution & Timing Decomposition (T1..T5)
  // ---------------------------------------------------------------------------
  console.log('\n' + '-'.repeat(80));
  console.log('SECTION 3: TIMING OBSERVABILITY & LATENCY DISTRIBUTION (COLD VS WARM)');
  console.log('-'.repeat(80));

  // 3.1 Cold Server Read (first read into partition 3)
  console.log('Measuring Cold Server Read (first unmapped partition seek)...');
  const coldRes = await DatasetRangeReader.readRange(DATASET_ID, 6_500_000, 200);
  console.log(`  Cold Server Total: ${coldRes.metrics.serverTotalMs.toFixed(2)} ms (Decode: ${coldRes.metrics.decodeMs.toFixed(2)} ms, Read: ${coldRes.metrics.fileReadMs.toFixed(2)} ms)`);

  const ITERATIONS = 30;
  console.log(`Executing ${ITERATIONS} varied warm range queries across all partitions...`);

  const t1Samples: number[] = [];
  const t2Samples: number[] = [];
  const t3Samples: number[] = [];
  const t4Samples: number[] = [];
  const t5Samples: number[] = [];

  const testRanges = [
    { start: 0, count: 200 },
    { start: 500_000, count: 500 },
    { start: 1_999_800, count: 400 }, // boundary 0->1
    { start: 2_500_000, count: 300 },
    { start: 3_999_700, count: 600 }, // boundary 1->2
    { start: 6_000_000, count: 250 },
    { start: 7_999_900, count: 300 }, // boundary 3->4
    { start: 10_000_000, count: 500 },
    { start: 11_809_000, count: 233 }, // final rows
  ];

  for (let i = 0; i < ITERATIONS; i++) {
    const range = testRanges[i % testRanges.length];
    const res = await DatasetRangeReader.readRange(DATASET_ID, range.start, range.count);
    t1Samples.push(res.metrics.partitionResolutionMs);
    t2Samples.push(res.metrics.indexLookupMs);
    t3Samples.push(res.metrics.fileReadMs);
    t4Samples.push(res.metrics.decodeMs);
    t5Samples.push(res.metrics.serverTotalMs);
  }

  const t1Stats = computeStats(t1Samples);
  const t2Stats = computeStats(t2Samples);
  const t3Stats = computeStats(t3Samples);
  const t4Stats = computeStats(t4Samples);
  const t5Stats = computeStats(t5Samples);

  console.log('\nWarm Stage Latencies Decomposition (ms across 30 iterations):');
  console.log('Stage                     | Min    | p50    | p95    | p99    | Max    | Mean');
  console.log('-'.repeat(75));
  const printRow = (name: string, s: LatencyStats) => {
    console.log(
      `${name.padEnd(25)} | ${s.min.toFixed(2).padStart(6)} | ${s.p50.toFixed(2).padStart(6)} | ${s.p95.toFixed(2).padStart(6)} | ${s.p99.toFixed(2).padStart(6)} | ${s.max.toFixed(2).padStart(6)} | ${s.mean.toFixed(2).padStart(6)}`
    );
  };

  printRow('T1: Partition Resolution', t1Stats);
  printRow('T2: Index Lookup (.idx)', t2Stats);
  printRow('T3: Bounded File Read', t3Stats);
  printRow('T4: RFC4180 Row Decode', t4Stats);
  printRow('T5: Server Total (E2E)', t5Stats);

  // 3.2 Worker Cache Hit vs Miss benchmark
  console.log('\nMeasuring Worker Cache Hit vs Miss latency:');
  const benchCache = new BoundedLruCache(25, 200);
  const tMiss0 = performance.now();
  const missData = await DatasetRangeReader.readRange(DATASET_ID, 1_000_000, 200);
  benchCache.set(DATASET_ID, 5000, {
    pageIndex: 5000,
    startRow: 1_000_000,
    rowCount: missData.rowCount,
    rows: missData.rows,
    cachedAt: Date.now(),
  });
  const tMiss1 = performance.now();
  const workerMissMs = tMiss1 - tMiss0;

  const tHit0 = performance.now();
  const hitPage = benchCache.get(DATASET_ID, 5000);
  const tHit1 = performance.now();
  const workerHitMs = tHit1 - tHit0;

  console.log(`  Worker Cache Miss (Server Read + Decode + Cache Insert): ${workerMissMs.toFixed(2)} ms`);
  console.log(`  Worker Cache Hit  (In-Memory BoundedLruCache lookup):     ${workerHitMs.toFixed(3)} ms`);

  // ---------------------------------------------------------------------------
  // SECTION 4: Viewport Memory Boundedness Simulation
  // ---------------------------------------------------------------------------
  console.log('\n' + '-'.repeat(80));
  console.log('SECTION 4: VIEWPORT MEMORY & HYDRATION BOUNDEDNESS');
  console.log('-'.repeat(80));

  // Measure memory of BoundedLruCache under sustained scrolling (1,000 page loads)
  const lru = new BoundedLruCache(25, 200); // 25 pages x 200 rows = max 5,000 rows
  const initialHeap = process.memoryUsage().heapUsed;

  console.log(`Simulating 1,000 viewport requests through BoundedLruCache...`);
  for (let i = 0; i < 1_000; i++) {
    const pageIndex = (i * 3) % 50_000;
    // Simulate 200-row page
    const sampleRows: any[] = [];
    for (let r = 0; r < 200; r++) {
      sampleRows.push({
        stableRowId: pageIndex * 200 + r,
        cells: ['8002531292', 'EPC5238', 'NY', 'PAS', '06/23/2013', '00', '00', '00'],
      });
    }

    lru.set(DATASET_ID, pageIndex, {
      pageIndex,
      startRow: pageIndex * 200,
      rowCount: 200,
      rows: sampleRows,
      cachedAt: Date.now(),
    });
  }

  const postScrollHeap = process.memoryUsage().heapUsed;
  const lruStats = lru.getStats();

  console.log(`  Cache Final Size:     ${lruStats.size} pages (configured max: ${lruStats.maxPages})`);
  console.log(`  Cache Max Rows Held:  ${lruStats.size * lruStats.pageSize} rows`);
  console.log(`  Total Evictions:      ${lruStats.evictions}`);
  console.log(`  Heap delta after 1K:  ${((postScrollHeap - initialHeap) / 1024 / 1024).toFixed(2)} MB`);

  // ---------------------------------------------------------------------------
  // SECTION 5: Write Results to JSON
  // ---------------------------------------------------------------------------
  const benchmarkSummary = {
    datasetId: DATASET_ID,
    totalRows: 11_809_233,
    columnCount: 54,
    partitionsCount: 6,
    timestamp: new Date().toISOString(),
    suiteResults,
    fidelityAudit,
    latencyStats: {
      t1PartitionResolution: t1Stats,
      t2IndexLookup: t2Stats,
      t3FileRead: t3Stats,
      t4Decode: t4Stats,
      t5ServerTotal: t5Stats,
    },
    memoryBoundedness: {
      maxPages: lruStats.maxPages,
      pageSize: lruStats.pageSize,
      maxRetainedRows: lruStats.maxPages * lruStats.pageSize,
      evictionsIn1000Scrolls: lruStats.evictions,
    },
  };

  fs.writeFileSync(RESULTS_JSON_PATH, JSON.stringify(benchmarkSummary, null, 2), 'utf8');
  console.log(`\nBenchmark results successfully recorded to: ${RESULTS_JSON_PATH}`);
  console.log('='.repeat(80));
}

runPhase1bBenchmarks()
  .catch((err) => {
    console.error('Fatal benchmark failure:', err);
    process.exit(1);
  })
  .finally(() => process.exit(0));
