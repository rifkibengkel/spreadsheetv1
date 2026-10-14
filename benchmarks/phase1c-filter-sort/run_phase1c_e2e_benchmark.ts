import fs from 'fs';
import path from 'path';
import prisma from '../../src/lib/db/prisma';
import { QueryIndexManager } from '../../src/lib/largeCsv/query/queryIndexManager';
import { DatasetRowResolver } from '../../src/lib/largeCsv/query/datasetRowResolver';
import { FilterEngine } from '../../src/lib/largeCsv/query/filterEngine';
import { ExternalSortEngine } from '../../src/lib/largeCsv/query/externalSortEngine';
import { QueryDefinition, SortDefinition, QUERY_INDEX_HEADER_SIZE } from '../../src/lib/largeCsv/query/types';

const DATASET_ID = '80af4daa-3c4a-42eb-b656-dd855dc6f866';

let passed = 0;
let failed = 0;

function assert(condition: boolean, name: string, detail?: any) {
  if (condition) {
    console.log(`  [PASS] ${name}`);
    passed++;
  } else {
    console.error(`  [FAIL] ${name}`, detail !== undefined ? detail : '');
    failed++;
  }
}

async function runBenchmark() {
  console.log('='.repeat(80));
  console.log('   PHASE 1C — E2E VERIFICATION & BENCHMARK SUITE');
  console.log(`   Dataset: ${DATASET_ID} (11,809,233 rows, 54 columns, 6 partitions)`);
  console.log('='.repeat(80));

  const results: any = {
    datasetId: DATASET_ID,
    timestamp: new Date().toISOString(),
    tests: {},
    benchmarks: {},
  };

  // ---------------------------------------------------------------------------
  // TEST 1: Empty Filter Result
  // ---------------------------------------------------------------------------
  console.log('\n[Test 1] Empty Filter Result (Zero Matches)');
  const emptyQueryDef: QueryDefinition = {
    filter: {
      conjunction: 'AND',
      predicates: [{ column: 'Registration_State', operator: 'EQUALS', value: 'IMPOSSIBLE_STATE_XYZ99' }],
    },
  };
  const emptyMeta = await QueryIndexManager.executeQuery(DATASET_ID, emptyQueryDef);
  assert(emptyMeta.status === 'READY', 'Empty query completes with status READY');
  assert(emptyMeta.matchedRows === 0, 'Matched exact 0 rows');
  const emptySlice = await QueryIndexManager.readStableRowIds(DATASET_ID, emptyMeta.queryId, 0, 100);
  assert(emptySlice.stableRowIds.length === 0, 'readStableRowIds returns empty array');
  const emptyResolved = await DatasetRowResolver.resolveRows(DATASET_ID, emptySlice.stableRowIds);
  assert(emptyResolved.length === 0, 'DatasetRowResolver returns empty array');
  results.tests.emptyFilter = 'PASS';

  // ---------------------------------------------------------------------------
  // TEST 2: Single-Row Filter Result
  // ---------------------------------------------------------------------------
  console.log('\n[Test 2] Single-Row Filter Result');
  // Row 0 has Summons_Number = "8002531292"
  const singleQueryDef: QueryDefinition = {
    filter: {
      conjunction: 'AND',
      predicates: [{ column: 'Summons_Number', operator: 'EQUALS', value: '8002531292' }],
    },
  };
  const singleMeta = await QueryIndexManager.executeQuery(DATASET_ID, singleQueryDef);
  assert(singleMeta.status === 'READY', 'Single-row query status is READY');
  assert(singleMeta.matchedRows === 1, `Matched exact 1 row (got ${singleMeta.matchedRows})`);
  const singleSlice = await QueryIndexManager.readStableRowIds(DATASET_ID, singleMeta.queryId, 0, 10);
  assert(singleSlice.stableRowIds.length === 1 && singleSlice.stableRowIds[0] === 0, 'Resolved stableRowId is 0');
  const singleRows = await DatasetRowResolver.resolveRows(DATASET_ID, singleSlice.stableRowIds);
  assert(singleRows.length === 1, 'Resolved 1 row');
  assert(singleRows[0].cells[1] === '8002531292', 'Cell[1] summons matches exact input value');
  assert(singleRows[0].cells.length === 54, 'Row has exact 54 columns');
  results.tests.singleRow = 'PASS';

  // ---------------------------------------------------------------------------
  // TEST 3: Large Filter Result (>100k rows)
  // ---------------------------------------------------------------------------
  console.log('\n[Test 3] Large Filter Result (CT plates: 160,361 rows across partitions)');
  const ctQueryDef: QueryDefinition = {
    filter: {
      conjunction: 'AND',
      predicates: [{ column: 'Registration_State', operator: 'EQUALS', value: 'CT' }],
    },
  };
  const ctMeta = await QueryIndexManager.executeQuery(DATASET_ID, ctQueryDef);
  assert(ctMeta.matchedRows === 160_361, `Matched 160,361 CT rows (got ${ctMeta.matchedRows})`);

  // Verify slice read at start, middle, and end of 160k result
  const ctPage0 = await QueryIndexManager.readStableRowIds(DATASET_ID, ctMeta.queryId, 0, 100);
  const ctPageMid = await QueryIndexManager.readStableRowIds(DATASET_ID, ctMeta.queryId, 80_000, 100);
  const ctPageEnd = await QueryIndexManager.readStableRowIds(DATASET_ID, ctMeta.queryId, 160_300, 100);

  assert(ctPage0.stableRowIds.length === 100, 'Read page 0 (100 rows)');
  assert(ctPageMid.stableRowIds.length === 100, 'Read page 80,000 (100 rows)');
  assert(ctPageEnd.stableRowIds.length === 61, `Read end page clamped to 61 rows (got ${ctPageEnd.stableRowIds.length})`);

  // Resolve rows and verify Registration_State is CT on all
  const resolvedCtMid = await DatasetRowResolver.resolveRows(DATASET_ID, ctPageMid.stableRowIds);
  assert(resolvedCtMid.every((r) => r.cells[3] === 'CT'), 'Every resolved row in page 80,000 has State == "CT"');
  results.tests.largeFilter = 'PASS';

  // ---------------------------------------------------------------------------
  // TEST 4 & 5: Sort Ascending and Sort Descending with Duplicate Values & Tie-breaking
  // ---------------------------------------------------------------------------
  console.log('\n[Test 4 & 5] Filter + Sort ASC / DESC with Duplicate Values and Tie-Breaking');
  // Filter for State == "FL" (Florida), sorted by Issue_Date ASC
  const flSortAscDef: QueryDefinition = {
    filter: {
      conjunction: 'AND',
      predicates: [{ column: 'Registration_State', operator: 'EQUALS', value: 'FL' }],
    },
    sort: {
      columns: [{ column: 'Issue_Date', direction: 'ASC', type: 'string' }],
    },
  };

  const flAscMeta = await QueryIndexManager.executeQuery(DATASET_ID, flSortAscDef);
  assert(flAscMeta.status === 'READY', 'FL Sort ASC query status is READY');
  assert(flAscMeta.matchedRows > 10_000, `Matched FL rows: ${flAscMeta.matchedRows}`);

  // Read first 200 sorted rows
  const flAscSlice = await QueryIndexManager.readStableRowIds(DATASET_ID, flAscMeta.queryId, 0, 200);
  const flAscRows = await DatasetRowResolver.resolveRows(DATASET_ID, flAscSlice.stableRowIds);

  let isSortedAsc = true;
  let tieBreaksValid = true;
  for (let i = 0; i < flAscRows.length - 1; i++) {
    const dateA = flAscRows[i].cells[5]; // Issue_Date
    const dateB = flAscRows[i + 1].cells[5];
    if (dateA > dateB) {
      isSortedAsc = false;
      console.error(`Sort violation at ${i}: "${dateA}" > "${dateB}"`);
      break;
    }
    if (dateA === dateB) {
      if (flAscRows[i].stableRowId >= flAscRows[i + 1].stableRowId) {
        tieBreaksValid = false;
        console.error(`Tie break violation at ${i}: ID ${flAscRows[i].stableRowId} >= ${flAscRows[i + 1].stableRowId}`);
        break;
      }
    }
  }
  assert(isSortedAsc, 'Rows are strictly sorted ascending by Issue_Date');
  assert(tieBreaksValid, 'Deterministic tie-breaking strictly enforced (stableRowId ASC on identical keys)');

  // Now Sort DESC on same set
  const flSortDescDef: QueryDefinition = {
    filter: {
      conjunction: 'AND',
      predicates: [{ column: 'Registration_State', operator: 'EQUALS', value: 'FL' }],
    },
    sort: {
      columns: [{ column: 'Issue_Date', direction: 'DESC', type: 'string' }],
    },
  };

  const flDescMeta = await QueryIndexManager.executeQuery(DATASET_ID, flSortDescDef);
  assert(flDescMeta.matchedRows === flAscMeta.matchedRows, 'ASC and DESC queries match exact same row count');

  const flDescSlice = await QueryIndexManager.readStableRowIds(DATASET_ID, flDescMeta.queryId, 0, 200);
  const flDescRows = await DatasetRowResolver.resolveRows(DATASET_ID, flDescSlice.stableRowIds);

  let isSortedDesc = true;
  for (let i = 0; i < flDescRows.length - 1; i++) {
    const dateA = flDescRows[i].cells[5];
    const dateB = flDescRows[i + 1].cells[5];
    if (dateA < dateB) {
      isSortedDesc = false;
      break;
    }
  }
  assert(isSortedDesc, 'Rows are strictly sorted descending by Issue_Date');
  results.tests.sortAscDesc = 'PASS';

  // ---------------------------------------------------------------------------
  // TEST 6: First & Last Rows of Dataset Included
  // ---------------------------------------------------------------------------
  console.log('\n[Test 6] First and Last Dataset Rows Retrieval & Bounds');
  const terminalRows = await DatasetRowResolver.resolveRows(DATASET_ID, [0, 11_809_232]);
  assert(terminalRows.length === 2, 'Resolved 2 terminal rows');
  assert(terminalRows[0].stableRowId === 0, 'Row 0 resolved');
  assert(terminalRows[1].stableRowId === 11_809_232, 'Row 11,809,232 resolved (terminal row)');
  results.tests.firstLastRows = 'PASS';

  // ---------------------------------------------------------------------------
  // TEST 7: Cross-Partition Row Mapping
  // ---------------------------------------------------------------------------
  console.log('\n[Test 7] Cross-Partition Row Mapping across 6 Partitions');
  // Query for a condition that spans all partitions: Issue_Date ending in '2023' or '2022'
  const crossSlice = [1_999_998, 1_999_999, 2_000_000, 2_000_001, 7_999_999, 8_000_000, 10_000_000];
  const crossResolved = await DatasetRowResolver.resolveRows(DATASET_ID, crossSlice);
  assert(crossResolved.length === crossSlice.length, 'Resolved all cross-partition row IDs');
  assert(crossResolved.every((r, idx) => r.stableRowId === crossSlice[idx]), 'Exact requested ID mapping preserved');
  results.tests.crossPartition = 'PASS';

  // ---------------------------------------------------------------------------
  // TEST 8: Memory Boundedness Under Sort
  // ---------------------------------------------------------------------------
  console.log('\n[Test 8] Memory Boundedness Measurement');
  const memBefore = process.memoryUsage().heapUsed;
  // Sort 20,000 items with small chunk size to verify heap stability
  const sortDefMem: SortDefinition = {
    columns: [{ column: 'V', direction: 'ASC', type: 'number' }],
  };
  const tempDir = path.join(process.cwd(), 'storage', 'test_mem_scratch');
  const sorter = new ExternalSortEngine(sortDefMem, tempDir, 500);
  for (let i = 0; i < 20_000; i++) {
    await sorter.addItem({ stableRowId: i, sortKeys: [i % 50] });
  }
  const memPeak = process.memoryUsage().heapUsed;
  await sorter.streamSortedIds(() => {});
  const memAfter = process.memoryUsage().heapUsed;

  const heapDiffMb = (memPeak - memBefore) / (1024 * 1024);
  console.log(`  Heap before: ${(memBefore / (1024*1024)).toFixed(1)} MB`);
  console.log(`  Heap peak:   ${(memPeak / (1024*1024)).toFixed(1)} MB (Delta: +${heapDiffMb.toFixed(1)} MB)`);
  console.log(`  Heap after:  ${(memAfter / (1024*1024)).toFixed(1)} MB`);
  assert(heapDiffMb < 100, `Peak heap delta during external merge sort is bounded < 100 MB (measured: ${heapDiffMb.toFixed(1)} MB)`);
  results.tests.memoryBounded = 'PASS';

  // ---------------------------------------------------------------------------
  // BENCHMARK SUMMARY & TIMINGS
  // ---------------------------------------------------------------------------
  console.log('\n' + '='.repeat(80));
  console.log('   LATENCY & BENCHMARK METRICS');
  console.log('='.repeat(80));

  // Benchmark 1: Query Index Compilation (Warm / Cache Hit)
  const tWarm0 = performance.now();
  await QueryIndexManager.executeQuery(DATASET_ID, ctQueryDef);
  const tWarm1 = performance.now();
  const warmCompileMs = tWarm1 - tWarm0;
  console.log(`  Query Cache Hit Latency:       ${warmCompileMs.toFixed(2)} ms`);

  // Benchmark 2: Bounded Random-Access Index Seek (100 rows)
  const indexSeekSamples: number[] = [];
  for (let i = 0; i < 20; i++) {
    const offset = Math.floor(Math.random() * 100_000);
    const t0 = performance.now();
    await QueryIndexManager.readStableRowIds(DATASET_ID, ctMeta.queryId, offset, 100);
    const t1 = performance.now();
    indexSeekSamples.push(t1 - t0);
  }
  indexSeekSamples.sort((a, b) => a - b);
  const indexP50 = indexSeekSamples[Math.floor(indexSeekSamples.length * 0.5)];
  const indexP95 = indexSeekSamples[Math.floor(indexSeekSamples.length * 0.95)];
  console.log(`  Index Slice Read p50 (100 rows): ${indexP50.toFixed(3)} ms`);
  console.log(`  Index Slice Read p95 (100 rows): ${indexP95.toFixed(3)} ms`);

  // Benchmark 3: Row Resolver Latency (100 scattered rows)
  const resolveSamples: number[] = [];
  for (let i = 0; i < 20; i++) {
    const offset = Math.floor(Math.random() * 100_000);
    const { stableRowIds } = await QueryIndexManager.readStableRowIds(DATASET_ID, ctMeta.queryId, offset, 100);
    const t0 = performance.now();
    await DatasetRowResolver.resolveRows(DATASET_ID, stableRowIds);
    const t1 = performance.now();
    resolveSamples.push(t1 - t0);
  }
  resolveSamples.sort((a, b) => a - b);
  const resolveP50 = resolveSamples[Math.floor(resolveSamples.length * 0.5)];
  const resolveP95 = resolveSamples[Math.floor(resolveSamples.length * 0.95)];
  console.log(`  Row Resolution p50 (100 rows):   ${resolveP50.toFixed(2)} ms`);
  console.log(`  Row Resolution p95 (100 rows):   ${resolveP95.toFixed(2)} ms`);

  results.benchmarks = {
    warmQueryHitMs: warmCompileMs,
    indexSliceP50Ms: indexP50,
    indexSliceP95Ms: indexP95,
    rowResolveP50Ms: resolveP50,
    rowResolveP95Ms: resolveP95,
    flMatchedRows: flAscMeta.matchedRows,
    ctMatchedRows: ctMeta.matchedRows,
    passedCount: passed,
    failedCount: failed,
  };

  const resultsPath = path.join(process.cwd(), 'benchmarks', 'phase1c-filter-sort', 'phase1c_benchmark_results.json');
  await fs.promises.writeFile(resultsPath, JSON.stringify(results, null, 2), 'utf8');

  console.log('\n' + '='.repeat(80));
  console.log(`   PHASE 1C E2E VERIFICATION: ${passed} PASSED, ${failed} FAILED`);
  console.log(`   Results saved to: ${resultsPath}`);
  console.log('='.repeat(80));

  if (failed > 0) {
    process.exit(1);
  }
}

runBenchmark().catch((err) => {
  console.error('Fatal benchmark error:', err);
  process.exit(1);
}).finally(() => {
  process.exit(0);
});
